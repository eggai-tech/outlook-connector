import datetime
from collections import OrderedDict
from collections.abc import Callable

import httpx
import structlog
from outlook_helper import (
    AppOnlyConfig,
    ClientSecretCredential,
    GraphError,
    OutlookAttachment,
    OutlookClient,
    OutlookMessage,
)
from pydantic import BaseModel

from outlook_connector.config import get_settings

# Errors that mean "the Graph call failed" — leave the cursor, try next cycle.
# The helper already retries 429/503 honoring Retry-After; everything else
# (other 5xx, network failures) surfaces as one of these.
GRAPH_ERRORS = (GraphError, httpx.HTTPError)

logger = structlog.getLogger()

# How long a boundary second must be in the past before the cursor may be
# bumped over it. Mail is stamped (receivedDateTime) at transport receipt but
# can become visible to listings tens of seconds later; bumping before the lag
# has safely elapsed would exclude such late-visible mail forever.
_BUMP_VISIBILITY_MARGIN = datetime.timedelta(seconds=120)


class PollSummary(BaseModel):
    """Per-cycle observability record (also emitted to the log)."""

    fetched: int = 0
    published: int = 0
    # emails skipped because publishing (or attachment fetch) failed mid-batch
    dropped: int = 0
    error: str | None = None


def _utcnow() -> datetime.datetime:
    return datetime.datetime.now(datetime.UTC)


def _build_client():
    settings = get_settings()
    credential = ClientSecretCredential(
        AppOnlyConfig(
            client_id=settings.azure_client_id,
            tenant_id=settings.azure_tenant_id,
            client_secret=settings.azure_client_secret,
        )
    )
    return OutlookClient(credential, settings.mailbox)


class Poller:
    """
    Mailbox poller.
    Keeps its own last polled state.

    All methods call the synchronous outlook-helper client; the service wraps
    them in ``asyncio.to_thread`` so a blocking request or ``Retry-After``
    sleep never stalls the event loop.
    """

    def __init__(
        self,
        *,
        client: OutlookClient | None = None,
        cursor: datetime.datetime | None = None,
        now: Callable[[], datetime.datetime] = _utcnow,
        source_folder: str = "inbox",
        batch_max_messages: int | None = None,
        max_attachment_bytes: int | None = None,
    ):
        self.client = (
            _build_client() if client is None else client
        )  # can inject client for testing
        self.cursor = cursor if cursor is not None else now()
        self.now = now
        self.source_folder = source_folder
        self.batch_max_messages = batch_max_messages
        self.max_attachment_bytes = max_attachment_bytes
        # Graph ids already published, with their receivedDateTime. Graph
        # returns receivedDateTime truncated to whole seconds but compares
        # filters against its finer stored value, so `gt <truncated cursor>`
        # keeps matching the boundary message — the newest mail would be
        # re-fetched and re-published every cycle forever (observed live).
        # Entries are evicted once the cursor has moved safely past their
        # second (they can never be returned again), so the dict is bounded by
        # the size of one boundary-second cohort — never by a magic number
        # that a large same-second bulk delivery could overflow.
        self._published_ids: OrderedDict[str, datetime.datetime | None] = OrderedDict()

    def poll_mailbox(self) -> list[OutlookMessage]:
        """Fetch messages received strictly after the cursor, oldest first,
        at most ``batch_max_messages`` per cycle.

        The Graph query orders ascending so the bound takes the *oldest* N —
        with max-seen cursor advancement, taking the newest N would skip the
        backlog behind them permanently.

        Graph/transport errors propagate to the caller, which owns the
        error policy (log, leave the cursor, retry next cycle).
        """
        raw = list(self._fetch_message(self.cursor, self.batch_max_messages))
        messages = [m for m in raw if m.id not in self._published_ids]
        truncated = (
            self.batch_max_messages is not None and len(raw) >= self.batch_max_messages
        )
        if raw and not messages and truncated:
            # The bounded window held only already-published mail: a boundary
            # cohort larger than the batch. The bound would starve everything
            # behind the cohort (Graph keeps returning the same first N), so
            # take one unbounded look at the window.
            raw = list(self._fetch_message(self.cursor, None))
            messages = [m for m in raw if m.id not in self._published_ids]
            truncated = False
        if raw and not messages:
            # The whole window is already published: we are stuck on the
            # truncated-timestamp boundary (see _published_ids). Step the
            # cursor past that second — but only once the boundary is
            # comfortably in the past: mail can be stamped in a second and
            # become visible to listings tens of seconds later, and bumping
            # early would exclude it forever. Until then the id filter keeps
            # cycles quiet at the cost of re-fetching the boundary window.
            stamps = [m.received_at for m in raw if m.received_at is not None]
            if stamps:
                boundary = max(stamps)
                bumped = boundary + datetime.timedelta(seconds=1)
                aged = self.now() - boundary >= _BUMP_VISIBILITY_MARGIN
                if aged and bumped > self.cursor:
                    logger.debug("Cursor bumped past truncated boundary", cursor=bumped)
                    self.cursor = bumped
        elif len(messages) < len(raw):
            logger.debug(
                "Skipping already-published messages", skipped=len(raw) - len(messages)
            )
        messages.sort(key=lambda m: (m.received_at is not None, m.received_at))
        if self.batch_max_messages is not None:
            messages = messages[: self.batch_max_messages]
        return messages

    def advance(self, message: OutlookMessage) -> None:
        """Record a successfully published message and move the cursor to it.

        The cursor only ever advances (max-seen), so an out-of-order timestamp
        can never pull it back; the id is remembered so the truncated-timestamp
        boundary message is not re-published next cycle (see poll_mailbox).
        """
        self._published_ids[message.id] = message.received_at
        received_at = message.received_at
        if received_at is not None and received_at > self.cursor:
            self.cursor = received_at
        # Evict ids the cursor has safely passed: `gt cursor` can never return
        # a message whose (truncated) receivedDateTime is a full second behind
        # it, so remembering it buys nothing.
        horizon = self.cursor - datetime.timedelta(seconds=1)
        for message_id, stamp in list(self._published_ids.items()):
            if stamp is not None and stamp < horizon:
                del self._published_ids[message_id]

    def fetch_attachments(self, message: OutlookMessage) -> list[OutlookAttachment]:
        """Fetch per-attachment metadata + content, only for attachment-bearing mail.

        Gated on the free native ``has_attachments`` flag so messages without
        attachments never incur the extra Graph call. Content larger than
        ``max_attachment_bytes`` is stripped to metadata (``content=None``).
        Any Graph/transport error propagates to the caller, which stops the
        batch.
        """
        if not message.has_attachments:
            return []
        return [self._apply_size_cap(a, message) for a in self.client.get_attachments(message.id)]

    def _apply_size_cap(
        self, attachment: OutlookAttachment, message: OutlookMessage
    ) -> OutlookAttachment:
        cap = self.max_attachment_bytes
        if cap is None or attachment.content is None or len(attachment.content) <= cap:
            return attachment
        logger.warning(
            "Attachment over size cap; publishing metadata only",
            message=message.id,
            attachment=attachment.name,
            size=len(attachment.content),
            cap=cap,
        )
        return attachment.model_copy(update={"content": None})

    def _fetch_message(self, cursor, top):
        return self.client.search_email(
            folder=self.source_folder,
            since_exclusive=cursor,  # strict >, sub-second precise
            top=top,
            oldest_first=True,  # bound takes the oldest N (see poll_mailbox)
            include_headers=True,  # internet_message_id + In-Reply-To/References
            html_body=True,  # body guaranteed HTML
        )
