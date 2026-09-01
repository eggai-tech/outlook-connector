import datetime
from collections.abc import Callable
from typing import Literal

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


class PollSummary(BaseModel):
    """Per-cycle observability record (also emitted to the log)."""

    fetched: int = 0
    published: int = 0
    # emails skipped because publishing (or attachment fetch) failed mid-batch
    dropped: int = 0
    # full error text — logs only: Graph errors embed mailbox addresses and
    # URLs, so this never leaves the process via the health endpoint
    error: str | None = None
    # the identity-free form the unauthenticated health endpoint may expose
    error_class: str | None = None
    error_status: int | None = None  # Graph HTTP status, when there is one
    # which dependency the error came from: the Graph API or the bus publish
    error_source: Literal["graph", "bus"] | None = None


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

    def poll_mailbox(self) -> list[OutlookMessage]:
        """Fetch messages received strictly after the cursor, oldest first,
        at most ``batch_max_messages`` per cycle.

        The Graph query orders ascending so the bound takes the *oldest* N —
        with max-seen cursor advancement, taking the newest N would skip the
        backlog behind them permanently.

        Graph/transport errors propagate to the caller, which owns the
        error policy (log, leave the cursor, retry next cycle).
        """
        messages = list(self._fetch_message(self.cursor))
        messages.sort(key=lambda m: (m.received_at is not None, m.received_at))
        return messages

    def advance(self, received_at: datetime.datetime | None) -> None:
        """Move the cursor forward to a successfully published message.

        Only ever advances (max-seen), so an out-of-order timestamp can never
        pull the cursor back and cause a re-publish.
        """
        if received_at is not None and received_at > self.cursor:
            self.cursor = received_at

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

    def _fetch_message(self, cursor):
        return self.client.search_email(
            folder=self.source_folder,
            since_exclusive=cursor,  # strict >, sub-second precise
            top=self.batch_max_messages,
            oldest_first=True,  # bound takes the oldest N (see poll_mailbox)
            include_headers=True,  # internet_message_id + In-Reply-To/References
            html_body=True,  # body guaranteed HTML
        )
