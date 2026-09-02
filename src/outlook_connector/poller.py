import datetime
import time
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
from outlook_helper.http import GraphSession
from pydantic import BaseModel

from outlook_connector.config import get_settings

# Errors that mean "the Graph call failed" — try again next cycle. The helper
# already retries 429/503 honoring Retry-After; everything else (other 5xx,
# network failures) surfaces as one of these.
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


def _noop() -> None:
    return None


def _beating_sleep(seconds: float, heartbeat: Callable[[], None]) -> None:
    """Sleep in chunks, beating between them.

    The helper honors ``Retry-After`` with an in-thread sleep; without beats a
    long backoff would read as a wedged poller and trip the staleness 503.
    """
    deadline = time.monotonic() + seconds
    while True:
        heartbeat()
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return
        time.sleep(min(remaining, 10.0))


def _build_client(heartbeat: Callable[[], None] = _noop):
    settings = get_settings()
    credential = ClientSecretCredential(
        AppOnlyConfig(
            client_id=settings.azure_client_id,
            tenant_id=settings.azure_tenant_id,
            client_secret=settings.azure_client_secret,
        )
    )
    session = GraphSession(credential, sleep=lambda s: _beating_sleep(s, heartbeat))
    return OutlookClient(credential, settings.mailbox, session=session)


class Poller:
    """Folder-rescan poller: the source folder itself is the work set.

    Every cycle lists the *whole* folder (ids only — cheap) and fetches the
    oldest unseen messages in full. There is no cursor and no durable state:
    a bounded in-memory set of already-published ids suppresses re-fetching
    within a process lifetime, and is pruned to the ids still present in the
    folder, so it can never grow past the folder size. A restart empties the
    set and everything still in the folder is published again — delivery is
    **at least once**, and consumers must be idempotent (dedupe on
    ``internet_message_id``).

    All methods call the synchronous outlook-helper client; the service wraps
    them in ``asyncio.to_thread`` so a blocking request or ``Retry-After``
    sleep never stalls the event loop. The poller owns the health heartbeat
    precisely because the work happens in threads: it beats per listed and
    per fetched message and through retry sleeps, so a long rescan never
    reads as a wedged loop.
    """

    def __init__(
        self,
        *,
        client: OutlookClient | None = None,
        now: Callable[[], datetime.datetime] = _utcnow,
        source_folder: str = "inbox",
        batch_max_messages: int | None = None,
        max_attachment_bytes: int | None = None,
        include_mime_content: bool = False,
        ignore_received_before: datetime.datetime | None = None,
        heartbeat: Callable[[], None] = _noop,
    ):
        self.heartbeat = heartbeat
        self.client = (
            _build_client(heartbeat) if client is None else client
        )  # can inject client for testing
        self.now = now
        self.source_folder = source_folder
        self.batch_max_messages = batch_max_messages
        self.max_attachment_bytes = max_attachment_bytes
        self.include_mime_content = include_mime_content
        self.ignore_received_before = ignore_received_before
        # Ids published this process lifetime that are still in the folder.
        self._published_ids: set[str] = set()

    def poll_mailbox(self) -> list[OutlookMessage]:
        """One rescan: list the folder, fetch the oldest unseen batch in full.

        Graph/transport errors propagate to the caller, which owns the error
        policy (log, retry next cycle — nothing is lost, the folder still
        holds the mail).
        """
        listed: list[OutlookMessage] = []
        for stub in self.client.search_email(
            folder=self.source_folder,
            since=self.ignore_received_before,
            oldest_first=True,
            ids_only=True,
        ):
            listed.append(stub)
            self.heartbeat()  # per listed message: pagination counts as progress

        # The listing is the truth: mail moved out of the folder needs no
        # memory (and if a human moves it back, it re-publishes and the
        # consumer dedupes). This also bounds the set by the folder size.
        self._published_ids &= {m.id for m in listed}

        unseen = [m for m in listed if m.id not in self._published_ids]
        unseen.sort(key=lambda m: (m.received_at is not None, m.received_at))
        if self.batch_max_messages is not None:
            unseen = unseen[: self.batch_max_messages]

        messages = []
        for stub in unseen:
            try:
                messages.append(
                    self.client.get_email(
                        stub.id,
                        include_headers=True,
                        html_body=True,
                        include_mime=self.include_mime_content,
                    )
                )
            except GraphError as exc:
                if exc.status_code == 404:
                    # Moved/deleted between the listing and the fetch — the
                    # downstream mover does exactly this concurrently. The
                    # next listing simply won't contain it.
                    logger.debug("Listed message gone before fetch", message=stub.id)
                    continue
                raise
            self.heartbeat()
        return messages

    def mark_published(self, message: OutlookMessage) -> None:
        """Record a successfully published message so rescans skip it."""
        self._published_ids.add(message.id)

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
