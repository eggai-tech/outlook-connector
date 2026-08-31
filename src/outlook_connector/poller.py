import datetime
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
    ):
        self.client = (
            _build_client() if client is None else client
        )  # can inject client for testing
        self.cursor = cursor if cursor is not None else now()
        self.now = now

    def poll_mailbox(self) -> list[OutlookMessage]:
        """Fetch messages received strictly after the cursor, oldest first.

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
        attachments never incur the extra Graph call. Any Graph/transport error
        propagates to the caller, which stops the batch.
        """
        if not message.has_attachments:
            return []
        return self.client.get_attachments(message.id)

    def _fetch_message(self, cursor):
        return self.client.search_email(
            folder="inbox",  # received mail only — exclude Sent Items
            since_exclusive=cursor,  # strict >, sub-second precise
            include_headers=True,  # internet_message_id + In-Reply-To/References
            html_body=True,  # body guaranteed HTML
        )
