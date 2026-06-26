"""The inbound poll loop's per-cycle, per-mailbox processing.

Implements `the spec <docs/implementation.md#per-cycle-per-mailbox-processing>`.
:meth:`Poller.run_cycle` is the unit of work the :class:`~service.Service`
fixed-delay loop calls each cycle; it polls every configured mailbox
**sequentially**, each inside its own ``try/except`` so one mailbox's failure
never crashes the loop or blocks the others.

Cursor / failure semantics (the "never duplicate, occasionally drop" principle):

- Each mailbox has its own in-memory ``receivedDateTime`` cursor. A poll fetches
  ``receivedDateTime > cursor`` (strict ``>``, served natively by the helper's
  ``since_exclusive``).
- The batch is processed **ascending**; the cursor advances to each message's
  ``receivedDateTime`` only **after** a successful publish.
- On the **first publish failure the batch stops**, leaving the cursor at the
  last success so the next cycle resumes there without duplicating.
- On any **Graph/transport error** the cursor is left untouched and the next
  mailbox is polled; the fixed-delay loop retries next cycle.

Attachment metadata is enriched inline: for each message whose free
``has_attachments`` flag is set, the poller makes one extra Graph call via the
injected ``fetch_attachments`` collaborator and the result is mapped onto the
event. Messages without attachments make **no** extra call. That call runs
inside the same per-message ``try`` as publish, so a Graph error fetching
metadata stops the batch with the cursor at the last success — the "never
duplicate" path, retried next cycle.

The ``fetch``, ``publish`` and ``fetch_attachments`` collaborators are injected
so the cursor logic is testable without a live Graph API or bus;
:func:`service.build_poller` wires the real ones.
"""

import logging
from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from datetime import datetime, timezone

import httpx
from outlook_helper import GraphError, OutlookAttachmentMeta, OutlookMessage

from mapping import build_event

logger = logging.getLogger(__name__)

# Fetch new messages for a mailbox given its current cursor (async-wrapped Graph
# call). Publish a built bus event. Both injected for testability.
Fetch = Callable[[str, datetime], Awaitable[list[OutlookMessage]]]
Publish = Callable[[object], Awaitable[None]]
# Fetch per-attachment metadata for one message (mailbox, graph_id) -> the extra
# async-wrapped Graph call, made only for attachment-bearing mail. Injected too.
FetchAttachments = Callable[[str, str], Awaitable[list[OutlookAttachmentMeta]]]

# Errors that mean "the Graph call failed" — leave the cursor, try next cycle.
# The helper already retries 429/503 honoring Retry-After; everything else
# (other 5xx, network failures) surfaces as one of these.
_GRAPH_ERRORS = (GraphError, httpx.HTTPError)


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


@dataclass
class MailboxSummary:
    """Per-cycle, per-mailbox observability record (also emitted to the log)."""

    mailbox: str
    fetched: int = 0
    published: int = 0
    error: str | None = None


class Poller:
    """Polls mailboxes sequentially and bridges new mail to the bus."""

    def __init__(
        self,
        *,
        mailboxes: list[str],
        fetch: Fetch,
        publish: Publish,
        cursors: dict[str, datetime],
        now: Callable[[], datetime] = _utcnow,
        fetch_attachments: FetchAttachments | None = None,
    ):
        self.mailboxes = mailboxes
        self._fetch = fetch
        self._publish = publish
        self.cursors = cursors
        self._now = now
        self._fetch_attachments = fetch_attachments

    async def run_cycle(self) -> list[MailboxSummary]:
        """Poll every mailbox once, in order, isolating per-mailbox failures."""
        summaries = []
        for mailbox in self.mailboxes:
            summaries.append(await self._poll_mailbox(mailbox))
        return summaries

    async def _poll_mailbox(self, mailbox: str) -> MailboxSummary:
        summary = MailboxSummary(mailbox)
        cursor = self.cursors[mailbox]
        fetched_at = self._now()

        try:
            messages = await self._fetch(mailbox, cursor)
        except _GRAPH_ERRORS as exc:
            summary.error = f"{type(exc).__name__}: {exc}"
            logger.warning("poll %s: fetch failed: %s", mailbox, summary.error)
            return summary

        # Drop any without a timestamp (can't be ordered or cursored), then sort
        # ascending so the cursor advances monotonically and we never duplicate.
        messages = [m for m in messages if m.received_at is not None]
        messages.sort(key=lambda m: m.received_at)
        summary.fetched = len(messages)

        for message in messages:
            try:
                attachments = await self._attachment_metadata(mailbox, message)
                event = build_event(
                    message,
                    source_mailbox=mailbox,
                    fetched_at=fetched_at,
                    attachments=attachments,
                )
                await self._publish(event)
            except Exception as exc:  # publish/map failure -> stop the batch
                summary.error = f"{type(exc).__name__}: {exc}"
                logger.warning(
                    "poll %s: stopped batch at publish failure: %s",
                    mailbox,
                    summary.error,
                )
                break
            self.cursors[mailbox] = message.received_at
            summary.published += 1

        logger.info(
            "poll %s: fetched=%d published=%d error=%s",
            mailbox,
            summary.fetched,
            summary.published,
            summary.error,
        )
        return summary

    async def _attachment_metadata(
        self, mailbox: str, message: OutlookMessage
    ) -> list[OutlookAttachmentMeta]:
        """Fetch per-attachment metadata, but only for attachment-bearing mail.

        Gated on the free native ``has_attachments`` flag so messages without
        attachments never incur the extra Graph call. Any Graph/transport error
        propagates to the caller's ``try``, which stops the batch.
        """
        if not message.has_attachments or self._fetch_attachments is None:
            return []
        return await self._fetch_attachments(mailbox, message.id)
