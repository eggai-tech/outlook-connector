"""Tests for the inbound poll loop's per-mailbox processing.

These exercise the cursor/failure semantics that implement the spec's
"never duplicate, occasionally drop" principle:

- new mail is published ascending and the cursor advances to the batch max;
- a publish failure mid-batch stops the batch with the cursor at the last
  success (so the next cycle resumes, no duplicates);
- a Graph error leaves the cursor untouched and never blocks other mailboxes.

The poller takes injected ``fetch``/``publish`` callables so these run without a
live Graph API or bus. Runnable standalone (`python -m tests.test_poller`) or pytest.
"""

import asyncio
from datetime import datetime, timedelta, timezone

import httpx
from outlook_helper import EmailAddress as HelperAddress
from outlook_helper import (
    GraphError,
    OutlookAttachment,
    OutlookBody,
    OutlookMessage,
)

from outlook_connector.poller import Poller
from outlook_connector.storage import StorageError, init_backends

_T0 = datetime(2026, 6, 26, 12, 0, 0, tzinfo=timezone.utc)
_NOW = datetime(2026, 6, 26, 12, 5, 0, tzinfo=timezone.utc)


def _msg(
    seconds: int, mid: str | None = None, has_attachments: bool = False
) -> OutlookMessage:
    return OutlookMessage(
        id=f"graph-{seconds}",
        internet_message_id=mid or f"<m{seconds}@example.com>",
        subject=f"msg {seconds}",
        from_=HelperAddress(address="alice@example.com"),
        received_at=_T0 + timedelta(seconds=seconds),
        body=OutlookBody(content_type="html", content="<p>hi</p>"),
        has_attachments=has_attachments,
    )


class _Recorder:
    """Captures published events; can be told to raise on the Nth publish."""

    def __init__(self, fail_on: int | None = None, exc: Exception | None = None):
        self.published = []
        self._fail_on = fail_on
        self._exc = exc or RuntimeError("bus down")

    async def publish(self, message):
        if self._fail_on is not None and len(self.published) + 1 == self._fail_on:
            raise self._exc
        self.published.append(message)


def _poller(*, fetch, publish, cursors, fetch_attachments=None, save=None):
    return Poller(
        mailboxes=list(cursors),
        fetch=fetch,
        publish=publish,
        cursors=cursors,
        now=lambda: _NOW,
        fetch_attachments=fetch_attachments,
        save=save,
    )


def _static_fetch(batches: dict[str, list]):
    async def fetch(mailbox, cursor):
        return list(batches.get(mailbox, []))

    return fetch


# --- happy path ------------------------------------------------------------


def test_publishes_ascending_and_advances_cursor_to_batch_max():
    async def go():
        rec = _Recorder()
        # Provided newest-first (as Graph/helper returns); poller must sort asc.
        batch = [_msg(30), _msg(10), _msg(20)]
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": batch}),
            publish=rec.publish,
            cursors=cursors,
        )
        summaries = await p.run_cycle()

        order = [m.data.email.received_datetime for m in rec.published]
        assert order == [_T0 + timedelta(seconds=s) for s in (10, 20, 30)]
        assert cursors["a@egg-ai.com"] == _T0 + timedelta(seconds=30)
        assert summaries[0].fetched == 3 and summaries[0].published == 3
        assert summaries[0].error is None

    asyncio.run(go())


def test_fetch_called_with_current_cursor_and_stamps_fetched_at():
    async def go():
        seen = {}

        async def fetch(mailbox, cursor):
            seen[mailbox] = cursor
            return [_msg(10)]

        rec = _Recorder()
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(fetch=fetch, publish=rec.publish, cursors=cursors)
        await p.run_cycle()

        assert seen["a@egg-ai.com"] == _T0
        assert rec.published[0].data.fetched_at == _NOW

    asyncio.run(go())


def test_empty_batch_leaves_cursor_unchanged():
    async def go():
        rec = _Recorder()
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": []}),
            publish=rec.publish,
            cursors=cursors,
        )
        summaries = await p.run_cycle()
        assert cursors["a@egg-ai.com"] == _T0
        assert rec.published == []
        assert summaries[0].fetched == 0 and summaries[0].published == 0

    asyncio.run(go())


# --- attachment metadata enrichment (Piece 4) ------------------------------


class _AttachmentFetcher:
    """Records (mailbox, graph_id) calls; returns canned metadata per graph_id."""

    def __init__(self, by_graph_id: dict[str, list] | None = None):
        self.calls = []
        self._by_graph_id = by_graph_id or {}

    async def __call__(self, mailbox, graph_id):
        self.calls.append((mailbox, graph_id))
        return list(self._by_graph_id.get(graph_id, []))


def test_attachment_bearing_message_is_enriched_with_content():
    async def go():
        atts = [
            OutlookAttachment(
                id="att-1",
                name="invoice.pdf",
                content_type="application/pdf",
                size=5,
                content=b"hello",
            ),
            # An item attachment carries no bytes -> content stays None.
            OutlookAttachment(
                id="att-2", name="forwarded.eml", content_type="message/rfc822", size=99
            ),
        ]
        fetcher = _AttachmentFetcher({"graph-10": atts})
        rec = _Recorder()
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": [_msg(10, has_attachments=True)]}),
            publish=rec.publish,
            cursors=cursors,
            fetch_attachments=fetcher,
        )
        await p.run_cycle()

        assert fetcher.calls == [("a@egg-ai.com", "graph-10")]
        attachments = rec.published[0].data.email.attachments
        assert [(a.filename, a.content_type, a.size, a.content) for a in attachments] == [
            ("invoice.pdf", "application/pdf", 5, b"hello"),
            ("forwarded.eml", "message/rfc822", 99, None),
        ]

    asyncio.run(go())


def test_message_without_attachments_makes_no_extra_call():
    async def go():
        fetcher = _AttachmentFetcher()
        rec = _Recorder()
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": [_msg(10, has_attachments=False)]}),
            publish=rec.publish,
            cursors=cursors,
            fetch_attachments=fetcher,
        )
        await p.run_cycle()

        assert fetcher.calls == []  # gated behind the free has_attachments flag
        assert rec.published[0].data.email.attachments == []

    asyncio.run(go())


def test_attachment_fetch_error_stops_batch_with_cursor_at_last_success():
    async def go():
        async def failing_fetcher(mailbox, graph_id):
            raise GraphError(503, "service unavailable")

        rec = _Recorder()
        # msg(10) has no attachments (publishes fine); msg(20) bears attachments
        # and its metadata fetch fails -> batch stops, cursor at msg(10).
        batch = [_msg(10), _msg(20, has_attachments=True), _msg(30)]
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": batch}),
            publish=rec.publish,
            cursors=cursors,
            fetch_attachments=failing_fetcher,
        )
        summaries = await p.run_cycle()

        assert len(rec.published) == 1
        assert cursors["a@egg-ai.com"] == _T0 + timedelta(seconds=10)
        assert summaries[0].error is not None

    asyncio.run(go())


# --- publish failure: stop the batch, cursor at last success ---------------


def test_publish_failure_mid_batch_stops_and_leaves_cursor_at_last_success():
    async def go():
        # Fail on the 2nd publish: msg(10) succeeds, msg(20) fails, msg(30) never sent.
        rec = _Recorder(fail_on=2)
        batch = [_msg(10), _msg(20), _msg(30)]
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": batch}),
            publish=rec.publish,
            cursors=cursors,
        )
        summaries = await p.run_cycle()

        assert len(rec.published) == 1
        assert rec.published[0].data.email.received_datetime == _T0 + timedelta(seconds=10)
        # cursor left at the last successfully-published message
        assert cursors["a@egg-ai.com"] == _T0 + timedelta(seconds=10)
        assert summaries[0].published == 1
        assert summaries[0].error is not None

    asyncio.run(go())


# --- storage: saved before the bus, a failed save never publishes ----------


def test_email_is_saved_before_it_is_published():
    async def go():
        order = []
        rec = _Recorder()

        async def save(email):
            order.append(("save", email.message_id))
            return {"memory": email.message_id}

        async def publish(message):
            order.append(("publish", message.data.email.message_id))
            await rec.publish(message)

        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": [_msg(10)]}),
            publish=publish,
            cursors=cursors,
            save=save,
        )
        await p.run_cycle()

        assert order == [
            ("save", "<m10@example.com>"),
            ("publish", "<m10@example.com>"),
        ]

    asyncio.run(go())


def test_save_failure_never_publishes_and_drops_only_that_email():
    async def go():
        rec = _Recorder()

        async def save(email):
            # msg(20) is the one storage cannot take.
            if email.message_id == "<m20@example.com>":
                raise StorageError("memory", email.message_id, RuntimeError("boom"))
            return {"memory": email.message_id}

        batch = [_msg(10), _msg(20), _msg(30)]
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": batch}),
            publish=rec.publish,
            cursors=cursors,
            save=save,
        )
        summaries = await p.run_cycle()

        # The unsaved message never reaches the bus; the batch continues.
        published = [m.data.email.message_id for m in rec.published]
        assert published == ["<m10@example.com>", "<m30@example.com>"]
        # The cursor still advances past it: dropped, never re-delivered.
        assert cursors["a@egg-ai.com"] == _T0 + timedelta(seconds=30)
        assert summaries[0].published == 2
        assert summaries[0].dropped == 1
        assert summaries[0].error is None  # a drop is not a stopped batch

    asyncio.run(go())


def test_save_log_line_holds_the_email_id_and_no_content(caplog):
    async def go():
        rec = _Recorder()

        async def save(email):
            return {"memory": "stored-1"}

        msg = _msg(10)
        msg.subject = "March invoice"
        msg.body = OutlookBody(content_type="html", content="<p>bank details</p>")
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": [msg]}),
            publish=rec.publish,
            cursors=cursors,
            save=save,
        )
        with caplog.at_level("INFO", logger="outlook_connector.poller"):
            await p.run_cycle()

        saved = [r for r in caplog.records if r.getMessage().startswith("saved ")]
        assert len(saved) == 1
        line = saved[0].getMessage()
        assert "<m10@example.com>" in line  # the email id
        assert "stored-1" in line  # what the backend returned
        assert "March invoice" not in line  # never the subject
        assert "bank details" not in line  # never the body
        assert "mime" not in line  # never the mime

    asyncio.run(go())


def test_storage_is_wired_by_default_and_stores_nothing_when_inactive():
    """No injected ``save`` means the real one, which is a no-op with no backends."""

    async def go():
        init_backends([])
        rec = _Recorder()
        cursors = {"a@egg-ai.com": _T0}
        p = Poller(
            mailboxes=["a@egg-ai.com"],
            fetch=_static_fetch({"a@egg-ai.com": [_msg(10)]}),
            publish=rec.publish,
            cursors=cursors,
            now=lambda: _NOW,
        )
        await p.run_cycle()

        assert len(rec.published) == 1

    asyncio.run(go())


def test_default_save_reaches_the_active_backend():
    async def go():
        (_, backend), = init_backends(["memory"])
        try:
            rec = _Recorder()
            cursors = {"a@egg-ai.com": _T0}
            p = Poller(
                mailboxes=["a@egg-ai.com"],
                fetch=_static_fetch({"a@egg-ai.com": [_msg(10)]}),
                publish=rec.publish,
                cursors=cursors,
                now=lambda: _NOW,
            )
            await p.run_cycle()

            assert list(backend.emails) == ["<m10@example.com>"]
            assert len(rec.published) == 1
        finally:
            init_backends([])

    asyncio.run(go())


# --- Graph errors: cursor untouched, isolation across mailboxes ------------


def test_graph_error_on_fetch_leaves_cursor_untouched():
    async def go():
        async def fetch(mailbox, cursor):
            raise GraphError(503, "service unavailable")

        rec = _Recorder()
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(fetch=fetch, publish=rec.publish, cursors=cursors)
        summaries = await p.run_cycle()

        assert cursors["a@egg-ai.com"] == _T0
        assert rec.published == []
        assert summaries[0].error is not None

    asyncio.run(go())


def test_network_error_on_fetch_is_caught():
    async def go():
        async def fetch(mailbox, cursor):
            raise httpx.ConnectError("no route to host")

        rec = _Recorder()
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(fetch=fetch, publish=rec.publish, cursors=cursors)
        summaries = await p.run_cycle()  # must not raise
        assert summaries[0].error is not None
        assert cursors["a@egg-ai.com"] == _T0

    asyncio.run(go())


def test_one_mailbox_failure_does_not_block_others():
    async def go():
        async def fetch(mailbox, cursor):
            if mailbox == "bad@egg-ai.com":
                raise GraphError(500, "boom")
            return [_msg(10)]

        rec = _Recorder()
        cursors = {"bad@egg-ai.com": _T0, "good@egg-ai.com": _T0}
        p = _poller(fetch=fetch, publish=rec.publish, cursors=cursors)
        summaries = await p.run_cycle()

        # good mailbox still processed and advanced
        assert cursors["good@egg-ai.com"] == _T0 + timedelta(seconds=10)
        assert cursors["bad@egg-ai.com"] == _T0
        assert len(rec.published) == 1
        by_mb = {s.mailbox: s for s in summaries}
        assert by_mb["bad@egg-ai.com"].error is not None
        assert by_mb["good@egg-ai.com"].error is None

    asyncio.run(go())


def test_messages_without_received_at_are_skipped():
    async def go():
        rec = _Recorder()
        no_dt = OutlookMessage(id="g0", internet_message_id="<m0>", received_at=None)
        batch = [no_dt, _msg(10)]
        cursors = {"a@egg-ai.com": _T0}
        p = _poller(
            fetch=_static_fetch({"a@egg-ai.com": batch}),
            publish=rec.publish,
            cursors=cursors,
        )
        await p.run_cycle()
        assert len(rec.published) == 1
        assert cursors["a@egg-ai.com"] == _T0 + timedelta(seconds=10)

    asyncio.run(go())


# --- build_poller wiring ---------------------------------------------------


def _settings(initial_cursor=None):
    from outlook_connector.config import BusConfig, Settings

    return Settings(
        mailboxes=["a@egg-ai.com", "b@egg-ai.com"],
        bus=BusConfig(transport="inmemory"),
        azure_tenant_id="t",
        azure_client_id="c",
        azure_client_secret="s3cret",
        initial_cursor=initial_cursor,
    )


def test_build_poller_seeds_cursors_from_initial_cursor():
    from eggai import InMemoryTransport

    from outlook_connector.service import build_poller

    p = build_poller(_settings(initial_cursor=_T0), InMemoryTransport())
    assert p.mailboxes == ["a@egg-ai.com", "b@egg-ai.com"]
    assert p.cursors == {"a@egg-ai.com": _T0, "b@egg-ai.com": _T0}


def test_build_poller_wires_attachment_fetcher():
    from eggai import InMemoryTransport

    from outlook_connector.service import build_poller

    p = build_poller(_settings(initial_cursor=_T0), InMemoryTransport())
    assert p._fetch_attachments is not None  # enrichment path is wired


def test_build_poller_defaults_cursors_to_now():
    from eggai import InMemoryTransport

    from outlook_connector.service import build_poller

    before = datetime.now(timezone.utc)
    p = build_poller(_settings(), InMemoryTransport())
    after = datetime.now(timezone.utc)
    for cursor in p.cursors.values():
        assert cursor.tzinfo is not None  # timezone-aware
        assert before <= cursor <= after


if __name__ == "__main__":
    test_publishes_ascending_and_advances_cursor_to_batch_max()
    test_attachment_bearing_message_is_enriched_with_content()
    test_message_without_attachments_makes_no_extra_call()
    test_attachment_fetch_error_stops_batch_with_cursor_at_last_success()
    test_fetch_called_with_current_cursor_and_stamps_fetched_at()
    test_empty_batch_leaves_cursor_unchanged()
    test_email_is_saved_before_it_is_published()
    test_save_failure_never_publishes_and_drops_only_that_email()
    # test_save_log_line_holds_the_email_id_and_no_content needs pytest's caplog
    test_storage_is_wired_by_default_and_stores_nothing_when_inactive()
    test_default_save_reaches_the_active_backend()
    test_publish_failure_mid_batch_stops_and_leaves_cursor_at_last_success()
    test_graph_error_on_fetch_leaves_cursor_untouched()
    test_network_error_on_fetch_is_caught()
    test_one_mailbox_failure_does_not_block_others()
    test_messages_without_received_at_are_skipped()
    test_build_poller_seeds_cursors_from_initial_cursor()
    test_build_poller_wires_attachment_fetcher()
    test_build_poller_defaults_cursors_to_now()
    print("All poller tests passed.")
