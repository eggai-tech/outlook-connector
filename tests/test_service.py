import asyncio
import datetime

import httpx
from conftest import T0, FakeChannel, FakeClient, make_attachment, make_message

from outlook_connector.bus import EMAIL_RECEIVED
from outlook_connector.poller import Poller
from outlook_connector.service import run_workflow


def _at(seconds: int) -> datetime.datetime:
    return T0 + datetime.timedelta(seconds=seconds)


def _context(client, channel, cursor=T0):
    poller = Poller(client=client, cursor=cursor, now=lambda: _at(100))
    return {
        "poller": poller,
        "channel": channel,
        "source_mailbox": "inbox@example.com",
    }


def test_publishes_envelopes_and_advances_cursor():
    client = FakeClient(
        messages=[
            make_message("m1", received_at=_at(10)),
            make_message("m2", received_at=_at(20), has_attachments=True),
        ],
        attachments=[make_attachment()],
    )
    channel = FakeChannel()
    context = _context(client, channel)

    summary = asyncio.run(run_workflow(context))

    assert (summary.fetched, summary.published, summary.dropped) == (2, 2, 0)
    assert summary.error is None
    assert [e.data.email.id for e in channel.published] == ["m1", "m2"]
    assert all(e.type == EMAIL_RECEIVED for e in channel.published)
    assert channel.published[1].data.email.attachments[0].file_name == "doc.pdf"
    assert all(e.data.fetched_at == _at(100) for e in channel.published)
    assert context["poller"].cursor == _at(20)


def test_publish_failure_stops_batch_at_last_success():
    client = FakeClient(
        messages=[
            make_message("m1", received_at=_at(10)),
            make_message("m2", received_at=_at(20)),
            make_message("m3", received_at=_at(30)),
        ]
    )
    channel = FakeChannel(fail_on={"m2"})
    context = _context(client, channel)

    summary = asyncio.run(run_workflow(context))

    assert (summary.fetched, summary.published, summary.dropped) == (3, 1, 2)
    assert "bus down" in summary.error
    assert [e.data.email.id for e in channel.published] == ["m1"]
    # cursor sits at the last published message: m2 and m3 retry next cycle
    assert context["poller"].cursor == _at(10)


def test_graph_error_leaves_cursor_untouched():
    class FailingClient(FakeClient):
        def search_email(self, **kwargs):
            raise httpx.ConnectError("boom")

    channel = FakeChannel()
    context = _context(FailingClient(), channel)

    summary = asyncio.run(run_workflow(context))

    assert summary.fetched == 0
    assert "ConnectError" in summary.error
    assert channel.published == []
    assert context["poller"].cursor == T0
