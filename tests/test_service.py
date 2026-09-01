import asyncio
import datetime

import httpx
from conftest import T0, FakeChannel, FakeClient, make_attachment, make_message

from outlook_connector.bus import EMAIL_RECEIVED
from outlook_connector.poller import Poller
from outlook_connector.service import run_workflow


def _at(seconds: int) -> datetime.datetime:
    return T0 + datetime.timedelta(seconds=seconds)


def _context(client, channel):
    poller = Poller(client=client, now=lambda: _at(100))
    return {
        "poller": poller,
        "channel": channel,
        "source_mailbox": "inbox@example.com",
    }


def test_publishes_envelopes_and_marks_published():
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
    assert context["poller"]._published_ids == {"m1", "m2"}
    # rescan with everything marked: the next cycle is quiet
    assert asyncio.run(run_workflow(context)).published == 0


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
    assert summary.error_source == "bus"
    assert [e.data.email.id for e in channel.published] == ["m1"]
    # only the published message is marked: m2 and m3 retry on the next rescan
    assert context["poller"]._published_ids == {"m1"}


def test_graph_error_publishes_nothing_and_marks_nothing():
    class FailingClient(FakeClient):
        def search_email(self, **kwargs):
            raise httpx.ConnectError("boom")

    channel = FakeChannel()
    context = _context(FailingClient(), channel)

    summary = asyncio.run(run_workflow(context))

    assert summary.fetched == 0
    assert "ConnectError" in summary.error
    assert summary.error_source == "graph"
    assert channel.published == []
    assert context["poller"]._published_ids == set()
