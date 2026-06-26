"""End-to-end check: a fetched message reaches the real bus exactly once.

Exercises the full inbound path — :func:`service.build_poller` -> the real
``OutlookClient`` wiring (with the Graph fetch stubbed out) -> mapping ->
publish over a real eggai ``InMemoryTransport`` -> a subscriber. This is the
"Done when" guarantee from Piece 3: a real email published to the bus exactly
once.

Runnable standalone (`python -m tests.test_integration`) or under pytest.
"""

import asyncio
from datetime import datetime, timezone

from eggai import Channel, InMemoryTransport
from outlook_helper import EmailAddress as HelperAddress
from outlook_helper import OutlookAttachmentMeta, OutlookBody, OutlookMessage

from config import AzureConfig, BusConfig, Settings
from schemas import EMAIL_RECEIVED
from service import build_poller

_T0 = datetime(2026, 6, 26, 12, 0, 0, tzinfo=timezone.utc)


def _settings():
    return Settings(
        mailboxes=["invoices@egg-ai.com"],
        bus=BusConfig(transport="inmemory", channel="emails"),
        azure=AzureConfig(tenant_id="t", client_id="c"),
        client_secret="s3cret",
        initial_cursor=_T0,
    )


def test_fetched_message_published_to_bus_exactly_once():
    async def go():
        transport = InMemoryTransport()
        received = []

        async def on_event(message):
            received.append(message)

        # Subscribe through a Channel so the namespace-prefixed topic matches
        # what the connector's publishing Channel uses.
        subscriber = Channel("emails", transport=transport)
        await subscriber.subscribe(on_event, group_id="test-consumer")
        await transport.connect()

        poller = build_poller(_settings(), transport)
        # Stub the Graph fetch: one new message after the cursor.
        msg = OutlookMessage(
            id="graph-1",
            internet_message_id="<m1@example.com>",
            subject="hello",
            from_=HelperAddress(address="alice@example.com"),
            received_at=datetime(2026, 6, 26, 12, 0, 30, tzinfo=timezone.utc),
            body=OutlookBody(content_type="html", content="<p>hi</p>"),
        )
        poller._fetch = _once({"invoices@egg-ai.com": [msg]})

        await poller.run_cycle()
        await poller.run_cycle()  # second cycle: cursor advanced, nothing new
        await asyncio.sleep(0.05)  # let the in-memory delivery settle

        assert len(received) == 1, f"expected exactly one event, got {len(received)}"
        event = received[0]
        # Delivered as the CloudEvents envelope dict.
        assert event["type"] == EMAIL_RECEIVED
        assert event["data"]["source_mailbox"] == "invoices@egg-ai.com"
        assert event["data"]["email"]["message_id"] == "<m1@example.com>"

        await transport.disconnect()

    asyncio.run(go())


def test_attachment_bearing_message_published_with_metadata():
    async def go():
        transport = InMemoryTransport()
        received = []

        async def on_event(message):
            received.append(message)

        subscriber = Channel("emails", transport=transport)
        await subscriber.subscribe(on_event, group_id="test-consumer")
        await transport.connect()

        poller = build_poller(_settings(), transport)
        msg = OutlookMessage(
            id="graph-att",
            internet_message_id="<att@example.com>",
            subject="with attachment",
            from_=HelperAddress(address="alice@example.com"),
            received_at=datetime(2026, 6, 26, 12, 0, 30, tzinfo=timezone.utc),
            body=OutlookBody(content_type="html", content="<p>see attached</p>"),
            has_attachments=True,
        )
        poller._fetch = _once({"invoices@egg-ai.com": [msg]})
        # Stub the extra Graph call the helper would make for attachment metadata.
        attachment_calls = []

        async def fetch_attachments(mailbox, graph_id):
            attachment_calls.append((mailbox, graph_id))
            return [
                OutlookAttachmentMeta(
                    id="a-1", name="invoice.pdf", content_type="application/pdf", size=2048
                )
            ]

        poller._fetch_attachments = fetch_attachments

        await poller.run_cycle()
        await asyncio.sleep(0.05)

        assert attachment_calls == [("invoices@egg-ai.com", "graph-att")]
        assert len(received) == 1
        attachments = received[0]["data"]["email"]["attachments"]
        assert attachments == [
            {"filename": "invoice.pdf", "content_type": "application/pdf", "size": 2048}
        ]

        await transport.disconnect()

    asyncio.run(go())


def _once(batches):
    """Fetch that returns the batch on the first call per mailbox, then nothing."""
    served = set()

    async def fetch(mailbox, cursor):
        if mailbox in served:
            return []
        served.add(mailbox)
        return list(batches.get(mailbox, []))

    return fetch


if __name__ == "__main__":
    test_fetched_message_published_to_bus_exactly_once()
    test_attachment_bearing_message_published_with_metadata()
    print("Integration test passed.")
