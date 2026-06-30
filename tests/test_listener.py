"""The outbound send listener consumes ``email.send`` off the real bus.

Mirrors the inbound integration test from the other direction: an ``email.send``
event published over a real eggai ``InMemoryTransport`` reaches the listener's
typed handler exactly once, and the injected ``send`` collaborator receives the
validated request. A co-published ``email.received`` event on the same channel
is ignored — the typed ``data_type`` filter keeps the listener to its contract.

Runnable standalone (`python -m tests.test_listener`) or under pytest.
"""

import asyncio

from eggai import Channel, InMemoryTransport

from listener import make_send_listener
from schemas import (
    EMAIL_RECEIVED,
    EmailAddress,
    EmailSend,
    EmailSendMessage,
)

_SOURCE = "/test-agent"


def _send_event(**overrides) -> EmailSendMessage:
    payload = EmailSend(
        mailbox="invoices@egg-ai.com",
        to=[EmailAddress(address="bob@example.com")],
        subject="hello",
        body="<p>hi</p>",
        **overrides,
    )
    return EmailSendMessage(source=_SOURCE, data=payload)


def test_email_send_event_reaches_the_handler_once():
    async def go():
        transport = InMemoryTransport()
        sent: list[EmailSend] = []

        async def send(request: EmailSend) -> None:
            sent.append(request)

        listener = make_send_listener(
            channel="emails", transport=transport, send=send
        )
        await listener.start()

        publisher = Channel("emails", transport=transport)
        await publisher.publish(_send_event())
        await asyncio.sleep(0.05)  # let the in-memory delivery settle

        assert len(sent) == 1, f"expected exactly one send, got {len(sent)}"
        request = sent[0]
        assert request.mailbox == "invoices@egg-ai.com"
        assert [a.address for a in request.to] == ["bob@example.com"]
        assert request.subject == "hello"
        assert request.reply_to_graph_id is None

        await listener.stop()

    asyncio.run(go())


def test_other_events_on_the_channel_are_ignored():
    async def go():
        transport = InMemoryTransport()
        sent: list[EmailSend] = []

        async def send(request: EmailSend) -> None:
            sent.append(request)

        listener = make_send_listener(
            channel="emails", transport=transport, send=send
        )
        await listener.start()

        # An unrelated event on the same channel must not trigger the handler.
        publisher = Channel("emails", transport=transport)
        await publisher.publish(
            {"type": EMAIL_RECEIVED, "source": _SOURCE, "data": {"anything": 1}}
        )
        await asyncio.sleep(0.05)

        assert sent == [], "listener must ignore non-email.send events"

        await listener.stop()

    asyncio.run(go())


def test_reply_request_carries_the_graph_id():
    async def go():
        transport = InMemoryTransport()
        sent: list[EmailSend] = []

        async def send(request: EmailSend) -> None:
            sent.append(request)

        listener = make_send_listener(
            channel="emails", transport=transport, send=send
        )
        await listener.start()

        publisher = Channel("emails", transport=transport)
        await publisher.publish(_send_event(reply_to_graph_id="graph-42"))
        await asyncio.sleep(0.05)

        assert len(sent) == 1
        assert sent[0].reply_to_graph_id == "graph-42"

        await listener.stop()

    asyncio.run(go())


if __name__ == "__main__":
    test_email_send_event_reaches_the_handler_once()
    test_other_events_on_the_channel_are_ignored()
    test_reply_request_carries_the_graph_id()
    print("Listener test passed.")
