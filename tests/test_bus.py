import pytest
from conftest import T0, make_attachment, make_message
from eggai import InMemoryTransport

from outlook_connector.bus import EMAIL_RECEIVED, build_bus_event, build_transport
from outlook_connector.config import BusConfig


def test_build_bus_event_envelope():
    message = make_message("m1", has_attachments=True)

    event = build_bus_event(
        message,
        source_mailbox="inbox@example.com",
        fetched_at=T0,
        attachments=[make_attachment()],
    )

    assert event.type == EMAIL_RECEIVED
    assert event.source == "/outlook-connector"
    assert event.data.source_mailbox == "inbox@example.com"
    assert event.data.fetched_at == T0
    assert event.data.email.id == "m1"
    assert len(event.data.email.attachments) == 1


def test_envelope_round_trips_through_json():
    binary = b"%PDF-1.4\xff\x00\x89"  # not valid UTF-8: must ride as base64
    event = build_bus_event(
        make_message("m1", has_attachments=True),
        source_mailbox="inbox@example.com",
        fetched_at=T0,
        attachments=[make_attachment(content=binary)],
    )

    raw = event.model_dump_json()
    parsed = type(event).model_validate_json(raw)

    assert parsed.data.email.internet_message_id == "<m1@example.com>"
    assert parsed.data.email.attachments[0].body == binary


def test_build_transport_inmemory():
    transport = build_transport(BusConfig(transport="inmemory"))
    assert isinstance(transport, InMemoryTransport)


def test_build_transport_rejects_unknown():
    config = BusConfig.model_construct(transport="carrier-pigeon")
    with pytest.raises(ValueError):
        build_transport(config)
