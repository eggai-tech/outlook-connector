"""Tests for the boundary mapping: outlook-helper's OutlookMessage -> owned model.

The mapping is the single seam between the dependency's schema and the bus
contract. It must be lossless for the fields the contract carries and defensive
about the helper's many ``Optional`` fields, since the owned ``Email`` model
makes several of them required.

Runnable standalone (`python test_mapping.py`) or under pytest.
"""

from datetime import datetime, timezone

from outlook_helper import EmailAddress as HelperAddress
from outlook_helper import OutlookBody, OutlookMessage
from outlook_helper.schemas import InternetMessageHeader

from mapping import build_event, map_email
from schemas import EMAIL_RECEIVED, EmailReceived

_RECEIVED_AT = datetime(2026, 6, 26, 9, 29, 12, 500000, tzinfo=timezone.utc)


def _full_message() -> OutlookMessage:
    return OutlookMessage(
        id="AAMkAGI2THVSAAA=",
        internet_message_id="<CADnf...@mail.gmail.com>",
        subject="March invoice",
        from_=HelperAddress(name="Alice", address="alice@example.com"),
        to=[HelperAddress(name="Invoices", address="invoices@egg-ai.com")],
        cc=[HelperAddress(address="carol@example.com")],
        received_at=_RECEIVED_AT,
        body=OutlookBody(content_type="html", content="<p>Please find attached.</p>"),
        body_preview="Please find attached.",
        conversation_id="AAQkAGI2-conv",
        has_attachments=True,
        internet_message_headers=[
            InternetMessageHeader(name="In-Reply-To", value="<parent@example.com>"),
            InternetMessageHeader(
                name="References", value="<root@example.com> <parent@example.com>"
            ),
            InternetMessageHeader(name="X-Mailer", value="Outlook"),
        ],
    )


def test_maps_identity():
    email = map_email(_full_message())
    assert email.message_id == "<CADnf...@mail.gmail.com>"
    assert email.graph_id == "AAMkAGI2THVSAAA="


def test_maps_core_headers():
    email = map_email(_full_message())
    assert (email.from_.name, email.from_.address) == ("Alice", "alice@example.com")
    assert [(a.name, a.address) for a in email.to] == [("Invoices", "invoices@egg-ai.com")]
    assert [(a.name, a.address) for a in email.cc] == [(None, "carol@example.com")]
    assert email.subject == "March invoice"
    assert email.received_datetime == _RECEIVED_AT


def test_maps_body_losslessly():
    email = map_email(_full_message())
    assert email.body == "<p>Please find attached.</p>"
    assert email.body_content_type == "html"
    assert email.preview == "Please find attached."


def test_maps_threading_from_headers():
    email = map_email(_full_message())
    assert email.conversation_id == "AAQkAGI2-conv"
    assert email.in_reply_to == "<parent@example.com>"
    assert email.references == ["<root@example.com>", "<parent@example.com>"]


def test_extra_headers_excludes_modeled_threading_headers():
    email = map_email(_full_message())
    # In-Reply-To / References are modeled fields, so they don't double up here.
    assert email.extra_headers == {"X-Mailer": "Outlook"}


def test_attachments_metadata_not_populated_yet():
    email = map_email(_full_message())
    assert email.has_attachments is True  # native flag is free
    assert email.attachments == []  # content/metadata fetch is Piece 4


def test_falls_back_to_graph_id_when_internet_message_id_missing():
    msg = _full_message()
    msg.internet_message_id = None
    email = map_email(msg)
    assert email.message_id == "AAMkAGI2THVSAAA="


def test_defensive_defaults_for_sparse_message():
    msg = OutlookMessage(id="g1", internet_message_id="<m1>", received_at=_RECEIVED_AT)
    email = map_email(msg)
    assert email.from_.address == ""  # no sender -> empty, never crashes
    assert email.subject == ""
    assert email.body == ""
    assert email.body_content_type == "html"  # HTML is requested globally
    assert email.to == [] and email.cc == []


def test_build_event_wraps_payload():
    msg = _full_message()
    fetched = datetime(2026, 6, 26, 9, 30, 0, tzinfo=timezone.utc)
    event = build_event(msg, source_mailbox="invoices@egg-ai.com", fetched_at=fetched)

    assert event.type == EMAIL_RECEIVED
    data = event.data
    assert isinstance(data, EmailReceived)
    assert data.source_mailbox == "invoices@egg-ai.com"
    assert data.fetched_at == fetched
    assert data.email.message_id == "<CADnf...@mail.gmail.com>"


if __name__ == "__main__":
    test_maps_identity()
    test_maps_core_headers()
    test_maps_body_losslessly()
    test_maps_threading_from_headers()
    test_extra_headers_excludes_modeled_threading_headers()
    test_attachments_metadata_not_populated_yet()
    test_falls_back_to_graph_id_when_internet_message_id_missing()
    test_defensive_defaults_for_sparse_message()
    test_build_event_wraps_payload()
    print("All mapping tests passed.")
