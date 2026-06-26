"""Tests for the owned email.received bus contract.

Runnable standalone (`python -m tests.test_schemas`) or under pytest.
"""

from datetime import datetime, timezone

import pytest
from pydantic import ValidationError

from schemas import (
    EMAIL_RECEIVED,
    Attachment,
    Email,
    EmailAddress,
    EmailReceived,
    EmailReceivedMessage,
)

_RECEIVED_AT = datetime(2026, 6, 16, 9, 29, 12, tzinfo=timezone.utc)
_FETCHED_AT = datetime(2026, 6, 16, 9, 30, 0, tzinfo=timezone.utc)


def _sample_email() -> Email:
    return Email(
        message_id="<CADnf...@mail.gmail.com>",
        graph_id="AAMkAGI2THVSAAA=",
        # use the wire alias to prove population-by-alias works
        **{"from": EmailAddress(name="Alice", address="alice@example.com")},
        to=[EmailAddress(name="Bob", address="bob@example.com")],
        cc=[EmailAddress(address="carol@example.com")],
        subject="March invoice",
        received_datetime=_RECEIVED_AT,
        extra_headers={"X-Mailer": "Outlook"},
        conversation_id="AAQkAGI2-conv",
        in_reply_to="<parent@example.com>",
        references=["<root@example.com>", "<parent@example.com>"],
        body="<p>Please find attached.</p>",
        body_content_type="html",
        preview="Please find attached.",
        has_attachments=True,
        attachments=[
            Attachment(
                filename="invoice-2026-03.pdf",
                content_type="application/pdf",
                size=12345,
            )
        ],
    )


def _sample_message() -> EmailReceivedMessage:
    payload = EmailReceived(
        source_mailbox="invoices@egg-ai.com",
        fetched_at=_FETCHED_AT,
        email=_sample_email(),
    )
    return EmailReceivedMessage(
        source="/outlook-connector", type=EMAIL_RECEIVED, data=payload
    )


def test_round_trip_preserves_fields():
    msg = _sample_message()
    wire = msg.model_dump_json(by_alias=True)
    restored = EmailReceivedMessage.model_validate_json(wire)

    assert restored.type == "email.received"
    assert restored.data.source_mailbox == "invoices@egg-ai.com"
    assert restored.data.fetched_at == _FETCHED_AT

    email = restored.data.email
    assert email.message_id == "<CADnf...@mail.gmail.com>"
    assert email.graph_id == "AAMkAGI2THVSAAA="
    assert email.from_ == EmailAddress(name="Alice", address="alice@example.com")
    assert email.to == [EmailAddress(name="Bob", address="bob@example.com")]
    assert email.cc == [EmailAddress(address="carol@example.com")]
    assert email.subject == "March invoice"
    assert email.received_datetime == _RECEIVED_AT
    assert email.extra_headers == {"X-Mailer": "Outlook"}
    assert email.conversation_id == "AAQkAGI2-conv"
    assert email.in_reply_to == "<parent@example.com>"
    assert email.references == ["<root@example.com>", "<parent@example.com>"]
    assert email.body == "<p>Please find attached.</p>"
    assert email.body_content_type == "html"
    assert email.preview == "Please find attached."
    assert email.has_attachments is True
    assert len(email.attachments) == 1
    att = email.attachments[0]
    assert att.filename == "invoice-2026-03.pdf"
    assert att.content_type == "application/pdf"
    assert att.size == 12345


def test_from_is_serialized_with_wire_key():
    wire = _sample_message().model_dump(by_alias=True)
    email = wire["data"]["email"]
    # On the wire the sender field is "from", not "from_".
    assert email["from"] == {"name": "Alice", "address": "alice@example.com"}
    assert "from_" not in email


def test_email_accepts_python_field_name():
    # Population by the Python attribute name (from_) must also work.
    email = Email(
        message_id="<x@y>",
        graph_id="g1",
        from_=EmailAddress(address="bob@example.com"),
        subject="hi",
        received_datetime=_RECEIVED_AT,
        body="body",
        body_content_type="text",
    )
    assert email.from_.address == "bob@example.com"


def test_optional_fields_default_to_empty():
    email = Email(
        message_id="<x@y>",
        graph_id="g1",
        from_=EmailAddress(address="bob@example.com"),
        subject="hi",
        received_datetime=_RECEIVED_AT,
        body="body",
        body_content_type="text",
    )
    assert email.to == []
    assert email.cc == []
    assert email.extra_headers == {}
    assert email.conversation_id is None
    assert email.in_reply_to is None
    assert email.references == []
    assert email.preview is None
    assert email.has_attachments is False
    assert email.attachments == []


def test_required_email_fields_are_enforced():
    with pytest.raises(ValidationError):
        # missing graph_id, body, body_content_type, etc.
        Email(message_id="<x@y>")


def test_body_content_type_rejects_unknown_value():
    with pytest.raises(ValidationError):
        Email(
            message_id="<x@y>",
            graph_id="g1",
            from_=EmailAddress(address="bob@example.com"),
            subject="hi",
            received_datetime=_RECEIVED_AT,
            body="body",
            body_content_type="markdown",  # not html|text
        )


def test_email_address_address_is_required():
    EmailAddress(address="a@b.com")  # name optional
    with pytest.raises(ValidationError):
        EmailAddress(name="No Address")


def test_attachment_carries_metadata_only():
    att = Attachment(filename="f.pdf", content_type="application/pdf", size=10)
    # The owned contract never carries attachment content.
    assert not hasattr(att, "content")
    assert "content_base64" not in att.model_dump()


if __name__ == "__main__":
    test_round_trip_preserves_fields()
    test_from_is_serialized_with_wire_key()
    test_email_accepts_python_field_name()
    test_optional_fields_default_to_empty()
    test_required_email_fields_are_enforced()
    test_body_content_type_rejects_unknown_value()
    test_email_address_address_is_required()
    test_attachment_carries_metadata_only()
    print("All tests passed.")
