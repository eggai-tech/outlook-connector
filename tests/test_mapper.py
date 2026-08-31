from conftest import T0, make_attachment, make_message

from outlook_connector.mapper import outlook_message_to_email


def test_maps_core_fields():
    email = outlook_message_to_email(make_message("m1"))

    assert email.id == "m1"
    assert email.internet_message_id == "<m1@example.com>"
    assert email.from_addresses == ["sender@example.com"]
    assert email.to_addresses == ["to@example.com"]
    assert email.subject == "subject m1"
    assert email.received_at == T0
    assert email.body_html == "<p>hi</p>"
    assert email.body_text is None


def test_text_body_lands_in_body_text():
    message = make_message("m1", body_content_type="text", body_content="plain")

    email = outlook_message_to_email(message)

    assert email.body_text == "plain"
    assert email.body_html is None


def test_missing_sender_and_body_are_safe():
    message = make_message("m1", sender=None)
    message.body = None

    email = outlook_message_to_email(message)

    assert email.from_addresses == []
    assert email.body_html is None
    assert email.body_text is None


def test_attachments_mapped_and_contentless_skipped():
    message = make_message("m1", has_attachments=True)
    attachments = [
        make_attachment("doc.pdf"),
        make_attachment("linked-item", content=None),  # item/reference attachment
    ]

    email = outlook_message_to_email(message, attachments)

    assert email.has_attachments is True
    assert [a.file_name for a in email.attachments] == ["doc.pdf"]
    assert email.attachments[0].content_type == "application/pdf"
    assert email.attachments[0].body == b"%PDF-1.4"
