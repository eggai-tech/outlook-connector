from datetime import datetime

from outlook_helper.schemas import OutlookAttachmentMeta, OutlookFolder, OutlookMessage


def test_message_parses_full_graph_payload(load_fixture):
    msg = OutlookMessage.model_validate(load_fixture("message.json"))
    assert msg.id == "AAMkAGI1"
    assert msg.subject == "Quarterly report"
    assert msg.is_read is False
    assert msg.has_attachments is True
    assert msg.importance == "high"
    assert msg.web_link.endswith("AAMkAGI1")
    assert msg.conversation_id == "CONV123"
    assert msg.parent_folder_id == "FOLDER_INBOX"


def test_message_flattens_sender_and_recipients(load_fixture):
    msg = OutlookMessage.model_validate(load_fixture("message.json"))
    assert msg.from_.name == "Alice Sender"
    assert msg.from_.address == "alice@example.com"
    assert [r.address for r in msg.to] == ["bob@example.com", "carol@example.com"]
    assert [r.address for r in msg.cc] == ["dan@example.com"]
    assert msg.bcc == []


def test_message_parses_dates_and_body(load_fixture):
    msg = OutlookMessage.model_validate(load_fixture("message.json"))
    assert isinstance(msg.received_at, datetime)
    assert msg.received_at.year == 2026
    assert msg.body.content_type == "html"
    assert "report" in msg.body.content


def test_message_minimal_payload_uses_defaults():
    msg = OutlookMessage.model_validate({"id": "X"})
    assert msg.id == "X"
    assert msg.subject is None
    assert msg.from_ is None
    assert msg.to == []
    assert msg.has_attachments is False


def test_message_parses_internet_message_id_and_headers():
    msg = OutlookMessage.model_validate(
        {
            "id": "X",
            "internetMessageId": "<abc@mail.example>",
            "internetMessageHeaders": [
                {"name": "In-Reply-To", "value": "<parent@mail.example>"},
                {"name": "References", "value": "<root@mail.example> <parent@mail.example>"},
            ],
        }
    )
    assert msg.internet_message_id == "<abc@mail.example>"
    assert len(msg.internet_message_headers) == 2
    assert msg.internet_message_headers[0].name == "In-Reply-To"
    assert msg.internet_message_headers[0].value == "<parent@mail.example>"


def test_message_in_reply_to_is_case_insensitive():
    msg = OutlookMessage.model_validate(
        {
            "id": "X",
            "internetMessageHeaders": [{"name": "in-reply-to", "value": "<p@mail.example>"}],
        }
    )
    assert msg.in_reply_to == "<p@mail.example>"


def test_message_references_splits_on_whitespace():
    msg = OutlookMessage.model_validate(
        {
            "id": "X",
            "internetMessageHeaders": [
                {"name": "References", "value": "<a@mail.example> <b@mail.example>\n<c@mail.example>"},
            ],
        }
    )
    assert msg.references == ["<a@mail.example>", "<b@mail.example>", "<c@mail.example>"]


def test_message_threading_defaults_when_absent():
    msg = OutlookMessage.model_validate({"id": "X"})
    assert msg.internet_message_id is None
    assert msg.internet_message_headers == []
    assert msg.in_reply_to is None
    assert msg.references == []


def test_attachment_meta_parses():
    att = OutlookAttachmentMeta.model_validate(
        {
            "id": "att1",
            "name": "report.pdf",
            "contentType": "application/pdf",
            "size": 20480,
            "isInline": False,
        }
    )
    assert att.id == "att1"
    assert att.name == "report.pdf"
    assert att.content_type == "application/pdf"
    assert att.size == 20480
    assert att.is_inline is False


def test_folder_parses():
    folder = OutlookFolder.model_validate(
        {
            "id": "F1",
            "displayName": "Projects",
            "parentFolderId": "ROOT",
            "childFolderCount": 2,
            "totalItemCount": 10,
            "unreadItemCount": 3,
        }
    )
    assert folder.id == "F1"
    assert folder.display_name == "Projects"
    assert folder.parent_folder_id == "ROOT"
    assert folder.child_folder_count == 2
    assert folder.total_item_count == 10
    assert folder.unread_item_count == 3
