import base64
import json

import httpx
import respx

from outlook_helper.attachments import Attachment, LARGE_ATTACHMENT_THRESHOLD
from outlook_helper.client import OutlookClient
from outlook_helper.http import GraphSession

BASE = "https://graph.microsoft.com/v1.0"


class FakeCredential:
    supports_me = True

    def get_token(self):
        return "tok"


def make_client():
    session = GraphSession(FakeCredential(), sleep=lambda s: None)
    return OutlookClient(FakeCredential(), session=session)


def body_of(route):
    return json.loads(route.calls.last.request.content)


@respx.mock
def test_send_email_plain_uses_sendmail():
    route = respx.post(f"{BASE}/me/sendMail").mock(
        return_value=httpx.Response(202)
    )
    make_client().send_email("bob@x.com", "Hi", "Body text")
    sent = body_of(route)
    assert sent["saveToSentItems"] is True
    msg = sent["message"]
    assert msg["subject"] == "Hi"
    assert msg["body"] == {"contentType": "text", "content": "Body text"}
    assert msg["toRecipients"] == [{"emailAddress": {"address": "bob@x.com"}}]


@respx.mock
def test_send_email_html_and_multiple_recipients():
    route = respx.post(f"{BASE}/me/sendMail").mock(return_value=httpx.Response(202))
    make_client().send_email(
        ["a@x.com", "b@x.com"], "S", "<b>hi</b>", cc="c@x.com", html=True
    )
    msg = body_of(route)["message"]
    assert msg["body"]["contentType"] == "html"
    assert [r["emailAddress"]["address"] for r in msg["toRecipients"]] == [
        "a@x.com",
        "b@x.com",
    ]
    assert msg["ccRecipients"] == [{"emailAddress": {"address": "c@x.com"}}]


@respx.mock
def test_send_email_small_attachment_inline_via_sendmail():
    route = respx.post(f"{BASE}/me/sendMail").mock(return_value=httpx.Response(202))
    att = Attachment("note.txt", b"hello", "text/plain")
    make_client().send_email("bob@x.com", "Hi", "body", attachments=[att])
    msg = body_of(route)["message"]
    assert len(msg["attachments"]) == 1
    payload = msg["attachments"][0]
    assert payload["name"] == "note.txt"
    assert base64.b64decode(payload["contentBytes"]) == b"hello"


@respx.mock
def test_send_email_large_attachment_uses_draft_and_upload_session():
    big = Attachment("big.bin", b"x" * (LARGE_ATTACHMENT_THRESHOLD + 5), "application/octet-stream")

    draft_route = respx.post(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(201, json={"id": "DRAFT1"})
    )
    session_route = respx.post(
        f"{BASE}/me/messages/DRAFT1/attachments/createUploadSession"
    ).mock(
        return_value=httpx.Response(
            201, json={"uploadUrl": "https://upload.example/up"}
        )
    )
    upload_route = respx.put("https://upload.example/up").mock(
        return_value=httpx.Response(200, json={"nextExpectedRanges": ["0-"]})
    )
    send_route = respx.post(f"{BASE}/me/messages/DRAFT1/send").mock(
        return_value=httpx.Response(202)
    )

    make_client().send_email("bob@x.com", "Big", "body", attachments=[big])

    assert draft_route.call_count == 1
    assert session_route.call_count == 1
    assert upload_route.call_count >= 1
    assert send_route.call_count == 1
    # sendMail is never registered here; respx would raise if it were called.


@respx.mock
def test_create_draft_returns_message_without_sending():
    draft_route = respx.post(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(201, json={"id": "D1", "subject": "Draft"})
    )
    draft = make_client().create_draft("bob@x.com", "Draft", "body")
    assert draft.id == "D1"
    assert draft_route.call_count == 1
    assert body_of(draft_route)["subject"] == "Draft"


@respx.mock
def test_update_draft_patches_fields():
    route = respx.patch(f"{BASE}/me/messages/D1").mock(
        return_value=httpx.Response(200, json={"id": "D1", "subject": "New"})
    )
    msg = make_client().update_draft("D1", subject="New")
    assert msg.subject == "New"
    assert body_of(route) == {"subject": "New"}


@respx.mock
def test_send_draft_posts_send():
    route = respx.post(f"{BASE}/me/messages/D1/send").mock(
        return_value=httpx.Response(202)
    )
    make_client().send_draft("D1")
    assert route.call_count == 1


@respx.mock
def test_discard_draft_deletes():
    route = respx.delete(f"{BASE}/me/messages/D1").mock(
        return_value=httpx.Response(204)
    )
    make_client().discard_draft("D1")
    assert route.call_count == 1


@respx.mock
def test_reply_creates_reply_patches_body_and_sends():
    create = respx.post(f"{BASE}/me/messages/M1/createReply").mock(
        return_value=httpx.Response(201, json={"id": "R1"})
    )
    patch = respx.patch(f"{BASE}/me/messages/R1").mock(
        return_value=httpx.Response(200, json={"id": "R1"})
    )
    send = respx.post(f"{BASE}/me/messages/R1/send").mock(
        return_value=httpx.Response(202)
    )
    make_client().reply("M1", "Thanks!")
    assert create.call_count == 1
    assert body_of(patch)["body"] == {"contentType": "text", "content": "Thanks!"}
    assert send.call_count == 1


@respx.mock
def test_reply_all_uses_create_reply_all():
    create = respx.post(f"{BASE}/me/messages/M1/createReplyAll").mock(
        return_value=httpx.Response(201, json={"id": "R2"})
    )
    respx.patch(f"{BASE}/me/messages/R2").mock(
        return_value=httpx.Response(200, json={"id": "R2"})
    )
    respx.post(f"{BASE}/me/messages/R2/send").mock(return_value=httpx.Response(202))
    make_client().reply("M1", "Reply all", reply_all=True)
    assert create.call_count == 1


@respx.mock
def test_create_reply_draft_returns_draft_with_conversation_id():
    create = respx.post(f"{BASE}/me/messages/M1/createReply").mock(
        return_value=httpx.Response(
            201, json={"id": "R1", "conversationId": "CONV1"}
        )
    )
    draft = make_client().create_reply_draft("M1")
    assert create.call_count == 1
    assert draft.id == "R1"
    assert draft.conversation_id == "CONV1"


@respx.mock
def test_create_reply_draft_reply_all_uses_create_reply_all():
    create = respx.post(f"{BASE}/me/messages/M1/createReplyAll").mock(
        return_value=httpx.Response(
            201, json={"id": "R2", "conversationId": "CONV2"}
        )
    )
    draft = make_client().create_reply_draft("M1", reply_all=True)
    assert create.call_count == 1
    assert draft.conversation_id == "CONV2"


@respx.mock
def test_threaded_reply_redirects_recipients_to_shared_mailbox():
    """Compose create_reply_draft + update_draft + send_draft to thread a
    message into a shared mailbox without notifying the original sender."""
    respx.post(f"{BASE}/me/messages/M1/createReply").mock(
        return_value=httpx.Response(
            201, json={"id": "R1", "conversationId": "CONV1"}
        )
    )
    patch = respx.patch(f"{BASE}/me/messages/R1").mock(
        return_value=httpx.Response(200, json={"id": "R1"})
    )
    send = respx.post(f"{BASE}/me/messages/R1/send").mock(
        return_value=httpx.Response(202)
    )

    client = make_client()
    draft = client.create_reply_draft("M1")
    client.update_draft(
        draft.id,
        to="sandbox@example.com",
        body="<p>Extra information</p>",
        html=True,
    )
    client.send_draft(draft.id)

    patched = body_of(patch)
    assert patched["toRecipients"] == [
        {"emailAddress": {"address": "sandbox@example.com"}}
    ]
    # Alice (the original sender) must not be a recipient.
    addresses = [r["emailAddress"]["address"] for r in patched["toRecipients"]]
    assert "alice@example.com" not in addresses
    assert patched["body"] == {
        "contentType": "html",
        "content": "<p>Extra information</p>",
    }
    assert send.call_count == 1


@respx.mock
def test_download_attachment_streams_to_file(tmp_path):
    respx.get(f"{BASE}/me/messages/M1/attachments/A1/$value").mock(
        return_value=httpx.Response(200, content=b"attachment-bytes")
    )
    dest = tmp_path / "out.bin"
    result = make_client().download_attachment("M1", "A1", dest)
    assert result == dest
    assert dest.read_bytes() == b"attachment-bytes"
