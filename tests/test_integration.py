"""One poll cycle through the real stack, with Graph mocked at the HTTP layer.

Unlike the unit tests (which inject fakes above the client), this runs the
real outlook-helper ``OutlookClient`` against respx-mocked Graph responses:
query building, pagination shape, ``contentBytes`` base64 decoding, the
mapper, and the CloudEvents envelope are all exercised as one path. The only
fakes are the credential (no MSAL) and the channel (captures the published
event instead of a broker).
"""

import asyncio
import base64
import datetime

import httpx
import respx
from conftest import T0, FakeChannel
from outlook_helper.client import OutlookClient
from outlook_helper.http import GraphSession

from outlook_connector.bus import EMAIL_RECEIVED
from outlook_connector.poller import Poller
from outlook_connector.service import run_workflow

BASE = "https://graph.microsoft.com/v1.0"
MAILBOX = "inbox@example.com"
PDF = b"%PDF-1.4\xff\x00\x89"  # not valid UTF-8: must survive as base64


class FakeCredential:
    supports_me = False

    def get_token(self):
        return "tok"


def make_client() -> OutlookClient:
    session = GraphSession(FakeCredential(), sleep=lambda s: None)
    return OutlookClient(FakeCredential(), mailbox=MAILBOX, session=session)


def graph_message(message_id: str, *, received: str, has_attachments: bool = False) -> dict:
    return {
        "id": message_id,
        "subject": f"subject {message_id}",
        "from": {"emailAddress": {"name": "Sender", "address": "sender@example.com"}},
        "toRecipients": [{"emailAddress": {"address": MAILBOX}}],
        "receivedDateTime": received,
        "hasAttachments": has_attachments,
        "internetMessageId": f"<{message_id}@example.com>",
        "body": {"contentType": "html", "content": "<p>hi</p>"},
    }


@respx.mock
def test_full_poll_cycle_publishes_wire_correct_events():
    # the rescan lists the folder ids-only, then fetches each message in full
    listing = respx.get(f"{BASE}/users/{MAILBOX}/mailFolders/inbox/messages").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {"id": "m1", "receivedDateTime": "2026-01-01T12:00:10Z"},
                    {"id": "m2", "receivedDateTime": "2026-01-01T12:00:20Z"},
                ]
            },
        )
    )
    respx.get(f"{BASE}/users/{MAILBOX}/messages/m1").mock(
        return_value=httpx.Response(
            200, json=graph_message("m1", received="2026-01-01T12:00:10Z")
        )
    )
    respx.get(f"{BASE}/users/{MAILBOX}/messages/m2").mock(
        return_value=httpx.Response(
            200,
            json=graph_message("m2", received="2026-01-01T12:00:20Z", has_attachments=True),
        )
    )
    respx.get(f"{BASE}/users/{MAILBOX}/messages/m2/attachments").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "id": "att1",
                        "name": "doc.pdf",
                        "contentType": "application/pdf",
                        "size": len(PDF),
                        "contentBytes": base64.b64encode(PDF).decode(),
                    }
                ]
            },
        )
    )

    poller = Poller(client=make_client(), now=lambda: T0)
    channel = FakeChannel()
    context = {"poller": poller, "channel": channel, "source_mailbox": MAILBOX}

    summary = asyncio.run(run_workflow(context))

    assert (summary.fetched, summary.published, summary.dropped) == (2, 2, 0)
    assert summary.error is None
    assert listing.calls.last.request.url.params["$select"] == "id,receivedDateTime"
    assert poller._published_ids == {"m1", "m2"}

    # what a consumer receives after a JSON round trip (the wire format)
    events = [type(e).model_validate_json(e.model_dump_json()) for e in channel.published]
    assert [e.data.email.id for e in events] == ["m1", "m2"]
    assert all(e.type == EMAIL_RECEIVED for e in events)
    assert events[0].data.source_mailbox == MAILBOX
    assert events[0].data.email.internet_message_id == "<m1@example.com>"

    (attachment,) = events[1].data.email.attachments
    assert attachment.file_name == "doc.pdf"
    assert attachment.size == len(PDF)
    assert attachment.body == PDF
