import json

import httpx
import respx

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
def test_list_folders_returns_folders():
    respx.get(f"{BASE}/me/mailFolders").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {"id": "F1", "displayName": "Inbox"},
                    {"id": "F2", "displayName": "Archive"},
                ]
            },
        )
    )
    folders = make_client().list_folders()
    assert [f.display_name for f in folders] == ["Inbox", "Archive"]


@respx.mock
def test_create_folder_top_level():
    route = respx.post(f"{BASE}/me/mailFolders").mock(
        return_value=httpx.Response(201, json={"id": "NEW", "displayName": "Reports"})
    )
    folder = make_client().create_folder("Reports")
    assert folder.id == "NEW"
    assert folder.display_name == "Reports"
    assert body_of(route) == {"displayName": "Reports"}


@respx.mock
def test_create_folder_under_parent_display_name():
    respx.get(
        f"{BASE}/me/mailFolders", params={"$filter": "displayName eq 'Projects'"}
    ).mock(return_value=httpx.Response(200, json={"value": [{"id": "PARENT"}]}))
    route = respx.post(f"{BASE}/me/mailFolders/PARENT/childFolders").mock(
        return_value=httpx.Response(201, json={"id": "CHILD", "displayName": "2026"})
    )
    folder = make_client().create_folder("2026", parent="Projects")
    assert folder.id == "CHILD"
    assert body_of(route) == {"displayName": "2026"}


@respx.mock
def test_move_email_resolves_destination_and_returns_message():
    route = respx.post(f"{BASE}/me/messages/M1/move").mock(
        return_value=httpx.Response(
            201, json={"id": "M1-NEW", "subject": "moved"}
        )
    )
    msg = make_client().move_email("M1", "archive")
    assert msg.id == "M1-NEW"
    assert body_of(route) == {"destinationId": "archive"}
