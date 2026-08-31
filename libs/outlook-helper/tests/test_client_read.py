from datetime import datetime, timedelta, timezone

import httpx
import pytest
import respx

from outlook_helper.client import OutlookClient, _MESSAGE_SELECT, _fmt_dt
from outlook_helper.http import IMMUTABLE_ID_PREFERENCE, GraphSession

BASE = "https://graph.microsoft.com/v1.0"


class FakeCredential:
    def __init__(self, supports_me=True):
        self.supports_me = supports_me

    def get_token(self):
        return "tok"


def make_client(supports_me=True, mailbox=None):
    session = GraphSession(FakeCredential(supports_me), sleep=lambda s: None)
    return OutlookClient(
        FakeCredential(supports_me), mailbox=mailbox, session=session
    )


def test_app_only_without_mailbox_raises():
    session = GraphSession(FakeCredential(False), sleep=lambda s: None)
    with pytest.raises(ValueError, match="mailbox"):
        OutlookClient(FakeCredential(supports_me=False), session=session)


def test_app_only_with_mailbox_uses_users_path():
    client = make_client(supports_me=False, mailbox="u@x.com")
    assert client.base_path == "/users/u@x.com"


def test_delegated_defaults_to_me():
    client = make_client()
    assert client.base_path == "/me"


@respx.mock
def test_get_email_returns_message(load_fixture):
    respx.get(f"{BASE}/me/messages/AAMkAGI1").mock(
        return_value=httpx.Response(200, json=load_fixture("message.json"))
    )
    msg = make_client().get_email("AAMkAGI1")
    assert msg.subject == "Quarterly report"
    assert msg.from_.address == "alice@example.com"


@respx.mock
def test_list_messages_default_inbox_orders_newest_first():
    route = respx.get(f"{BASE}/me/mailFolders/inbox/messages").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "1"}, {"id": "2"}]}
        )
    )
    msgs = list(make_client().list_messages())
    assert [m.id for m in msgs] == ["1", "2"]
    assert route.calls.last.request.url.params["$orderby"] == "receivedDateTime desc"


@respx.mock
def test_list_messages_top_caps_results():
    respx.get(f"{BASE}/me/mailFolders/inbox/messages").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "1"}, {"id": "2"}, {"id": "3"}]}
        )
    )
    msgs = list(make_client().list_messages(top=2))
    assert [m.id for m in msgs] == ["1", "2"]


@respx.mock
def test_list_messages_resolves_display_name_folder():
    respx.get(f"{BASE}/me/mailFolders", params={"$filter": "displayName eq 'Projects'"}).mock(
        return_value=httpx.Response(200, json={"value": [{"id": "PF1"}]})
    )
    msgs_route = respx.get(f"{BASE}/me/mailFolders/PF1/messages").mock(
        return_value=httpx.Response(200, json={"value": [{"id": "9"}]})
    )
    msgs = list(make_client().list_messages(folder="Projects"))
    assert [m.id for m in msgs] == ["9"]
    assert msgs_route.call_count == 1


@respx.mock
def test_search_email_builds_filter():
    route = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(200, json={"value": [{"id": "1"}]})
    )
    list(
        make_client().search_email(
            subject_contains="ABC",
            since=datetime(2026, 6, 1),
            has_attachments=True,
        )
    )
    flt = route.calls.last.request.url.params["$filter"]
    assert "contains(subject,'ABC')" in flt
    assert "receivedDateTime ge 2026-06-01T00:00:00.000000Z" in flt
    assert "hasAttachments eq true" in flt
    assert " and " in flt


def test_fmt_dt_preserves_subsecond_precision_for_naive_utc():
    assert (
        _fmt_dt(datetime(2026, 6, 1, 9, 30, 15, 123456))
        == "2026-06-01T09:30:15.123456Z"
    )


def test_fmt_dt_converts_aware_datetime_to_utc():
    aware = datetime(2026, 6, 1, 11, 30, 0, tzinfo=timezone(timedelta(hours=2)))
    assert _fmt_dt(aware) == "2026-06-01T09:30:00.000000Z"


def test_fmt_dt_passes_strings_through():
    assert _fmt_dt("2026-06-01T00:00:00Z") == "2026-06-01T00:00:00Z"


@respx.mock
def test_search_email_strict_operators():
    route = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    list(
        make_client().search_email(
            since_exclusive=datetime(2026, 6, 1, 0, 0, 0),
            until_exclusive=datetime(2026, 6, 2, 0, 0, 0),
        )
    )
    flt = route.calls.last.request.url.params["$filter"]
    assert "receivedDateTime gt 2026-06-01T00:00:00.000000Z" in flt
    assert "receivedDateTime lt 2026-06-02T00:00:00.000000Z" in flt


@respx.mock
def test_search_email_include_headers_applies_full_select():
    route = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    list(make_client().search_email(include_headers=True))
    assert route.calls.last.request.url.params["$select"] == _MESSAGE_SELECT
    assert "internetMessageHeaders" in _MESSAGE_SELECT


@respx.mock
def test_search_email_html_body_sets_prefer_header():
    route = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    list(make_client().search_email(html_body=True))
    assert (
        route.calls.last.request.headers["Prefer"]
        == f'{IMMUTABLE_ID_PREFERENCE}, outlook.body-content-type="html"'
    )


@respx.mock
def test_search_email_without_options_has_no_select_or_prefer():
    route = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    list(make_client().search_email())
    assert "$select" not in route.calls.last.request.url.params
    # The only preference is the always-on immutable-id request.
    assert route.calls.last.request.headers["Prefer"] == IMMUTABLE_ID_PREFERENCE


@respx.mock
def test_get_email_supports_headers_and_html_body(load_fixture):
    route = respx.get(f"{BASE}/me/messages/AAMkAGI1").mock(
        return_value=httpx.Response(200, json=load_fixture("message.json"))
    )
    make_client().get_email("AAMkAGI1", include_headers=True, html_body=True)
    req = route.calls.last.request
    assert req.url.params["$select"] == _MESSAGE_SELECT
    assert req.headers["Prefer"] == (
        f'{IMMUTABLE_ID_PREFERENCE}, outlook.body-content-type="html"'
    )


@respx.mock
def test_get_email_with_include_mime_fetches_the_raw_mime(load_fixture):
    respx.get(f"{BASE}/me/messages/AAMkAGI1").mock(
        return_value=httpx.Response(200, json=load_fixture("message.json"))
    )
    route = respx.get(f"{BASE}/me/messages/AAMkAGI1/$value").mock(
        return_value=httpx.Response(
            200,
            headers={"Content-Type": "text/plain"},
            text="Subject: Quarterly report\r\n\r\nBody text\r\n",
        )
    )
    msg = make_client().get_email("AAMkAGI1", include_mime=True)
    assert msg.subject == "Quarterly report"
    assert msg.mime_content == "Subject: Quarterly report\r\n\r\nBody text\r\n"
    assert route.calls.last.request.headers["Prefer"] == IMMUTABLE_ID_PREFERENCE


@respx.mock
def test_get_email_without_include_mime_leaves_mime_content_unset(load_fixture):
    respx.get(f"{BASE}/me/messages/AAMkAGI1").mock(
        return_value=httpx.Response(200, json=load_fixture("message.json"))
    )
    value_route = respx.get(f"{BASE}/me/messages/AAMkAGI1/$value")
    msg = make_client().get_email("AAMkAGI1")
    assert msg.mime_content is None
    assert not value_route.called


@respx.mock
def test_list_messages_supports_headers_and_html_body():
    route = respx.get(f"{BASE}/me/mailFolders/inbox/messages").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    list(make_client().list_messages(include_headers=True, html_body=True))
    req = route.calls.last.request
    assert req.url.params["$select"] == _MESSAGE_SELECT
    assert req.headers["Prefer"] == (
        f'{IMMUTABLE_ID_PREFERENCE}, outlook.body-content-type="html"'
    )


@respx.mock
def test_search_email_sender_and_unread_filters():
    route = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    list(make_client().search_email(sender="boss@x.com", unread=True))
    flt = route.calls.last.request.url.params["$filter"]
    assert "from/emailAddress/address eq 'boss@x.com'" in flt
    assert "isRead eq false" in flt


@respx.mock
def test_list_attachments_returns_meta():
    respx.get(f"{BASE}/me/messages/M1/attachments").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "a1",
                        "name": "f.pdf",
                        "contentType": "application/pdf",
                        "size": 10,
                        "isInline": False,
                    }
                ]
            },
        )
    )
    atts = make_client().list_attachments("M1")
    assert len(atts) == 1
    assert atts[0].name == "f.pdf"


@respx.mock
def test_list_attachments_selects_metadata_only():
    route = respx.get(f"{BASE}/me/messages/M1/attachments").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    make_client().list_attachments("M1")
    select = route.calls.last.request.url.params["$select"]
    assert select == "id,name,contentType,size,isInline"
    assert "contentBytes" not in select


@respx.mock
def test_get_attachments_decodes_content_and_omits_select():
    route = respx.get(f"{BASE}/me/messages/M1/attachments").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "id": "a1",
                        "name": "f.pdf",
                        "contentType": "application/pdf",
                        "size": 5,
                        "isInline": False,
                        "contentBytes": "aGVsbG8=",  # base64 of b"hello"
                    }
                ]
            },
        )
    )
    atts = make_client().get_attachments("M1")
    assert len(atts) == 1
    assert atts[0].name == "f.pdf"
    assert atts[0].content == b"hello"
    # No metadata-only $select: contentBytes must come back inline.
    assert "$select" not in route.calls.last.request.url.params


@respx.mock
def test_get_attachments_content_none_when_no_bytes():
    respx.get(f"{BASE}/me/messages/M1/attachments").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "@odata.type": "#microsoft.graph.itemAttachment",
                        "id": "a2",
                        "name": "forwarded.eml",
                        "contentType": "message/rfc822",
                        "size": 99,
                        "isInline": False,
                    }
                ]
            },
        )
    )
    atts = make_client().get_attachments("M1")
    assert len(atts) == 1
    assert atts[0].content is None
    assert atts[0].name == "forwarded.eml"
    assert atts[0].size == 99


@respx.mock
def test_get_attachments_follows_pagination():
    next_link = f"{BASE}/me/messages/M1/attachments?%24skiptoken=X"
    # First page must match ONLY the unparameterized request; without params__eq
    # this route also matches the $skiptoken follow-up (respx ignores query when
    # none is specified), re-serving a page with @odata.nextLink → infinite loop.
    respx.get(f"{BASE}/me/messages/M1/attachments", params__eq={}).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [{"id": "a1", "name": "one", "contentBytes": "AAA="}],
                "@odata.nextLink": next_link,
            },
        )
    )
    respx.get(next_link).mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "a2", "name": "two", "contentBytes": "AQE="}]},
        )
    )
    atts = make_client().get_attachments("M1")
    assert [a.id for a in atts] == ["a1", "a2"]


@respx.mock
def test_get_email_returns_the_immutable_id(load_fixture):
    """With ``Prefer: IdType="ImmutableId"`` Graph puts the immutable id in ``id``."""
    route = respx.get(f"{BASE}/me/messages/AAkALgAAAAA").mock(
        return_value=httpx.Response(200, json={"id": "AAkALgAAAAA", "subject": "hi"})
    )
    msg = make_client().get_email("AAkALgAAAAA")
    assert route.calls.last.request.headers["Prefer"] == IMMUTABLE_ID_PREFERENCE
    assert msg.id == "AAkALgAAAAA"


@respx.mock
def test_list_messages_requests_immutable_ids():
    route = respx.get(f"{BASE}/me/mailFolders/inbox/messages").mock(
        return_value=httpx.Response(200, json={"value": [{"id": "AAkALg"}]})
    )
    assert [m.id for m in make_client().list_messages()] == ["AAkALg"]
    assert route.calls.last.request.headers["Prefer"] == IMMUTABLE_ID_PREFERENCE


@respx.mock
def test_move_email_keeps_the_immutable_id():
    """The point of immutable ids: a move no longer changes ``.id``."""
    route = respx.post(f"{BASE}/me/messages/AAkALg/move").mock(
        return_value=httpx.Response(201, json={"id": "AAkALg", "subject": "hi"})
    )
    respx.get(f"{BASE}/me/mailFolders").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "f1", "displayName": "Processed"}]}
        )
    )
    moved = make_client().move_email("AAkALg", "Processed")
    assert route.calls.last.request.headers["Prefer"] == IMMUTABLE_ID_PREFERENCE
    assert moved.id == "AAkALg"
