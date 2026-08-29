import httpx
import pytest
import respx

from outlook_helper.exceptions import GraphError
from outlook_helper.http import GraphSession

BASE = "https://graph.microsoft.com/v1.0"


class FakeCredential:
    supports_me = True

    def __init__(self, token="tok-xyz"):
        self._token = token

    def get_token(self):
        return self._token


def make_session(**kwargs):
    sleeps = []
    session = GraphSession(
        FakeCredential(),
        sleep=lambda s: sleeps.append(s),
        **kwargs,
    )
    return session, sleeps


@respx.mock
def test_get_json_injects_bearer_token():
    route = respx.get(f"{BASE}/me/messages/1").mock(
        return_value=httpx.Response(200, json={"id": "1", "subject": "hi"})
    )
    session, _ = make_session()
    data = session.get_json("/me/messages/1")
    assert data["subject"] == "hi"
    assert route.calls.last.request.headers["Authorization"] == "Bearer tok-xyz"


@respx.mock
def test_non_2xx_maps_to_graph_error():
    respx.get(f"{BASE}/me/messages/missing").mock(
        return_value=httpx.Response(
            404,
            json={"error": {"code": "ErrorItemNotFound", "message": "not found"}},
            headers={"request-id": "req-7"},
        )
    )
    session, _ = make_session()
    with pytest.raises(GraphError) as exc:
        session.get_json("/me/messages/missing")
    assert exc.value.status_code == 404
    assert exc.value.code == "ErrorItemNotFound"
    assert exc.value.message == "not found"
    assert exc.value.request_id == "req-7"


@respx.mock
def test_retries_on_429_then_succeeds_and_honors_retry_after():
    route = respx.get(f"{BASE}/me/messages")
    route.side_effect = [
        httpx.Response(429, headers={"Retry-After": "2"}, json={}),
        httpx.Response(200, json={"value": []}),
    ]
    session, sleeps = make_session()
    data = session.get_json("/me/messages")
    assert data == {"value": []}
    assert route.call_count == 2
    assert sleeps == [2.0]


@respx.mock
def test_retries_exhausted_raises():
    route = respx.get(f"{BASE}/me/messages")
    route.side_effect = [httpx.Response(503, json={}) for _ in range(5)]
    session, sleeps = make_session(max_retries=2)
    with pytest.raises(GraphError) as exc:
        session.get_json("/me/messages")
    assert exc.value.status_code == 503
    # initial try + 2 retries = 3 calls
    assert route.call_count == 3
    assert len(sleeps) == 2


@respx.mock
def test_paginate_follows_next_link_lazily():
    # Register the param-constrained page first so the nextLink request (which
    # carries ?$skip=1) matches it rather than the param-less first-page route.
    page2 = respx.get(f"{BASE}/me/messages", params={"$skip": "1"}).mock(
        return_value=httpx.Response(200, json={"value": [{"id": "2"}]})
    )
    page1 = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [{"id": "1"}],
                "@odata.nextLink": f"{BASE}/me/messages?$skip=1",
            },
        )
    )
    session, _ = make_session()

    it = session.paginate("/me/messages")
    first = next(it)
    assert first["id"] == "1"
    # second page not fetched until we ask for more
    assert page2.call_count == 0

    rest = list(it)
    assert [m["id"] for m in rest] == ["2"]
    assert page1.call_count == 1
    assert page2.call_count == 1


@respx.mock
def test_get_json_forwards_extra_headers():
    route = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    session, _ = make_session()
    session.get_json("/me/messages", headers={"Prefer": 'outlook.body-content-type="html"'})
    assert (
        route.calls.last.request.headers["Prefer"]
        == 'outlook.body-content-type="html"'
    )


@respx.mock
def test_paginate_forwards_headers_on_every_page():
    page2 = respx.get(f"{BASE}/me/messages", params={"$skip": "1"}).mock(
        return_value=httpx.Response(200, json={"value": [{"id": "2"}]})
    )
    page1 = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [{"id": "1"}],
                "@odata.nextLink": f"{BASE}/me/messages?$skip=1",
            },
        )
    )
    session, _ = make_session()
    list(session.paginate("/me/messages", headers={"Prefer": "pref"}))
    assert page1.calls.last.request.headers["Prefer"] == "pref"
    assert page2.calls.last.request.headers["Prefer"] == "pref"


@respx.mock
def test_upload_chunk_puts_without_auth_and_sets_content_range():
    route = respx.put("https://upload.example/session").mock(
        return_value=httpx.Response(200, json={"nextExpectedRanges": ["4-"]})
    )
    session, _ = make_session()
    resp = session.upload_chunk(
        "https://upload.example/session", b"data", "bytes 0-3/10"
    )
    assert resp.status_code == 200
    req = route.calls.last.request
    assert "Authorization" not in req.headers
    assert req.headers["Content-Range"] == "bytes 0-3/10"
    assert req.content == b"data"


@respx.mock
def test_upload_chunk_raises_on_error():
    respx.put("https://upload.example/session").mock(
        return_value=httpx.Response(416, json={"error": {"message": "bad range"}})
    )
    session, _ = make_session()
    with pytest.raises(GraphError):
        session.upload_chunk("https://upload.example/session", b"data", "bytes 0-3/10")


@respx.mock
def test_download_streams_to_file(tmp_path):
    respx.get(f"{BASE}/me/messages/1/attachments/a/$value").mock(
        return_value=httpx.Response(200, content=b"file-bytes-here")
    )
    session, _ = make_session()
    dest = tmp_path / "out.bin"
    result = session.download("/me/messages/1/attachments/a/$value", dest)
    assert result == dest
    assert dest.read_bytes() == b"file-bytes-here"
