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


@respx.mock
def test_delete_email_soft_uses_delete_verb():
    route = respx.delete(f"{BASE}/me/messages/M1").mock(
        return_value=httpx.Response(204)
    )
    make_client().delete_email("M1")
    assert route.call_count == 1


@respx.mock
def test_delete_email_permanent_uses_permanent_delete_action():
    route = respx.post(f"{BASE}/me/messages/M1/permanentDelete").mock(
        return_value=httpx.Response(204)
    )
    make_client().delete_email("M1", permanent=True)
    assert route.call_count == 1
