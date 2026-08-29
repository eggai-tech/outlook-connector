import httpx
import respx

from outlook_helper.folders import FolderResolver
from outlook_helper.http import GraphSession

BASE = "https://graph.microsoft.com/v1.0"


class FakeCredential:
    supports_me = True

    def get_token(self):
        return "tok"


def make_resolver(base_path="/me"):
    session = GraphSession(FakeCredential(), sleep=lambda s: None)
    return FolderResolver(session, base_path)


@respx.mock
def test_well_known_name_needs_no_lookup():
    route = respx.get(f"{BASE}/me/mailFolders")
    resolver = make_resolver()
    assert resolver.resolve("Inbox") == "inbox"
    assert resolver.resolve("sentitems") == "sentitems"
    assert route.call_count == 0


@respx.mock
def test_display_name_resolves_to_id_and_caches():
    route = respx.get(f"{BASE}/me/mailFolders").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "F123", "displayName": "Projects"}]}
        )
    )
    resolver = make_resolver()
    assert resolver.resolve("Projects") == "F123"
    # cached: a second resolve does not hit the network again
    assert resolver.resolve("Projects") == "F123"
    assert route.call_count == 1


@respx.mock
def test_unknown_reference_is_treated_as_raw_id():
    respx.get(f"{BASE}/me/mailFolders").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    resolver = make_resolver()
    assert resolver.resolve("AAMkRAWID==") == "AAMkRAWID=="


@respx.mock
def test_resolve_uses_configured_base_path():
    route = respx.get(f"{BASE}/users/u@x.com/mailFolders").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "F9", "displayName": "Team"}]}
        )
    )
    resolver = make_resolver(base_path="/users/u@x.com")
    assert resolver.resolve("Team") == "F9"
    assert route.call_count == 1
