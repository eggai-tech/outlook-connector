import pytest

from outlook_helper.auth import ClientSecretCredential, DeviceCodeCredential
from outlook_helper.config import AppOnlyConfig, DelegatedConfig
from outlook_helper.exceptions import GraphError


class FakeConfidentialApp:
    def __init__(self, result):
        self._result = result
        self.calls = []

    def acquire_token_for_client(self, scopes):
        self.calls.append(scopes)
        return self._result


class FakePublicApp:
    def __init__(self, *, accounts=None, silent=None, flow=None, by_flow=None):
        self._accounts = accounts or []
        self._silent = silent
        self._flow = flow or {"user_code": "ABCD", "message": "go here"}
        self._by_flow = by_flow
        self.flow_started = False

    def get_accounts(self):
        return self._accounts

    def acquire_token_silent(self, scopes, account):
        return self._silent

    def initiate_device_flow(self, scopes):
        return self._flow

    def acquire_token_by_device_flow(self, flow):
        self.flow_started = True
        return self._by_flow


# --- app-only / client credentials ---


def test_client_secret_credential_returns_access_token():
    app = FakeConfidentialApp({"access_token": "tok-123"})
    cred = ClientSecretCredential(
        AppOnlyConfig(client_id="c", tenant_id="t", client_secret="s"), app=app
    )
    assert cred.get_token() == "tok-123"
    assert app.calls == [["https://graph.microsoft.com/.default"]]


def test_client_secret_credential_raises_on_error():
    app = FakeConfidentialApp(
        {"error": "invalid_client", "error_description": "bad secret"}
    )
    cred = ClientSecretCredential(
        AppOnlyConfig(client_id="c", tenant_id="t", client_secret="s"), app=app
    )
    with pytest.raises(GraphError, match="bad secret"):
        cred.get_token()


def test_app_only_does_not_support_me():
    cred = ClientSecretCredential(
        AppOnlyConfig(client_id="c", tenant_id="t", client_secret="s"),
        app=FakeConfidentialApp({"access_token": "x"}),
    )
    assert cred.supports_me is False


# --- delegated / device code ---


def test_device_code_uses_cached_token_silently():
    app = FakePublicApp(accounts=["acct"], silent={"access_token": "silent-tok"})
    cred = DeviceCodeCredential(DelegatedConfig(client_id="c"), app=app)
    assert cred.get_token() == "silent-tok"
    assert app.flow_started is False


def test_device_code_runs_flow_and_prompts_when_no_cache():
    prompts = []
    app = FakePublicApp(
        accounts=[],
        flow={"user_code": "WXYZ", "message": "enter WXYZ"},
        by_flow={"access_token": "flow-tok"},
    )
    cred = DeviceCodeCredential(
        DelegatedConfig(client_id="c"), app=app, prompt_callback=prompts.append
    )
    assert cred.get_token() == "flow-tok"
    assert app.flow_started is True
    assert prompts and prompts[0]["user_code"] == "WXYZ"


def test_device_code_raises_when_flow_init_fails():
    app = FakePublicApp(accounts=[], flow={"error": "bad"})
    cred = DeviceCodeCredential(DelegatedConfig(client_id="c"), app=app)
    with pytest.raises(GraphError):
        cred.get_token()


def test_device_code_raises_when_token_acquisition_fails():
    app = FakePublicApp(
        accounts=[],
        flow={"user_code": "X", "message": "m"},
        by_flow={"error": "expired_token", "error_description": "timed out"},
    )
    cred = DeviceCodeCredential(DelegatedConfig(client_id="c"), app=app)
    with pytest.raises(GraphError, match="timed out"):
        cred.get_token()


def test_delegated_supports_me():
    app = FakePublicApp(accounts=["a"], silent={"access_token": "t"})
    cred = DeviceCodeCredential(DelegatedConfig(client_id="c"), app=app)
    assert cred.supports_me is True
