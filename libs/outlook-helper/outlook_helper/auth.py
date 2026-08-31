"""Credentials that produce Microsoft Graph bearer tokens.

Two flows are supported behind a common ``Credential`` protocol so the rest of
the library never branches on auth model:

* :class:`ClientSecretCredential` -- app-only (client credentials).
* :class:`DeviceCodeCredential` -- delegated (signed-in user) via device code.

Both accept an injected MSAL ``app`` for testing; in normal use they build the
real MSAL application (and, for the delegated flow, an encrypted persistent
token cache) themselves.
"""

from __future__ import annotations

import sys
from typing import Any, Callable, Protocol, runtime_checkable

from outlook_helper.config import AppOnlyConfig, DelegatedConfig
from outlook_helper.exceptions import GraphError

PromptCallback = Callable[[dict], None]


@runtime_checkable
class Credential(Protocol):
    """Anything that can mint a Graph bearer token."""

    #: Whether the ``/me`` shortcut is usable (delegated only).
    supports_me: bool

    def get_token(self) -> str:
        """Return a valid bearer token, refreshing or signing in as needed."""
        ...


def _authority(tenant_id: str) -> str:
    return f"https://login.microsoftonline.com/{tenant_id}"


def _raise_token_error(result: dict[str, Any], default: str) -> "None":
    message = result.get("error_description") or result.get("error") or default
    raise GraphError(status_code=401, message=message, code=result.get("error"))


class ClientSecretCredential:
    """App-only authentication using the client-credentials grant."""

    supports_me = False

    def __init__(self, config: AppOnlyConfig, *, app: Any | None = None) -> None:
        self._config = config
        self._scopes = list(config.scopes)
        self._app = app  # built lazily on first use to avoid I/O at construction

    def _ensure_app(self) -> Any:
        if self._app is None:
            self._app = self._build_app()
        return self._app

    def _build_app(self) -> Any:
        import msal

        credential: Any
        if self._config.certificate is not None:
            cert = self._config.certificate
            with open(cert.private_key_path) as fh:
                private_key = fh.read()
            credential = {"private_key": private_key, "thumbprint": cert.thumbprint}
            if cert.public_certificate_path is not None:
                with open(cert.public_certificate_path) as fh:
                    credential["public_certificate"] = fh.read()
        else:
            credential = self._config.client_secret

        return msal.ConfidentialClientApplication(
            client_id=self._config.client_id,
            authority=_authority(self._config.tenant_id),
            client_credential=credential,
        )

    def get_token(self) -> str:
        result = self._ensure_app().acquire_token_for_client(scopes=self._scopes)
        if "access_token" not in result:
            _raise_token_error(result, "Failed to acquire app-only token")
        return result["access_token"]


def _default_prompt(flow: dict) -> None:
    print(flow.get("message", ""), file=sys.stderr, flush=True)


class DeviceCodeCredential:
    """Delegated authentication via the OAuth 2.0 device-code flow."""

    supports_me = True

    def __init__(
        self,
        config: DelegatedConfig,
        *,
        app: Any | None = None,
        prompt_callback: PromptCallback | None = None,
    ) -> None:
        self._config = config
        self._scopes = list(config.scopes)
        self._prompt = prompt_callback or _default_prompt
        self._app = app  # built lazily on first use to avoid I/O at construction

    def _ensure_app(self) -> Any:
        if self._app is None:
            self._app = self._build_app()
        return self._app

    def _build_app(self) -> Any:
        import msal

        return msal.PublicClientApplication(
            client_id=self._config.client_id,
            authority=_authority(self._config.tenant_id),
            token_cache=_build_token_cache(self._config.cache_path),
        )

    def get_token(self) -> str:
        app = self._ensure_app()
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(self._scopes, account=accounts[0])
            if result and "access_token" in result:
                return result["access_token"]

        flow = app.initiate_device_flow(scopes=self._scopes)
        if "user_code" not in flow:
            _raise_token_error(flow, "Failed to start device-code flow")
        self._prompt(flow)

        result = app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            _raise_token_error(result, "Failed to acquire delegated token")
        return result["access_token"]


def _build_token_cache(cache_path: Any | None) -> Any:
    """Build a token cache, encrypted at rest when a path is given.

    Uses ``msal-extensions`` keyring-backed encryption when available, falling
    back to an unencrypted file cache (with a warning) and finally to an
    in-memory cache when no path is configured.
    """
    if cache_path is None:
        return None

    from msal_extensions import (
        FilePersistence,
        PersistedTokenCache,
        build_encrypted_persistence,
    )

    path = str(cache_path)
    try:
        persistence = build_encrypted_persistence(path)
    except Exception:
        import warnings

        warnings.warn(
            "Encrypted token cache unavailable; falling back to plaintext file cache",
            RuntimeWarning,
            stacklevel=2,
        )
        persistence = FilePersistence(path)
    return PersistedTokenCache(persistence)
