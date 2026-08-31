"""Configuration objects for the two supported authentication models.

The library is environment-agnostic: callers build one of these in code and pass
it to the matching credential. Nothing here reads environment variables.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

DELEGATED_DEFAULT_SCOPES: tuple[str, ...] = ("Mail.ReadWrite", "Mail.Send", "User.Read")
APP_ONLY_DEFAULT_SCOPES: tuple[str, ...] = ("https://graph.microsoft.com/.default",)


@dataclass(frozen=True)
class CertificateConfig:
    """A certificate credential for app-only authentication."""

    private_key_path: str
    thumbprint: str
    public_certificate_path: str | None = None


@dataclass(frozen=True)
class DelegatedConfig:
    """Delegated (signed-in user) authentication via the device-code flow."""

    client_id: str
    tenant_id: str = "common"
    scopes: tuple[str, ...] = DELEGATED_DEFAULT_SCOPES
    cache_path: Path | None = None


@dataclass(frozen=True)
class AppOnlyConfig:
    """App-only (client-credentials) authentication.

    Exactly one of ``client_secret`` or ``certificate`` must be supplied.
    """

    client_id: str
    tenant_id: str
    client_secret: str | None = None
    certificate: CertificateConfig | None = None
    scopes: tuple[str, ...] = APP_ONLY_DEFAULT_SCOPES

    def __post_init__(self) -> None:
        if self.client_secret is None and self.certificate is None:
            raise ValueError("AppOnlyConfig requires a client_secret or certificate")
        if self.client_secret is not None and self.certificate is not None:
            raise ValueError(
                "AppOnlyConfig accepts a client_secret or certificate, not both"
            )
