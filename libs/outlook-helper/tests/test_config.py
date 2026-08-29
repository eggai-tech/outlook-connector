import pytest

from outlook_helper.config import AppOnlyConfig, CertificateConfig, DelegatedConfig


def test_delegated_config_defaults():
    cfg = DelegatedConfig(client_id="cid")
    assert cfg.client_id == "cid"
    assert cfg.tenant_id == "common"
    assert "Mail.ReadWrite" in cfg.scopes
    assert "Mail.Send" in cfg.scopes
    assert cfg.cache_path is None


def test_app_only_config_with_secret():
    cfg = AppOnlyConfig(client_id="cid", tenant_id="tid", client_secret="shh")
    assert cfg.client_secret == "shh"
    assert cfg.certificate is None
    assert cfg.scopes == ("https://graph.microsoft.com/.default",)


def test_app_only_config_with_certificate():
    cert = CertificateConfig(private_key_path="/k.pem", thumbprint="AA11")
    cfg = AppOnlyConfig(client_id="cid", tenant_id="tid", certificate=cert)
    assert cfg.certificate is cert
    assert cfg.client_secret is None


def test_app_only_config_requires_a_credential():
    with pytest.raises(ValueError, match="client_secret or certificate"):
        AppOnlyConfig(client_id="cid", tenant_id="tid")


def test_app_only_config_rejects_both_credentials():
    cert = CertificateConfig(private_key_path="/k.pem", thumbprint="AA11")
    with pytest.raises(ValueError, match="not both"):
        AppOnlyConfig(
            client_id="cid", tenant_id="tid", client_secret="shh", certificate=cert
        )
