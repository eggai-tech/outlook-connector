import pytest
from pydantic import ValidationError

from outlook_connector.config import Settings, get_settings


@pytest.fixture
def azure_env(monkeypatch):
    monkeypatch.setenv("AZURE_TENANT_ID", "t")
    monkeypatch.setenv("AZURE_CLIENT_ID", "c")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "s")


def _write_config(tmp_path, monkeypatch, text="mailbox: inbox@example.com\n"):
    cfg = tmp_path / "config.yaml"
    cfg.write_text(text)
    monkeypatch.setenv("CONFIG_FILE", str(cfg))
    return cfg


def test_loads_minimal_config(tmp_path, monkeypatch, azure_env):
    _write_config(tmp_path, monkeypatch)

    settings = Settings()

    assert settings.mailbox == "inbox@example.com"
    assert settings.bus.transport == "kafka"


def test_empty_azure_credential_rejected(tmp_path, monkeypatch, azure_env):
    """Compose files pass `${AZURE_...:-}` defaults; empty must fail fast."""
    _write_config(tmp_path, monkeypatch)
    monkeypatch.setenv("AZURE_TENANT_ID", "")

    with pytest.raises(ValidationError):
        Settings()


def test_azure_keys_in_yaml_rejected(tmp_path, monkeypatch, azure_env):
    _write_config(
        tmp_path, monkeypatch, "mailbox: inbox@example.com\nazure_client_secret: nope\n"
    )
    get_settings.cache_clear()

    with pytest.raises(ValueError, match="azure_client_secret"):
        get_settings()

    get_settings.cache_clear()
