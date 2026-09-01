import datetime
import os

import pytest
from pydantic import ValidationError

from outlook_connector.config import Settings, get_settings

# Env names Settings reads: ambient values (a developer's exported vars or the
# repo-root .env, which env_file points at) must not leak into the tests.
_SETTINGS_ENV = {
    "MAILBOX",
    "SOURCE_FOLDER",
    "POLL_INTERVAL_SECONDS",
    "BATCH_MAX_MESSAGES",
    "MAX_ATTACHMENT_BYTES",
    "INITIAL_CURSOR",
    "LOG_LEVEL",
    "HEALTH_PORT",
}


@pytest.fixture(autouse=True)
def isolated_env(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)  # env_file=".env" no longer finds the repo's
    for key in list(os.environ):
        if key in _SETTINGS_ENV or key.startswith(("BUS__", "AZURE_")):
            monkeypatch.delenv(key)


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
    assert settings.source_folder == "inbox"
    assert settings.batch_max_messages is None
    assert settings.max_attachment_bytes == 8 * 1024 * 1024


def test_polling_fields_configurable(tmp_path, monkeypatch, azure_env):
    _write_config(
        tmp_path,
        monkeypatch,
        "mailbox: inbox@example.com\n"
        "source_folder: Bankbestätigungen\n"
        "batch_max_messages: 50\n"
        "max_attachment_bytes: 1048576\n",
    )

    settings = Settings()

    assert settings.source_folder == "Bankbestätigungen"
    assert settings.batch_max_messages == 50
    assert settings.max_attachment_bytes == 1048576


def test_empty_azure_credential_rejected(tmp_path, monkeypatch, azure_env):
    """Compose files pass `${AZURE_...:-}` defaults; empty must fail fast."""
    _write_config(tmp_path, monkeypatch)
    monkeypatch.setenv("AZURE_TENANT_ID", "")

    with pytest.raises(ValidationError):
        Settings()


def test_naive_initial_cursor_coerced_to_utc(tmp_path, monkeypatch, azure_env):
    """A naive cursor would TypeError against Graph's aware datetimes on the
    first advance() — a crash-restart loop. Bare ISO strings count as UTC."""
    _write_config(
        tmp_path,
        monkeypatch,
        "mailbox: inbox@example.com\ninitial_cursor: 2026-06-26T08:00:00\n",
    )

    settings = Settings()

    assert settings.initial_cursor == datetime.datetime(
        2026, 6, 26, 8, 0, 0, tzinfo=datetime.UTC
    )


def test_azure_keys_in_yaml_rejected(tmp_path, monkeypatch, azure_env):
    _write_config(
        tmp_path, monkeypatch, "mailbox: inbox@example.com\nazure_client_secret: nope\n"
    )
    get_settings.cache_clear()

    with pytest.raises(ValueError, match="azure_client_secret"):
        get_settings()

    get_settings.cache_clear()
