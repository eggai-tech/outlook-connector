"""Tests for the layered service configuration.

Structural config comes from a YAML file; every Azure connection parameter comes
from the environment only (``AZURE_TENANT_ID`` / ``AZURE_CLIENT_ID`` /
``AZURE_CLIENT_SECRET``) and is rejected outright if found in the YAML file.

Runnable standalone (`python -m tests.test_config`) or under pytest.
"""

import os
from pathlib import Path

import pytest
from pydantic import ValidationError

from outlook_connector.config import load_settings

_VALID_YAML = """\
mailboxes:
  - invoices@egg-ai.com
  - support@egg-ai.com
poll_interval_seconds: 30
log_level: DEBUG
bus:
  transport: kafka
  broker_url: broker:19092
  channel: emails
"""

_AZURE_ENV = {
    "AZURE_TENANT_ID": "tenant-123",
    "AZURE_CLIENT_ID": "client-abc",
    "AZURE_CLIENT_SECRET": "s3cret",
}


def _write(tmp_path: Path, text: str) -> Path:
    cfg = tmp_path / "config.yaml"
    cfg.write_text(text)
    return cfg


def _load(tmp_path: Path, text: str, monkeypatch, *, omit: str | None = None, **env):
    """Load ``text`` as the config file with the Azure env vars set.

    ``omit`` drops one Azure variable; ``env`` overrides individual values.
    """
    cfg = _write(tmp_path, text)
    monkeypatch.setenv("CONFIG_FILE", str(cfg))
    for key, value in {**_AZURE_ENV, **env}.items():
        monkeypatch.setenv(key, value)
    if omit is not None:
        monkeypatch.delenv(omit, raising=False)
    return load_settings()


def test_loads_structural_config_from_yaml(tmp_path, monkeypatch):
    settings = _load(tmp_path, _VALID_YAML, monkeypatch)

    assert settings.mailboxes == ["invoices@egg-ai.com", "support@egg-ai.com"]
    assert settings.poll_interval_seconds == 30
    assert settings.log_level == "DEBUG"
    assert settings.bus.transport == "kafka"
    assert settings.bus.broker_url == "broker:19092"
    assert settings.bus.channel == "emails"


def test_azure_credentials_come_from_env(tmp_path, monkeypatch):
    settings = _load(
        tmp_path, _VALID_YAML, monkeypatch, AZURE_CLIENT_SECRET="top-secret"
    )
    assert settings.azure_tenant_id == "tenant-123"
    assert settings.azure_client_id == "client-abc"
    assert settings.azure_client_secret == "top-secret"


def test_defaults_applied(tmp_path, monkeypatch):
    minimal = """\
mailboxes:
  - invoices@egg-ai.com
bus:
  broker_url: broker:19092
"""
    settings = _load(tmp_path, minimal, monkeypatch)
    assert settings.poll_interval_seconds == 60.0
    assert settings.log_level == "INFO"
    assert settings.bus.transport == "kafka"
    assert settings.bus.channel == "emails"


def test_missing_client_secret_fails_fast(tmp_path, monkeypatch):
    with pytest.raises(ValidationError):
        _load(tmp_path, _VALID_YAML, monkeypatch, omit="AZURE_CLIENT_SECRET")


def test_missing_tenant_id_fails_fast(tmp_path, monkeypatch):
    with pytest.raises(ValidationError):
        _load(tmp_path, _VALID_YAML, monkeypatch, omit="AZURE_TENANT_ID")


def test_missing_required_structural_field_fails_fast(tmp_path, monkeypatch):
    no_mailboxes = """\
bus:
  broker_url: broker:19092
"""
    with pytest.raises(ValidationError):
        _load(tmp_path, no_mailboxes, monkeypatch)


def test_empty_mailbox_list_fails_fast(tmp_path, monkeypatch):
    no_mailboxes = """\
mailboxes: []
bus:
  broker_url: broker:19092
"""
    with pytest.raises(ValidationError):
        _load(tmp_path, no_mailboxes, monkeypatch)


def test_secret_in_config_file_is_rejected(tmp_path, monkeypatch):
    """A secret in the config file is an error, never a silently ignored value."""
    with_secret = _VALID_YAML + "azure_client_secret: from-file\n"
    with pytest.raises(ValueError, match="must not contain azure_client_secret"):
        _load(tmp_path, with_secret, monkeypatch)


def test_legacy_azure_block_is_rejected(tmp_path, monkeypatch):
    """A pre-migration ``azure:`` block fails loudly instead of looking effective."""
    legacy = _VALID_YAML + "azure:\n  tenant_id: t\n  client_id: c\n"
    with pytest.raises(ValueError, match="must not contain azure"):
        _load(tmp_path, legacy, monkeypatch)


def test_initial_cursor_defaults_to_none(tmp_path, monkeypatch):
    settings = _load(tmp_path, _VALID_YAML, monkeypatch)
    assert settings.initial_cursor is None


def test_initial_cursor_parses_iso_datetime(tmp_path, monkeypatch):
    from datetime import datetime, timezone

    yaml = _VALID_YAML + "initial_cursor: 2026-06-26T08:00:00Z\n"
    settings = _load(tmp_path, yaml, monkeypatch)
    assert settings.initial_cursor == datetime(2026, 6, 26, 8, 0, tzinfo=timezone.utc)


def test_storage_backends_default_to_none_active(tmp_path, monkeypatch):
    """Storing nothing is a valid configuration."""
    settings = _load(tmp_path, _VALID_YAML, monkeypatch)
    assert settings.storage_backends == []


def test_storage_backends_loaded_from_yaml(tmp_path, monkeypatch):
    yaml = _VALID_YAML + "storage_backends:\n  - memory\n"
    settings = _load(tmp_path, yaml, monkeypatch)
    assert settings.storage_backends == ["memory"]


def test_storage_backends_loaded_from_env(tmp_path, monkeypatch):
    """``STORAGE_BACKENDS`` is a JSON list on the environment."""
    settings = _load(tmp_path, _VALID_YAML, monkeypatch, STORAGE_BACKENDS='["memory"]')
    assert settings.storage_backends == ["memory"]


def test_unknown_storage_backend_fails_at_config_load(tmp_path, monkeypatch):
    """A name no backend answers to is a startup error, not a runtime surprise."""
    yaml = _VALID_YAML + "storage_backends:\n  - carrier-pigeon\n"
    with pytest.raises(ValidationError, match="unknown storage backend"):
        _load(tmp_path, yaml, monkeypatch)


def test_repeated_storage_backend_fails_at_config_load(tmp_path, monkeypatch):
    """A name may appear only once — at most one instance of each backend."""
    yaml = _VALID_YAML + "storage_backends:\n  - memory\n  - memory\n"
    with pytest.raises(ValidationError, match="listed more than once"):
        _load(tmp_path, yaml, monkeypatch)


def test_unknown_transport_fails_fast(tmp_path, monkeypatch):
    bad = _VALID_YAML.replace("transport: kafka", "transport: carrier-pigeon")
    with pytest.raises(ValidationError):
        _load(tmp_path, bad, monkeypatch)


def test_missing_config_file_fails_fast(tmp_path, monkeypatch):
    monkeypatch.setenv("CONFIG_FILE", str(tmp_path / "does-not-exist.yaml"))
    for key, value in _AZURE_ENV.items():
        monkeypatch.setenv(key, value)
    with pytest.raises(FileNotFoundError):
        load_settings()


if __name__ == "__main__":
    import tempfile

    class _MP:
        """Minimal monkeypatch stand-in for standalone runs."""

        def __init__(self):
            self._saved = {}

        def setenv(self, k, v):
            self._saved.setdefault(k, os.environ.get(k))
            os.environ[k] = v

        def delenv(self, k, raising=True):
            self._saved.setdefault(k, os.environ.get(k))
            os.environ.pop(k, None)

        def undo(self):
            for k, v in self._saved.items():
                if v is None:
                    os.environ.pop(k, None)
                else:
                    os.environ[k] = v
            self._saved.clear()

    def _run(fn):
        mp = _MP()
        with tempfile.TemporaryDirectory() as d:
            try:
                fn(Path(d), mp)
            finally:
                mp.undo()

    _run(test_loads_structural_config_from_yaml)
    _run(test_azure_credentials_come_from_env)
    _run(test_defaults_applied)
    _run(test_missing_client_secret_fails_fast)
    _run(test_missing_tenant_id_fails_fast)
    _run(test_missing_required_structural_field_fails_fast)
    _run(test_empty_mailbox_list_fails_fast)
    _run(test_secret_in_config_file_is_rejected)
    _run(test_legacy_azure_block_is_rejected)
    _run(test_initial_cursor_defaults_to_none)
    _run(test_initial_cursor_parses_iso_datetime)
    _run(test_storage_backends_default_to_none_active)
    _run(test_storage_backends_loaded_from_yaml)
    _run(test_storage_backends_loaded_from_env)
    _run(test_unknown_storage_backend_fails_at_config_load)
    _run(test_repeated_storage_backend_fails_at_config_load)
    _run(test_unknown_transport_fails_fast)
    _run(test_missing_config_file_fails_fast)
    print("All config tests passed.")
