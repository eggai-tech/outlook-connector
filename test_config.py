"""Tests for the layered service configuration.

Structural config comes from a YAML file; the ``client_secret`` is supplied via
an environment variable only and must never appear in the YAML.

Runnable standalone (`python test_config.py`) or under pytest.
"""

import os
from pathlib import Path

import pytest
from pydantic import ValidationError

from config import SecretInConfigError, Settings, load_settings

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
azure:
  tenant_id: tenant-123
  client_id: client-abc
"""


def _write(tmp_path: Path, text: str) -> Path:
    cfg = tmp_path / "config.yaml"
    cfg.write_text(text)
    return cfg


def _load(tmp_path: Path, text: str, monkeypatch, *, secret: str | None = "s3cret"):
    cfg = _write(tmp_path, text)
    monkeypatch.setenv("CONFIG_FILE", str(cfg))
    if secret is None:
        monkeypatch.delenv("client_secret", raising=False)
    else:
        monkeypatch.setenv("client_secret", secret)
    return load_settings()


def test_loads_structural_config_from_yaml(tmp_path, monkeypatch):
    settings = _load(tmp_path, _VALID_YAML, monkeypatch)

    assert settings.mailboxes == ["invoices@egg-ai.com", "support@egg-ai.com"]
    assert settings.poll_interval_seconds == 30
    assert settings.log_level == "DEBUG"
    assert settings.bus.transport == "kafka"
    assert settings.bus.broker_url == "broker:19092"
    assert settings.bus.channel == "emails"
    assert settings.azure.tenant_id == "tenant-123"
    assert settings.azure.client_id == "client-abc"


def test_client_secret_comes_from_env(tmp_path, monkeypatch):
    settings = _load(tmp_path, _VALID_YAML, monkeypatch, secret="top-secret")
    assert settings.client_secret == "top-secret"


def test_defaults_applied(tmp_path, monkeypatch):
    minimal = """\
mailboxes:
  - invoices@egg-ai.com
bus:
  broker_url: broker:19092
azure:
  tenant_id: t
  client_id: c
"""
    settings = _load(tmp_path, minimal, monkeypatch)
    assert settings.poll_interval_seconds == 60.0
    assert settings.log_level == "INFO"
    assert settings.bus.transport == "kafka"
    assert settings.bus.channel == "emails"


def test_missing_client_secret_fails_fast(tmp_path, monkeypatch):
    with pytest.raises(ValidationError):
        _load(tmp_path, _VALID_YAML, monkeypatch, secret=None)


def test_missing_required_structural_field_fails_fast(tmp_path, monkeypatch):
    no_azure = """\
mailboxes:
  - invoices@egg-ai.com
bus:
  broker_url: broker:19092
"""
    with pytest.raises(ValidationError):
        _load(tmp_path, no_azure, monkeypatch)


def test_empty_mailbox_list_fails_fast(tmp_path, monkeypatch):
    no_mailboxes = """\
mailboxes: []
bus:
  broker_url: broker:19092
azure:
  tenant_id: t
  client_id: c
"""
    with pytest.raises(ValidationError):
        _load(tmp_path, no_mailboxes, monkeypatch)


def test_secret_in_yaml_is_rejected(tmp_path, monkeypatch):
    with_secret = _VALID_YAML + "client_secret: leaked-in-file\n"
    with pytest.raises(SecretInConfigError):
        _load(tmp_path, with_secret, monkeypatch)


def test_unknown_transport_fails_fast(tmp_path, monkeypatch):
    bad = _VALID_YAML.replace("transport: kafka", "transport: carrier-pigeon")
    with pytest.raises(ValidationError):
        _load(tmp_path, bad, monkeypatch)


def test_missing_config_file_fails_fast(tmp_path, monkeypatch):
    monkeypatch.setenv("CONFIG_FILE", str(tmp_path / "does-not-exist.yaml"))
    monkeypatch.setenv("client_secret", "s3cret")
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
    _run(test_client_secret_comes_from_env)
    _run(test_defaults_applied)
    _run(test_missing_client_secret_fails_fast)
    _run(test_missing_required_structural_field_fails_fast)
    _run(test_empty_mailbox_list_fails_fast)
    _run(test_secret_in_yaml_is_rejected)
    _run(test_unknown_transport_fails_fast)
    _run(test_missing_config_file_fails_fast)
    print("All config tests passed.")
