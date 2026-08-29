"""Tests for the entrypoint's fail-fast exit codes.

`Done when <development/steps/02-service-skeleton.md>`: missing/invalid config
exits non-zero with a clear message.

Runnable standalone (`python -m tests.test_main`) or under pytest.
"""

from pathlib import Path

from outlook_connector.main import main

_VALID_YAML = """\
mailboxes:
  - invoices@egg-ai.com
bus:
  transport: inmemory
"""


def _set_azure_env(monkeypatch) -> None:
    monkeypatch.setenv("AZURE_TENANT_ID", "t")
    monkeypatch.setenv("AZURE_CLIENT_ID", "c")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "s3cret")


def test_exits_nonzero_when_config_file_missing(tmp_path, monkeypatch):
    monkeypatch.setenv("CONFIG_FILE", str(tmp_path / "absent.yaml"))
    _set_azure_env(monkeypatch)
    assert main() == 1


def test_exits_nonzero_when_client_secret_missing(tmp_path, monkeypatch):
    cfg = tmp_path / "config.yaml"
    cfg.write_text(_VALID_YAML)
    monkeypatch.setenv("CONFIG_FILE", str(cfg))
    _set_azure_env(monkeypatch)
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)
    assert main() == 1


def test_config_error_never_logs_the_secret(tmp_path, monkeypatch, caplog):
    """A missing field must not drag the secret into the log.

    Pydantic reports the whole merged input for a "field required" error, so the
    entrypoint renders validation errors without their input values.
    """
    cfg = tmp_path / "config.yaml"
    cfg.write_text(_VALID_YAML)
    monkeypatch.setenv("CONFIG_FILE", str(cfg))
    _set_azure_env(monkeypatch)
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "sup3r-s3cret-value")
    monkeypatch.delenv("AZURE_TENANT_ID", raising=False)

    with caplog.at_level("ERROR"):
        assert main() == 1

    assert "azure_tenant_id" in caplog.text
    assert "sup3r-s3cret-value" not in caplog.text


if __name__ == "__main__":
    import os
    import tempfile

    class _MP:
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

    # test_config_error_never_logs_the_secret needs pytest's caplog, so the
    # standalone runner covers the exit-code cases only.
    for fn in (
        test_exits_nonzero_when_config_file_missing,
        test_exits_nonzero_when_client_secret_missing,
    ):
        mp = _MP()
        with tempfile.TemporaryDirectory() as d:
            try:
                fn(Path(d), mp)
            finally:
                mp.undo()
    print("All main tests passed.")
