"""Tests for the entrypoint's fail-fast exit codes.

`Done when <development/steps/02-service-skeleton.md>`: missing/invalid config
exits non-zero with a clear message.

Runnable standalone (`python -m tests.test_main`) or under pytest.
"""

from pathlib import Path

from main import main

_VALID_YAML = """\
mailboxes:
  - invoices@egg-ai.com
bus:
  transport: inmemory
azure:
  tenant_id: t
  client_id: c
"""


def test_exits_nonzero_when_config_file_missing(tmp_path, monkeypatch):
    monkeypatch.setenv("CONFIG_FILE", str(tmp_path / "absent.yaml"))
    monkeypatch.setenv("client_secret", "s3cret")
    assert main() == 1


def test_exits_nonzero_when_client_secret_missing(tmp_path, monkeypatch):
    cfg = tmp_path / "config.yaml"
    cfg.write_text(_VALID_YAML)
    monkeypatch.setenv("CONFIG_FILE", str(cfg))
    monkeypatch.delenv("client_secret", raising=False)
    assert main() == 1


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
