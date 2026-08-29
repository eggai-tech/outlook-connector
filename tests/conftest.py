"""Shared test setup.

Keeps the suite hermetic: :class:`~outlook_connector.config.Settings` always
consults ``CONFIG_FILE`` (default ``config.yaml`` in the working directory), so
without this a developer's own config file would leak into every test that
constructs settings directly. Pointing it at a path that does not exist makes
the YAML source contribute an empty document; tests that care about file
contents set ``CONFIG_FILE`` themselves and override this.
"""

import pytest


@pytest.fixture(autouse=True)
def _no_ambient_config_file(tmp_path, monkeypatch):
    monkeypatch.setenv("CONFIG_FILE", str(tmp_path / "no-such-config.yaml"))
