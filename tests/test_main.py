import asyncio
import tomllib
from pathlib import Path

from outlook_connector import main as main_module


def test_console_entrypoint_is_sync():
    """The Docker ENTRYPOINT calls the console script directly: it must be a
    plain callable, not a coroutine function that returns un-awaited."""
    assert not asyncio.iscoroutinefunction(main_module.main)


def test_pyproject_script_points_at_sync_main():
    pyproject = Path(__file__).parents[1] / "pyproject.toml"
    scripts = tomllib.loads(pyproject.read_text())["project"]["scripts"]
    assert scripts["outlook-connector"] == "outlook_connector.main:main"
