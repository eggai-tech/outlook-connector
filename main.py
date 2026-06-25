"""Outlook connector entrypoint.

Loads and validates configuration (fail fast → exit non-zero on any error),
then runs the long-running service until ``SIGTERM``/``SIGINT`` drains it
cleanly (exit 0). See `the spec <docs/implementation.md#startup--shutdown>`.
"""

import asyncio
import logging
import sys

from config import load_settings
from service import run_service

logger = logging.getLogger("outlook_connector")


def _configure_logging(level: str = "INFO") -> None:
    logging.basicConfig(
        level=level,
        stream=sys.stdout,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )


def main() -> int:
    _configure_logging()

    try:
        settings = load_settings()
    except Exception as exc:
        # Configuration is user error — a clear message, no traceback.
        logger.error("Configuration error: %s", exc)
        return 1

    logging.getLogger().setLevel(settings.log_level)

    try:
        asyncio.run(run_service(settings))
    except Exception:
        logger.exception("Fatal error; shutting down")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
