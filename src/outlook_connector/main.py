"""Outlook connector entrypoint.

Loads and validates configuration (fail fast → exit non-zero on any error),
then runs the long-running service until ``SIGTERM``/``SIGINT`` drains it
cleanly (exit 0). See `the spec <docs/DESIGN.md#startup--shutdown>`.
"""

import asyncio
import logging
import sys

from pydantic import ValidationError

from outlook_connector.config import load_settings
from outlook_connector.service import run_service

logger = logging.getLogger("outlook_connector")


def _configure_logging(level: str = "INFO") -> None:
    logging.basicConfig(
        level=level,
        stream=sys.stdout,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )


def _describe(exc: ValidationError) -> str:
    """One line per invalid field — field name and reason, never the value."""
    return "; ".join(
        ".".join(str(part) for part in error["loc"]) + f": {error['msg']}"
        for error in exc.errors(include_input=False, include_url=False)
    )


def main() -> int:
    _configure_logging()

    try:
        settings = load_settings()
    except ValidationError as exc:
        # Configuration is user error — a clear message, no traceback. Rendered
        # without input values: a missing-field error reports the whole merged
        # input, which carries AZURE_CLIENT_SECRET.
        logger.error("Configuration error: %s", _describe(exc))
        return 1
    except Exception as exc:
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
