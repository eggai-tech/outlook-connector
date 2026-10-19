"""Outlook Connector entrypoint"""

import asyncio
import logging

import structlog
from pydantic import ValidationError

from outlook_connector.config import get_settings
from outlook_connector.service import run_service

logger = structlog.get_logger()


def configure_logging(level: str) -> None:
    """Drop events below ``level`` (``log_level`` in the settings).

    Without this structlog's default logger prints every level, DEBUG included,
    whatever the setting says.
    """
    levels = logging.getLevelNamesMapping()
    name = level.upper()
    if name not in levels:
        raise ValueError(f"log_level must be one of DEBUG, INFO, WARNING, ERROR; got {level!r}")
    structlog.configure(wrapper_class=structlog.make_filtering_bound_logger(levels[name]))


async def _run() -> None:
    logger.debug("OUTLOOK-CONNECTOR")

    logger.debug("Loading settings...")
    try:
        settings = get_settings()
    except ValidationError:
        logger.exception("Configuration error")
        raise
    configure_logging(settings.log_level)

    logger.debug("Running service...")
    await run_service()


def main() -> None:
    """Sync console-script entrypoint (the Docker ENTRYPOINT)."""
    asyncio.run(_run())


if __name__ == "__main__":
    main()
