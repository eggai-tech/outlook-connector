"""Outlook Connector entrypoint"""

import asyncio

import structlog
from pydantic import ValidationError

from outlook_connector.config import get_settings
from outlook_connector.service import run_service

logger = structlog.get_logger()


async def main() -> int:
    logger.debug("OUTLOOK-CONNECTOR")

    logger.debug("Loading settings...")
    try:
        get_settings()
    except ValidationError:
        logger.exception("Configuration error")
        raise

    logger.debug("Running service...")
    await run_service()


if __name__ == "__main__":
    logger.debug("MAIN")
    asyncio.run(main())
