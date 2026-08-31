"""
Main service workflow
"""

import asyncio
import datetime
from time import sleep

import structlog
from eggai import Channel

from outlook_connector.bus import build_bus_event, build_transport
from outlook_connector.config import get_settings
from outlook_connector.poller import Poller
from outlook_connector.storage import save_to_storage

logger = structlog.get_logger()


def build_poller():
    settings = get_settings()
    initial_cursor = settings.initial_cursor or datetime.datetime.now(datetime.UTC)

    return Poller(
        cursor=initial_cursor,
    )


async def process_message(context, message):
    """
    Single message flow
    """
    logger.debug("Processing message", message=message.id)

    # write to storage
    save_to_storage(message)

    # send on bus
    bus_event = build_bus_event(message)
    logger.debug("Publishing bus event...", new_event=bus_event)
    await context["channel"].publish(message)


async def run_workflow(context):
    messages = context["poller"].poll_mailbox()

    for message in messages:
        await process_message(context, message)


async def run_service() -> None:
    settings = get_settings()
    poller = build_poller()
    transport = build_transport(settings.bus)
    channel = Channel(settings.bus.channel, transport=transport)

    context = {
        "poller": poller,
        "channel": channel,
    }

    while True:
        logger.debug("Tick.", poll_interval_seconds=settings.poll_interval_seconds)

        await run_workflow(context)
        await asyncio.sleep(settings.poll_interval_seconds)
