"""
Main service workflow
"""

import asyncio
import datetime

import structlog
from eggai import Channel

from outlook_connector.bus import build_bus_event, build_transport
from outlook_connector.config import get_settings
from outlook_connector.poller import GRAPH_ERRORS, Poller, PollSummary
from outlook_connector.storage import save_to_storage

logger = structlog.get_logger()


def build_poller():
    settings = get_settings()
    initial_cursor = settings.initial_cursor or datetime.datetime.now(datetime.UTC)

    return Poller(
        cursor=initial_cursor,
        source_folder=settings.source_folder,
        batch_max_messages=settings.batch_max_messages,
        max_attachment_bytes=settings.max_attachment_bytes,
    )


async def process_message(context, message, *, fetched_at):
    """
    Single message flow: fetch attachment content, wrap in the CloudEvents
    envelope, publish. Raising aborts the batch with the cursor untouched
    for this message, so it is retried next cycle — never duplicated.
    """
    logger.debug("Processing message", message=message.id)
    poller = context["poller"]

    attachments = await asyncio.to_thread(poller.fetch_attachments, message)

    # write to storage
    save_to_storage(message)

    # send on bus
    bus_event = build_bus_event(
        message,
        source_mailbox=context["source_mailbox"],
        fetched_at=fetched_at,
        attachments=attachments,
    )
    logger.debug("Publishing bus event...", new_event=bus_event)
    await context["channel"].publish(bus_event)


async def run_workflow(context) -> PollSummary:
    """One poll cycle: fetch, publish oldest-first, advance the cursor per
    successfully published message.

    Bias is "never duplicate, occasionally drop": the first failure stops the
    batch and leaves the cursor at the last success, so the next cycle resumes
    from there without re-publishing anything.
    """
    poller = context["poller"]
    summary = PollSummary()

    try:
        messages = await asyncio.to_thread(poller.poll_mailbox)
    except GRAPH_ERRORS as exc:
        summary.error = f"{type(exc).__name__}: {exc}"
        logger.warning("Poll fetch failed", error=summary.error)
        return summary

    summary.fetched = len(messages)
    fetched_at = poller.now()

    for index, message in enumerate(messages):
        try:
            await process_message(context, message, fetched_at=fetched_at)
        except Exception as exc:
            summary.error = f"{type(exc).__name__}: {exc}"
            summary.dropped = summary.fetched - index
            logger.warning(
                "Publish failed; stopping batch until next cycle",
                message=message.id,
                error=summary.error,
            )
            break
        poller.advance(message.received_at)
        summary.published += 1

    logger.info(
        "Poll cycle complete",
        fetched=summary.fetched,
        published=summary.published,
        dropped=summary.dropped,
        error=summary.error,
    )
    return summary


async def run_service() -> None:
    settings = get_settings()
    if settings.bus.transport == "kafka" and (
        settings.max_attachment_bytes is None or settings.max_attachment_bytes > 700_000
    ):
        # base64 inflates content 4/3, and kafka's default max_request_size is
        # 1 MiB — a larger cap means big attachments fail to publish and, under
        # the stop-batch policy, block everything behind them.
        logger.warning(
            "max_attachment_bytes exceeds what kafka's default 1MiB message "
            "limit can carry; lower the cap or raise the broker/producer limit",
            max_attachment_bytes=settings.max_attachment_bytes,
        )
    poller = build_poller()
    transport = build_transport(settings.bus)
    channel = Channel(settings.bus.channel, transport=transport)

    context = {
        "poller": poller,
        "channel": channel,
        "source_mailbox": settings.mailbox,
    }

    while True:
        logger.debug("Tick.", poll_interval_seconds=settings.poll_interval_seconds)

        await run_workflow(context)
        await asyncio.sleep(settings.poll_interval_seconds)
