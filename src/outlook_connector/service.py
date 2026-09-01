"""
Main service workflow
"""

import asyncio
import datetime

import structlog
from eggai import Channel

from outlook_connector.bus import build_bus_event, build_transport
from outlook_connector.config import get_settings
from outlook_connector.health import HealthMonitor, start_health_server
from outlook_connector.poller import GRAPH_ERRORS, Poller, PollSummary
from outlook_connector.storage import save_to_storage

logger = structlog.get_logger()


def _record_error(summary: PollSummary, exc: Exception, source: str) -> None:
    summary.error = f"{type(exc).__name__}: {exc}"
    summary.error_class = type(exc).__name__
    # GraphError carries the HTTP status; transport errors don't.
    status = getattr(exc, "status_code", None)
    summary.error_status = status if isinstance(status, int) else None
    summary.error_source = source


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
    # Intra-cycle liveness for the health monitor: a long backfill cycle must
    # not read as a wedged loop while it is visibly making progress.
    heartbeat = context.get("heartbeat") or (lambda: None)
    summary = PollSummary()

    try:
        messages = await asyncio.to_thread(poller.poll_mailbox)
    except GRAPH_ERRORS as exc:
        _record_error(summary, exc, "graph")
        logger.warning("Poll fetch failed", error=summary.error)
        return summary

    heartbeat()
    summary.fetched = len(messages)
    fetched_at = poller.now()

    for index, message in enumerate(messages):
        try:
            await process_message(context, message, fetched_at=fetched_at)
            heartbeat()
        except Exception as exc:
            # attachment fetch raises a Graph error; anything else is the bus
            _record_error(summary, exc, "graph" if isinstance(exc, GRAPH_ERRORS) else "bus")
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
    poller = build_poller()
    transport = build_transport(settings.bus)
    channel = Channel(settings.bus.channel, transport=transport)

    context = {
        "poller": poller,
        "channel": channel,
        "source_mailbox": settings.mailbox,
    }

    monitor = HealthMonitor(poll_interval_seconds=settings.poll_interval_seconds)
    context["heartbeat"] = monitor.beat
    runner = None
    if settings.health_port is not None:
        runner = await start_health_server(monitor, settings.health_port)

    try:
        while True:
            logger.debug("Tick.", poll_interval_seconds=settings.poll_interval_seconds)

            summary = await run_workflow(context)
            monitor.record_cycle(summary)
            await asyncio.sleep(settings.poll_interval_seconds)
    finally:
        if runner is not None:
            await runner.cleanup()
