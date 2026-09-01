"""
Main service workflow
"""

import asyncio

import structlog
from eggai import Channel

from outlook_helper import GraphError

from outlook_connector.bus import build_bus_event, build_transport
from outlook_connector.config import get_settings
from outlook_connector.health import HealthMonitor, start_health_server
from outlook_connector.poller import GRAPH_ERRORS, Poller, PollSummary
from outlook_connector.storage import save_to_storage

logger = structlog.get_logger()


def _record_error(summary: PollSummary, exc: Exception, source: str) -> None:
    summary.error = f"{type(exc).__name__}: {exc}"
    summary.error_class = type(exc).__name__
    # Only GraphError's status is meaningful here; duck-typing status_code off
    # arbitrary bus/storage exceptions would fabricate Graph-style statuses in
    # the public health payload.
    summary.error_status = exc.status_code if isinstance(exc, GraphError) else None
    summary.error_source = source


def build_poller(heartbeat=None):
    settings = get_settings()

    kwargs = {} if heartbeat is None else {"heartbeat": heartbeat}
    return Poller(
        source_folder=settings.source_folder,
        batch_max_messages=settings.batch_max_messages,
        max_attachment_bytes=settings.max_attachment_bytes,
        ignore_received_before=settings.ignore_received_before,
        **kwargs,
    )


async def process_message(context, message, *, fetched_at):
    """
    Single message flow: fetch attachment content, wrap in the CloudEvents
    envelope, publish. Raising aborts the batch; the message stays unmarked
    and still in the folder, so the next rescan retries it.
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
    """One poll cycle: rescan the folder, publish the oldest unseen batch.

    Delivery is **at least once**: the first failure stops the batch, and the
    next rescan simply finds the unpublished mail still in the folder. A
    restart re-publishes everything still present (the seen-set is in-memory
    only) — consumers must be idempotent, deduping on ``internet_message_id``.
    """
    poller = context["poller"]
    summary = PollSummary()

    try:
        messages = await asyncio.to_thread(poller.poll_mailbox)
    except GRAPH_ERRORS as exc:
        _record_error(summary, exc, "graph")
        logger.warning("Poll fetch failed", error=summary.error)
        return summary

    summary.fetched = len(messages)
    fetched_at = poller.now()

    for index, message in enumerate(messages):
        try:
            await process_message(context, message, fetched_at=fetched_at)
            # intra-cycle liveness: a long publish drain is progress, not a wedge
            poller.heartbeat()
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
        poller.mark_published(message)
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
            "limit can carry; lower the cap or raise the broker/producer "
            "limit. Note the cap is PER attachment: a multi-attachment "
            "message can exceed the broker limit even under a smaller cap",
            max_attachment_bytes=settings.max_attachment_bytes,
        )
    monitor = HealthMonitor(poll_interval_seconds=settings.poll_interval_seconds)
    # The poller owns the heartbeat: beats fire per fetched message, through
    # retry sleeps, and per published message — so neither a long fetch nor a
    # long publish drain reads as a wedged loop.
    poller = build_poller(heartbeat=monitor.beat)
    transport = build_transport(settings.bus)
    channel = Channel(settings.bus.channel, transport=transport)

    context = {
        "poller": poller,
        "channel": channel,
        "source_mailbox": settings.mailbox,
    }

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
