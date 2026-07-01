"""Service skeleton: bus lifecycle + idle loop (no email logic yet).

Per `the spec <docs/implementation.md#startup--shutdown>`:

- **Startup — fail fast.** :meth:`Service.start` connects to the bus eagerly; a
  failure propagates so the entrypoint can exit non-zero.
- **Idle loop.** :meth:`Service.run` runs one ``cycle`` per iteration (a no-op
  in the skeleton — Piece 3 plugs polling in here) then sleeps the configured
  interval. The sleep is *cancellable*: a stop request wakes it immediately
  while letting any in-flight cycle finish (fixed-delay, never overlapping).
- **Shutdown — graceful drain.** :func:`run_service` wires ``SIGTERM``/``SIGINT``
  to :meth:`Service.request_stop`; the loop stops starting new cycles and
  :meth:`Service.stop` closes the bus connection.
"""

import asyncio
import logging
import signal
from collections.abc import Awaitable, Callable
from datetime import datetime, timezone

from eggai import Channel
from eggai.transport.base import Transport

logger = logging.getLogger(__name__)

# A unit of per-cycle work. The skeleton uses a no-op; Piece 3 supplies polling.
Cycle = Callable[[], Awaitable[None]]


async def _noop() -> None:
    return None


class Service:
    """Owns the bus connection and the fixed-delay idle loop."""

    def __init__(
        self,
        *,
        mailboxes: list[str],
        channel: str,
        poll_interval_seconds: float,
        transport: Transport,
        cycle: Cycle | None = None,
    ):
        self.mailboxes = mailboxes
        self.poll_interval_seconds = poll_interval_seconds
        self._transport = transport
        self.channel = Channel(channel, transport=transport)
        self._cycle = cycle or _noop
        self._stop = asyncio.Event()

    async def start(self) -> None:
        """Connect to the bus eagerly. Any failure propagates (fail fast)."""
        await self._transport.connect()
        logger.info(
            "Connected to bus; polling %d mailbox(es) every %ss",
            len(self.mailboxes),
            self.poll_interval_seconds,
        )

    async def run(self) -> None:
        """Run cycles until stop is requested, sleeping the interval between."""
        while not self._stop.is_set():
            await self._cycle()
            await self._sleep_or_stop(self.poll_interval_seconds)
        logger.info("Loop stopped")

    def request_stop(self) -> None:
        """Signal the loop to stop after the current cycle (wakes the sleep)."""
        self._stop.set()

    async def stop(self) -> None:
        """Close the bus connection."""
        self.request_stop()
        try:
            await self._transport.disconnect()
        except Exception:  # pragma: no cover - best-effort shutdown
            logger.exception("Error closing bus connection")
        logger.info("Bus connection closed")

    async def _sleep_or_stop(self, seconds: float) -> None:
        """Sleep ``seconds``, returning early if stop is requested."""
        try:
            await asyncio.wait_for(self._stop.wait(), timeout=seconds)
        except asyncio.TimeoutError:
            pass


def _install_signal_handlers(loop: asyncio.AbstractEventLoop, handler) -> None:
    for sig in (signal.SIGTERM, signal.SIGINT):
        try:
            loop.add_signal_handler(sig, handler)
        except NotImplementedError:  # pragma: no cover - non-Unix fallback
            signal.signal(sig, lambda *_: handler())


def _build_clients(settings) -> dict:
    """One app-only :class:`OutlookClient` per configured mailbox.

    The helper client is bound to a single mailbox, so both the inbound poller
    and the outbound listener need one per mailbox; this shared seam builds them
    from the validated Azure credentials.
    """
    from outlook_helper import AppOnlyConfig, ClientSecretCredential, OutlookClient

    credential = ClientSecretCredential(
        AppOnlyConfig(
            client_id=settings.azure.client_id,
            tenant_id=settings.azure.tenant_id,
            client_secret=settings.client_secret,
        )
    )
    return {
        mailbox: OutlookClient(credential, mailbox=mailbox)
        for mailbox in settings.mailboxes
    }


def build_poller(settings, transport: Transport):
    """Wire the real :class:`~poller.Poller` from validated ``settings``.

    Builds one app-only :class:`OutlookClient` per mailbox, seeds each cursor to
    ``initial_cursor`` or process-start "now", and publishes to the bus channel
    on ``transport``. Every (synchronous) helper call runs via
    :func:`asyncio.to_thread` so a blocking request or a ``Retry-After`` sleep
    never stalls the event loop.
    """
    from poller import Poller

    clients = _build_clients(settings)

    async def fetch(mailbox: str, cursor: datetime):
        client = clients[mailbox]
        return await asyncio.to_thread(
            lambda: list(
                client.search_email(
                    since_exclusive=cursor,  # strict >, sub-second precise
                    include_headers=True,  # internet_message_id + In-Reply-To/References
                    html_body=True,  # body guaranteed HTML
                )
            )
        )

    async def fetch_attachments(mailbox: str, graph_id: str):
        # The extra Graph call, made only for attachment-bearing mail; fetches
        # metadata + content. The helper is synchronous so it runs off the event
        # loop like every other call.
        client = clients[mailbox]
        return await asyncio.to_thread(client.get_attachments, graph_id)

    channel = Channel(settings.bus.channel, transport=transport)
    seed = settings.initial_cursor or datetime.now(timezone.utc)
    cursors = {mailbox: seed for mailbox in settings.mailboxes}

    return Poller(
        mailboxes=list(settings.mailboxes),
        fetch=fetch,
        publish=channel.publish,
        cursors=cursors,
        fetch_attachments=fetch_attachments,
    )


def build_send_listener(settings, transport: Transport):
    """Wire the real :class:`~listener.Agent` from validated ``settings``.

    Builds one app-only :class:`OutlookClient` per mailbox and supplies a ``send``
    collaborator that maps an :class:`~schemas.EmailSend` onto the helper: a fresh
    ``send_email``, or ``reply`` when ``reply_to_graph_id`` is set (threaded
    delivery). The send-from ``mailbox`` is validated against the configured set —
    an unknown mailbox raises rather than silently sending from a default. Like
    the poller, every (synchronous) helper call runs via :func:`asyncio.to_thread`.
    """
    from listener import make_send_listener
    from schemas import EmailSend

    clients = _build_clients(settings)

    async def send(request: EmailSend) -> None:
        client = clients.get(request.mailbox)
        if client is None:
            raise ValueError(f"Unknown send mailbox: {request.mailbox!r}")
        html = request.body_content_type != "text"
        if request.reply_to_graph_id:
            await asyncio.to_thread(
                lambda: client.reply(
                    request.reply_to_graph_id, request.body, html=html
                )
            )
            return
        to = [a.address for a in request.to]
        cc = [a.address for a in request.cc] or None
        bcc = [a.address for a in request.bcc] or None
        await asyncio.to_thread(
            lambda: client.send_email(
                to, request.subject, request.body, cc=cc, bcc=bcc, html=html
            )
        )

    return make_send_listener(
        channel=settings.bus.channel, transport=transport, send=send
    )


async def run_service(
    settings,
    transport: Transport | None = None,
    cycle: Cycle | None = None,
    listener=None,
):
    """Build, run, and drain the service from validated ``settings``.

    Runs both directions on the one shared bus connection: the inbound poll loop
    (``cycle``) and the outbound ``email.send`` ``listener``. ``cycle``/
    ``listener`` are injectable for tests; either ``None`` builds the real one.

    Ordering matters: the listener is started **first**. ``Agent.start`` registers
    its subscription and *then* connects the transport, and a transport like Kafka
    only begins consuming for subscribers registered **before** the broker is
    started — so subscribing must happen before any other code (the Service, the
    poller's publishing channel) connects the shared transport. On shutdown the
    listener is drained first so no new send is accepted once we begin stopping.
    """
    from bus import build_transport

    transport = transport or build_transport(settings.bus)
    if cycle is None:
        cycle = build_poller(settings, transport).run_cycle
    if listener is None:
        listener = build_send_listener(settings, transport)
    service = Service(
        mailboxes=settings.mailboxes,
        channel=settings.bus.channel,
        poll_interval_seconds=settings.poll_interval_seconds,
        transport=transport,
        cycle=cycle,
    )
    _install_signal_handlers(asyncio.get_running_loop(), service.request_stop)
    # Subscribe-then-connect: register the email.send consumer before the shared
    # transport is connected, or Kafka never starts consuming for it.
    await listener.start()
    await service.start()
    try:
        await service.run()
    finally:
        await listener.stop()
        await service.stop()
    return service
