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


async def run_service(settings, transport: Transport | None = None, cycle: Cycle | None = None):
    """Build, run, and drain the service from validated ``settings``."""
    from bus import build_transport

    transport = transport or build_transport(settings.bus)
    service = Service(
        mailboxes=settings.mailboxes,
        channel=settings.bus.channel,
        poll_interval_seconds=settings.poll_interval_seconds,
        transport=transport,
        cycle=cycle,
    )
    _install_signal_handlers(asyncio.get_running_loop(), service.request_stop)
    await service.start()
    try:
        await service.run()
    finally:
        await service.stop()
    return service
