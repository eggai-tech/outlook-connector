"""Health/status HTTP endpoint.

The poll loop reports each cycle to a :class:`HealthMonitor`; a stdlib HTTP
server (daemon thread, no extra dependencies) serves the current snapshot as
JSON on ``GET /health``.

Status codes are chosen for orchestrator probes: the endpoint answers 200 as
long as the poll loop is alive and ticking — even while Graph or the bus is
erroring, because restarting the connector cannot fix an external outage (the
body still reports the errors) — and 503 only once the loop itself is wedged:
no completed cycle within the staleness window.

The endpoint is unauthenticated, so the payload deliberately carries no
identity: no mailbox address, no folder name, and errors reduced to exception
class + Graph HTTP status (full error text — which can embed mailbox addresses
and URLs — stays in the logs). Expose the port to internal networks only.
"""

import datetime
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Literal

import structlog
from pydantic import BaseModel

from outlook_connector.poller import PollSummary

logger = structlog.get_logger()


def _utcnow() -> datetime.datetime:
    return datetime.datetime.now(datetime.UTC)


class ProbeStatus(BaseModel):
    """Last observed outcome for one dependency (Graph or the bus).

    ``unknown`` means the dependency has not been exercised yet — e.g. the
    bus before the first cycle that actually had something to publish.
    """

    status: Literal["ok", "error", "unknown"] = "unknown"
    last_success_at: datetime.datetime | None = None
    last_error: str | None = None
    last_error_at: datetime.datetime | None = None

    def succeeded(self, at: datetime.datetime) -> "ProbeStatus":
        return self.model_copy(update={"status": "ok", "last_success_at": at})

    def failed(self, error: str | None, at: datetime.datetime) -> "ProbeStatus":
        return self.model_copy(
            update={"status": "error", "last_error": error, "last_error_at": at}
        )


class CycleStats(BaseModel):
    """Identity-free view of a :class:`PollSummary` for the public payload."""

    fetched: int
    published: int
    dropped: int
    error: str | None  # exception class + Graph status, never the full text
    error_source: Literal["graph", "bus"] | None


def _public_error(summary: PollSummary) -> str | None:
    if summary.error_class is None:
        return None
    if summary.error_status is not None:
        return f"{summary.error_class} [{summary.error_status}]"
    return summary.error_class


def _cycle_stats(summary: PollSummary) -> CycleStats:
    return CycleStats(
        fetched=summary.fetched,
        published=summary.published,
        dropped=summary.dropped,
        error=_public_error(summary),
        error_source=summary.error_source,
    )


class HealthSnapshot(BaseModel):
    """The ``GET /health`` response body."""

    status: Literal["starting", "ok", "degraded", "stale"]
    started_at: datetime.datetime
    uptime_seconds: float
    poll_interval_seconds: float
    last_cycle_completed_at: datetime.datetime | None
    last_successful_cycle_at: datetime.datetime | None
    last_cycle: CycleStats | None
    graph: ProbeStatus
    bus: ProbeStatus


class HealthMonitor:
    """Aggregates poll-cycle outcomes into a point-in-time health snapshot.

    ``record_cycle`` runs on the event loop thread, ``snapshot`` on HTTP
    server threads — hence the lock.
    """

    def __init__(
        self,
        *,
        poll_interval_seconds: float,
        now=_utcnow,
    ):
        self._now = now
        self._lock = threading.Lock()
        self._poll_interval_seconds = poll_interval_seconds
        self._started_at = now()
        self._last_cycle: PollSummary | None = None
        self._last_cycle_at: datetime.datetime | None = None
        self._last_success_at: datetime.datetime | None = None
        self._graph = ProbeStatus()
        self._bus = ProbeStatus()
        # A loop that misses this much wall-clock (a few intervals, with slack
        # for one slow in-flight cycle) is reported stale -> HTTP 503.
        self._stale_after = max(3 * poll_interval_seconds, poll_interval_seconds + 60)

    def record_cycle(self, summary: PollSummary) -> None:
        at = self._now()
        with self._lock:
            self._last_cycle = summary
            self._last_cycle_at = at
            if summary.error is None:
                self._last_success_at = at
            # The fetch either failed (error_source == "graph" with nothing
            # fetched) or succeeded; mid-batch attachment failures also count
            # against Graph.
            if summary.error_source == "graph":
                self._graph = self._graph.failed(_public_error(summary), at)
            else:
                self._graph = self._graph.succeeded(at)
            # The bus is only exercised when there was something to publish.
            if summary.error_source == "bus":
                self._bus = self._bus.failed(_public_error(summary), at)
            elif summary.published > 0:
                self._bus = self._bus.succeeded(at)

    def snapshot(self) -> HealthSnapshot:
        now = self._now()
        with self._lock:
            last_beat = self._last_cycle_at or self._started_at
            if (now - last_beat).total_seconds() > self._stale_after:
                status = "stale"
            elif self._last_cycle is None:
                status = "starting"
            elif self._last_cycle.error is None:
                status = "ok"
            else:
                status = "degraded"
            return HealthSnapshot(
                status=status,
                started_at=self._started_at,
                uptime_seconds=(now - self._started_at).total_seconds(),
                poll_interval_seconds=self._poll_interval_seconds,
                last_cycle_completed_at=self._last_cycle_at,
                last_successful_cycle_at=self._last_success_at,
                last_cycle=(
                    _cycle_stats(self._last_cycle) if self._last_cycle is not None else None
                ),
                graph=self._graph,
                bus=self._bus,
            )


class _HealthHandler(BaseHTTPRequestHandler):
    server: "HealthServer"

    def do_GET(self) -> None:
        if self.path.split("?", 1)[0] not in ("/health", "/"):
            self.send_error(404)
            return
        snapshot = self.server.monitor.snapshot()
        body = snapshot.model_dump_json(indent=2).encode() + b"\n"
        self.send_response(503 if snapshot.status == "stale" else 200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args) -> None:  # noqa: A002 (stdlib signature)
        logger.debug("Health request", detail=format % args)


class HealthServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, monitor: HealthMonitor, port: int):
        super().__init__(("", port), _HealthHandler)
        self.monitor = monitor

    @property
    def port(self) -> int:
        return self.server_address[1]


def start_health_server(monitor: HealthMonitor, port: int) -> HealthServer:
    """Serve ``GET /health`` on all interfaces in a daemon thread."""
    server = HealthServer(monitor, port)
    threading.Thread(
        target=server.serve_forever, name="health-http", daemon=True
    ).start()
    logger.info("Health endpoint listening", port=server.port)
    return server
