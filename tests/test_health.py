import asyncio
import datetime
import json
import urllib.request

from conftest import T0

from outlook_connector.health import HealthMonitor, start_health_server
from outlook_connector.poller import PollSummary


class Clock:
    """Injectable clock: starts at T0, advanced manually."""

    def __init__(self):
        self.now = T0

    def __call__(self) -> datetime.datetime:
        return self.now

    def advance(self, seconds: float) -> None:
        self.now += datetime.timedelta(seconds=seconds)


def make_monitor(poll_interval_seconds: float = 60.0) -> tuple[HealthMonitor, Clock]:
    clock = Clock()
    monitor = HealthMonitor(poll_interval_seconds=poll_interval_seconds, now=clock)
    return monitor, clock


def test_starting_before_first_cycle():
    monitor, clock = make_monitor()
    clock.advance(10)

    snapshot = monitor.snapshot()

    assert snapshot.status == "starting"
    assert snapshot.uptime_seconds == 10
    assert snapshot.last_cycle is None
    assert snapshot.graph.status == "unknown"
    assert snapshot.bus.status == "unknown"


def test_successful_cycle_marks_ok():
    monitor, clock = make_monitor()
    clock.advance(60)

    monitor.record_cycle(PollSummary(fetched=2, published=2))
    snapshot = monitor.snapshot()

    assert snapshot.status == "ok"
    assert snapshot.last_cycle_completed_at == clock.now
    assert snapshot.last_successful_cycle_at == clock.now
    assert snapshot.graph.status == "ok"
    assert snapshot.bus.status == "ok"


def test_empty_cycle_leaves_bus_unexercised():
    monitor, _ = make_monitor()

    monitor.record_cycle(PollSummary())
    snapshot = monitor.snapshot()

    assert snapshot.status == "ok"
    assert snapshot.graph.status == "ok"
    assert snapshot.bus.status == "unknown"


def test_graph_error_is_attributed_to_graph():
    monitor, clock = make_monitor()

    monitor.record_cycle(
        PollSummary(
            error="GraphError: [503] server said no to inbox@example.com",
            error_class="GraphError",
            error_status=503,
            error_source="graph",
        )
    )
    snapshot = monitor.snapshot()

    assert snapshot.status == "degraded"
    assert snapshot.last_successful_cycle_at is None
    assert snapshot.graph.status == "error"
    # identity-free: class + status only, never the full error text
    assert snapshot.graph.last_error == "GraphError [503]"
    assert snapshot.last_cycle.error == "GraphError [503]"
    assert snapshot.graph.last_error_at == clock.now
    assert snapshot.bus.status == "unknown"


def test_bus_error_is_attributed_to_bus():
    monitor, _ = make_monitor()

    monitor.record_cycle(
        PollSummary(
            fetched=3,
            published=1,
            dropped=2,
            error="RuntimeError: bus down at redis://internal-host/0",
            error_class="RuntimeError",
            error_source="bus",
        )
    )
    snapshot = monitor.snapshot()

    assert snapshot.status == "degraded"
    # the fetch worked, so Graph is fine even though the cycle errored
    assert snapshot.graph.status == "ok"
    assert snapshot.bus.status == "error"
    assert snapshot.bus.last_error == "RuntimeError"


def test_recovery_clears_degraded_but_keeps_error_history():
    monitor, clock = make_monitor()

    monitor.record_cycle(
        PollSummary(error="boom", error_class="RuntimeError", error_source="graph")
    )
    clock.advance(60)
    monitor.record_cycle(PollSummary(fetched=1, published=1))
    snapshot = monitor.snapshot()

    assert snapshot.status == "ok"
    assert snapshot.graph.status == "ok"
    assert snapshot.graph.last_error == "RuntimeError"  # kept for forensics


def test_wedged_loop_goes_stale():
    monitor, clock = make_monitor(poll_interval_seconds=60)
    monitor.record_cycle(PollSummary())

    clock.advance(179)
    assert monitor.snapshot().status == "ok"
    clock.advance(2)  # past 3 * poll_interval
    assert monitor.snapshot().status == "stale"


def test_never_completing_first_cycle_goes_stale():
    monitor, clock = make_monitor(poll_interval_seconds=60)

    clock.advance(181)

    assert monitor.snapshot().status == "stale"


def test_long_backfill_with_beats_is_not_stale():
    """An unbounded first backfill can outlast the staleness window without a
    completed cycle; per-message beats must keep the probe green, or a
    liveness kill would restart the drain (and the in-memory cursor) forever."""
    monitor, clock = make_monitor(poll_interval_seconds=60)

    for _ in range(10):  # 10 minutes of publishing, one beat a minute
        clock.advance(60)
        monitor.beat()

    assert monitor.snapshot().status == "starting"  # alive, first cycle still running

    clock.advance(181)  # beats stop -> genuinely wedged
    assert monitor.snapshot().status == "stale"


async def test_http_endpoint_serves_snapshot():
    from aiohttp.test_utils import TestClient, TestServer

    from outlook_connector.health import build_app

    monitor, clock = make_monitor()
    async with TestClient(TestServer(build_app(monitor))) as client:
        monitor.record_cycle(PollSummary(fetched=1, published=1))

        response = await client.get("/health")
        assert response.status == 200
        payload = json.loads(await response.text())
        assert payload["status"] == "ok"
        assert payload["last_cycle"]["published"] == 1
        # unauthenticated endpoint: no identity in the payload
        assert "mailbox" not in payload
        assert "source_folder" not in payload

        response = await client.get("/nope")
        assert response.status == 404

        clock.advance(3600)  # wedge the loop
        response = await client.get("/health")
        assert response.status == 503
        assert json.loads(await response.text())["status"] == "stale"


async def test_start_health_server_binds_and_cleans_up():
    monitor, _ = make_monitor()
    runner = await start_health_server(monitor, port=0)  # ephemeral port
    try:
        site = next(iter(runner.sites))
        port = site._server.sockets[0].getsockname()[1]  # noqa: SLF001
        response = await asyncio.to_thread(
            urllib.request.urlopen, f"http://localhost:{port}/health", None, 5
        )
        with response:
            assert response.status == 200
    finally:
        await runner.cleanup()
