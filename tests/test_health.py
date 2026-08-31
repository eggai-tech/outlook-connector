import datetime
import json
import urllib.error
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
    monitor = HealthMonitor(
        mailbox="inbox@example.com",
        source_folder="inbox",
        poll_interval_seconds=poll_interval_seconds,
        now=clock,
    )
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

    monitor.record_cycle(PollSummary(error="ConnectError: boom", error_source="graph"))
    snapshot = monitor.snapshot()

    assert snapshot.status == "degraded"
    assert snapshot.last_successful_cycle_at is None
    assert snapshot.graph.status == "error"
    assert snapshot.graph.last_error == "ConnectError: boom"
    assert snapshot.graph.last_error_at == clock.now
    assert snapshot.bus.status == "unknown"


def test_bus_error_is_attributed_to_bus():
    monitor, _ = make_monitor()

    monitor.record_cycle(
        PollSummary(
            fetched=3, published=1, dropped=2, error="bus down", error_source="bus"
        )
    )
    snapshot = monitor.snapshot()

    assert snapshot.status == "degraded"
    # the fetch worked, so Graph is fine even though the cycle errored
    assert snapshot.graph.status == "ok"
    assert snapshot.bus.status == "error"
    assert snapshot.bus.last_error == "bus down"


def test_recovery_clears_degraded_but_keeps_error_history():
    monitor, clock = make_monitor()

    monitor.record_cycle(PollSummary(error="boom", error_source="graph"))
    clock.advance(60)
    monitor.record_cycle(PollSummary(fetched=1, published=1))
    snapshot = monitor.snapshot()

    assert snapshot.status == "ok"
    assert snapshot.graph.status == "ok"
    assert snapshot.graph.last_error == "boom"  # kept for forensics


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


def _get(port: int, path: str):
    return urllib.request.urlopen(f"http://localhost:{port}{path}", timeout=5)


def test_http_endpoint_serves_snapshot():
    monitor, clock = make_monitor()
    server = start_health_server(monitor, port=0)  # ephemeral port
    try:
        monitor.record_cycle(PollSummary(fetched=1, published=1))

        with _get(server.port, "/health") as response:
            assert response.status == 200
            payload = json.loads(response.read())
        assert payload["status"] == "ok"
        assert payload["mailbox"] == "inbox@example.com"
        assert payload["last_cycle"]["published"] == 1

        try:
            _get(server.port, "/nope")
            raise AssertionError("expected 404")
        except urllib.error.HTTPError as exc:
            assert exc.code == 404

        clock.advance(3600)  # wedge the loop
        try:
            _get(server.port, "/health")
            raise AssertionError("expected 503")
        except urllib.error.HTTPError as exc:
            assert exc.code == 503
            assert json.loads(exc.read())["status"] == "stale"
    finally:
        server.shutdown()
        server.server_close()
