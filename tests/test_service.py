"""Tests for the service skeleton: bus transport factory + lifecycle.

The skeleton has no email logic yet — it connects to the bus eagerly (fail
fast), idles in a cancellable loop, and drains cleanly on request.

Runnable standalone (`python -m tests.test_service`) or under pytest.
"""

import asyncio

from eggai import InMemoryTransport, KafkaTransport
from eggai.transport.base import Transport

from bus import build_transport
from config import BusConfig
from service import Service


# --- transport factory -----------------------------------------------------


def test_build_inmemory_transport():
    t = build_transport(BusConfig(transport="inmemory"))
    assert isinstance(t, InMemoryTransport)


def test_build_kafka_transport_uses_broker_url():
    t = build_transport(BusConfig(transport="kafka", broker_url="broker:19092"))
    assert isinstance(t, KafkaTransport)


def test_build_redis_transport():
    from eggai import RedisTransport

    t = build_transport(BusConfig(transport="redis", broker_url="redis://x:6379"))
    assert isinstance(t, RedisTransport)


# --- lifecycle -------------------------------------------------------------


class _FailingTransport(Transport):
    """A transport whose connect() always fails — to exercise fail-fast."""

    async def connect(self):
        raise ConnectionError("broker unreachable")

    async def disconnect(self):
        pass

    async def publish(self, channel, message):
        pass

    async def subscribe(self, channel, callback, **kwargs):
        pass


def _service(transport, **kw):
    return Service(
        mailboxes=["a@egg-ai.com"],
        channel="emails",
        poll_interval_seconds=kw.pop("poll_interval_seconds", 60.0),
        transport=transport,
        **kw,
    )


def test_start_connects_the_bus():
    async def go():
        t = InMemoryTransport()
        svc = _service(t)
        await svc.start()
        assert t._connected is True
        await svc.stop()

    asyncio.run(go())


def test_start_fails_fast_when_bus_unreachable():
    async def go():
        svc = _service(_FailingTransport())
        raised = False
        try:
            await svc.start()
        except ConnectionError:
            raised = True
        assert raised, "start() must propagate the connection failure"

    asyncio.run(go())


def test_stop_disconnects_the_bus():
    async def go():
        t = InMemoryTransport()
        svc = _service(t)
        await svc.start()
        await svc.stop()
        assert t._connected is False

    asyncio.run(go())


def test_run_loop_invokes_cycle_each_iteration():
    async def go():
        cycles = 0

        async def cycle():
            nonlocal cycles
            cycles += 1
            if cycles >= 3:
                svc.request_stop()

        svc = _service(
            InMemoryTransport(), poll_interval_seconds=0.001, cycle=cycle
        )
        await svc.start()
        await svc.run()
        await svc.stop()
        assert cycles == 3

    asyncio.run(go())


def test_request_stop_cancels_the_sleep_promptly():
    async def go():
        # A long interval: if the sleep were not cancellable, run() would block
        # for 100s. request_stop() must wake it so run() returns near-instantly.
        svc = _service(InMemoryTransport(), poll_interval_seconds=100.0)
        await svc.start()
        runner = asyncio.create_task(svc.run())
        await asyncio.sleep(0.05)  # let it reach the sleep
        svc.request_stop()
        await asyncio.wait_for(runner, timeout=2.0)  # must not wait 100s
        await svc.stop()

    asyncio.run(go())


def test_in_flight_cycle_finishes_before_drain():
    async def go():
        started = asyncio.Event()
        finished = False

        async def slow_cycle():
            nonlocal finished
            started.set()
            await asyncio.sleep(0.1)
            finished = True
            svc.request_stop()

        svc = _service(
            InMemoryTransport(), poll_interval_seconds=100.0, cycle=slow_cycle
        )
        await svc.start()
        runner = asyncio.create_task(svc.run())
        await started.wait()
        await asyncio.wait_for(runner, timeout=2.0)
        # The cycle in flight when stop was requested ran to completion.
        assert finished is True
        await svc.stop()

    asyncio.run(go())


if __name__ == "__main__":
    test_build_inmemory_transport()
    test_build_kafka_transport_uses_broker_url()
    test_build_redis_transport()
    test_start_connects_the_bus()
    test_start_fails_fast_when_bus_unreachable()
    test_stop_disconnects_the_bus()
    test_run_loop_invokes_cycle_each_iteration()
    test_request_stop_cancels_the_sleep_promptly()
    test_in_flight_cycle_finishes_before_drain()
    print("All service tests passed.")
