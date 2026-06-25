"""EggAI bus transport factory.

Maps the structural :class:`~config.BusConfig` onto a concrete eggai transport.
A ``broker_url`` of ``None`` leaves the transport's own default in place.
"""

from eggai import InMemoryTransport, KafkaTransport, RedisTransport
from eggai.transport.base import Transport

from config import BusConfig


def build_transport(bus: BusConfig) -> Transport:
    """Construct the transport described by ``bus``."""
    if bus.transport == "kafka":
        if bus.broker_url:
            return KafkaTransport(bootstrap_servers=bus.broker_url)
        return KafkaTransport()
    if bus.transport == "redis":
        if bus.broker_url:
            return RedisTransport(url=bus.broker_url)
        return RedisTransport()
    if bus.transport == "inmemory":
        return InMemoryTransport()
    raise ValueError(f"Unknown transport: {bus.transport!r}")
