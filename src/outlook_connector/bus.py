"""EggAI bus transport factory.

Maps the structural :class:`~config.BusConfig` onto a concrete eggai transport.
A ``broker_url`` of ``None`` leaves the transport's own default in place.
"""

import datetime
from collections.abc import Sequence

from eggai import InMemoryTransport, KafkaTransport, RedisTransport
from eggai.schemas import BaseMessage as EggaiBaseMessage
from eggai.transport.base import Transport
from outlook_helper.schemas import OutlookAttachment, OutlookMessage
from pydantic import BaseModel
from structlog import get_logger

from outlook_connector.config import BusConfig
from outlook_connector.mapper import outlook_message_to_email
from outlook_connector.schemas import Email

logger = get_logger()

# CloudEvents `source` for events this connector emits.
_EVENT_SOURCE = "/outlook-connector"
# The event types carried on the CloudEvents envelope.
EMAIL_RECEIVED = "email.received"  # inbound: mail observed -> bus


class EmailReceived(BaseModel):
    """The ``data`` payload of an ``email.received`` event.

    ``source_mailbox`` is the address of the mailbox that received the mail
    (provenance/routing). ``fetched_at`` is when the connector *observed* the
    message (the poll-run timestamp).
    """

    source_mailbox: str
    fetched_at: datetime.datetime
    email: Email


# Typed CloudEvents envelope: validates the payload on publish and subscribe.
EmailReceivedMessage = EggaiBaseMessage[EmailReceived]


def build_bus_event(
    message: OutlookMessage,
    *,
    source_mailbox: str,
    fetched_at: datetime.datetime,
    attachments: Sequence[OutlookAttachment] = (),
) -> EmailReceivedMessage:
    """Wrap a mapped email in the typed CloudEvents ``email.received`` envelope."""
    logger.debug("Building bus event", message_id=message.id)
    payload = EmailReceived(
        source_mailbox=source_mailbox,
        fetched_at=fetched_at,
        email=outlook_message_to_email(message, attachments),
    )
    return EmailReceivedMessage(source=_EVENT_SOURCE, type=EMAIL_RECEIVED, data=payload)


def build_transport(bus_config: BusConfig) -> Transport:
    """Construct the transport described by ``bus``."""
    if bus_config.transport == "kafka":
        if bus_config.broker_url:
            return KafkaTransport(bootstrap_servers=bus_config.broker_url)
        return KafkaTransport()
    if bus_config.transport == "redis":
        if bus_config.broker_url:
            return RedisTransport(url=bus_config.broker_url, max_len=bus_config.max_len)
        return RedisTransport(max_len=bus_config.max_len)
    if bus_config.transport == "inmemory":
        return InMemoryTransport()
    raise ValueError(f"Unknown transport: {bus_config.transport!r}")
