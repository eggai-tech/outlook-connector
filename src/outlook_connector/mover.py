"""Bus listener that files a message into another mailbox folder.

The mirror image of the poller: the connector publishes ``email.received``
events, and a consumer asks it back — over the bus — to move one of those
messages. The move itself is outlook-helper's
:meth:`~outlook_helper.OutlookClient.move_email`, which resolves the
destination on its own (a well-known Graph name like ``archive``, a folder's
display name, or a raw folder id).

Commands ride the *same* channel as the outbound events: the subscription is
typed on :class:`MoveEmailMessage`, so eggai hands the handler only envelopes
whose ``type`` is ``email.move`` and skips everything else — including this
connector's own ``email.received`` events — without ever entering it.

Per `the spec <development/specs/initial.md#commands>` a command gets **no
reply**: a failure is logged and the command is dropped rather than retried, so
one unmovable message never stalls the stream. Note that Graph gives the moved
message a *new* id in the destination folder — a replayed command therefore
carries a stale id and fails with a logged "not found" instead of moving
anything twice.
"""

import asyncio

import structlog
from eggai import Agent, Channel
from eggai.schemas import BaseMessage as EggaiBaseMessage
from eggai.transport.base import Transport
from outlook_helper import OutlookClient
from pydantic import BaseModel, Field

from outlook_connector.poller import GRAPH_ERRORS

logger = structlog.get_logger()

# The event type carried on the CloudEvents envelope.
EMAIL_MOVE = "email.move"  # inbound: bus -> mailbox command


class MoveEmail(BaseModel):
    """The ``data`` payload of an ``email.move`` command.

    ``message_id`` is the Graph id published as ``email.id``. ``mailbox`` is
    optional provenance: when set it must name the mailbox this connector
    serves, otherwise the command is not ours to apply.
    """

    message_id: str = Field(min_length=1)
    # Well-known Graph name ("archive"), a folder display name, or a folder id.
    destination_folder: str = Field(min_length=1)
    mailbox: str | None = None


class MoveEmailMessage(EggaiBaseMessage[MoveEmail]):
    """Typed CloudEvents envelope for ``email.move``.

    ``type`` carries a default — that is the discriminator eggai matches a
    ``data_type`` subscription on, so envelopes of any other type on the same
    channel are skipped instead of handled.
    """

    type: str = EMAIL_MOVE


class Mover:
    """Applies ``email.move`` commands to one mailbox.

    Takes the client rather than building one: a connector serves a single
    mailbox, so it shares the poller's — same credentials, same folder-name
    cache, one token to refresh.

    Calls the synchronous outlook-helper client; the listener wraps it in
    ``asyncio.to_thread`` so a blocking request or ``Retry-After`` sleep never
    stalls the event loop.
    """

    def __init__(self, *, mailbox: str, client: OutlookClient):
        self.mailbox = mailbox
        self.client = client

    def apply(self, command: MoveEmail) -> None:
        """Move one message, logging the outcome. Never raises: see the module
        docstring on why a failed command is dropped rather than retried."""
        if command.mailbox is not None and command.mailbox != self.mailbox:
            logger.warning(
                "Ignoring move command addressed to another mailbox",
                message=command.message_id,
                requested_mailbox=command.mailbox,
                mailbox=self.mailbox,
            )
            return

        try:
            moved = self.client.move_email(command.message_id, command.destination_folder)
        except GRAPH_ERRORS as exc:
            logger.error(
                "Move failed",
                message=command.message_id,
                destination=command.destination_folder,
                error=f"{type(exc).__name__}: {exc}",
            )
            return

        logger.info(
            "Message moved",
            message=command.message_id,
            destination=command.destination_folder,
            moved_message=moved.id,  # Graph reissues the id in the new folder
        )


def build_mover_agent(mover: Mover, *, channel: str, transport: Transport) -> Agent:
    """Wire ``mover`` to the ``email.move`` commands on ``channel``.

    The returned agent is inert until awaited on ``start()``.
    """
    commands = Channel(channel, transport=transport)
    agent = Agent("outlook-connector-mover", transport=transport)

    @agent.subscribe(channel=commands, data_type=MoveEmailMessage)
    async def move_email(event: MoveEmailMessage) -> None:
        logger.debug("Move command received", message=event.data.message_id)
        await asyncio.to_thread(mover.apply, event.data)

    return agent
