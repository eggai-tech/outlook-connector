"""The outbound send listener: consume ``email.send`` events off the bus.

The mirror of the inbound :class:`~poller.Poller`. Where the poller fetches mail
and *publishes* ``email.received``, this *subscribes* to the owned
:class:`~schemas.EmailSend` contract and hands each request to an injected
``send`` collaborator, which delivers it as mail.

Subscription uses EggAI's **typed** delivery (``data_type=EmailSendMessage``):
the transport validates each raw event against the envelope and invokes the
handler only for events whose ``type`` is ``email.send`` — every other event on
the channel (notably the connector's own ``email.received``) is skipped before
the handler runs. The handler therefore always receives a validated
:class:`~schemas.EmailSendMessage`, never a stray dict.

The ``send`` collaborator is injected so the listener's event handling is
testable without a live Graph API or bus; :func:`service.build_send_listener`
wires the real one (one app-only ``OutlookClient`` per mailbox).
"""

import logging
from collections.abc import Awaitable, Callable

from eggai import Agent, Channel
from eggai.transport.base import Transport

from schemas import EmailSend, EmailSendMessage

logger = logging.getLogger(__name__)

# Deliver one outbound request as mail (async-wrapped Graph call). Injected for
# testability; raising aborts handling of that event (the broker may redeliver).
Send = Callable[[EmailSend], Awaitable[None]]


def make_send_listener(*, channel: str, transport: Transport, send: Send) -> Agent:
    """Build the agent that turns ``email.send`` events into ``send`` calls.

    Subscribes on the same configured ``channel`` the connector publishes to; the
    typed ``data_type`` filter keeps it from reacting to its own inbound events.
    The returned :class:`~eggai.Agent` is inert until ``await agent.start()``.
    """
    agent = Agent("outlook-send-listener", transport=transport)
    subscribe_channel = Channel(channel, transport=transport)

    @agent.subscribe(channel=subscribe_channel, data_type=EmailSendMessage)
    async def on_send(message: EmailSendMessage) -> None:
        request = message.data
        logger.info(
            "email.send: from=%s to=%d recipient(s) reply=%s subject=%r",
            request.mailbox,
            len(request.to),
            bool(request.reply_to_graph_id),
            request.subject,
        )
        await send(request)

    return agent
