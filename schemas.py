"""The owned ``email.received`` bus contract.

The on-the-wire form is EggAI's CloudEvents 1.0 envelope (JSON) wrapping a typed
:class:`EmailReceived` payload. The models here are **owned** by the connector:
``outlook-helper``'s ``OutlookMessage`` is mapped into :class:`Email` at the
boundary, so the bus contract is decoupled from that dependency's schema.

Attachment *content* is carried inline alongside metadata when available
(``Base64Bytes`` keeps it JSON-safe on the wire). Item/reference attachments
have no bytes, so their ``content`` is ``None``. Large attachments inflate the
event and may exceed broker message limits (e.g. Kafka's ~1MB default
``max.message.bytes``) — a conscious tradeoff; the helper's
``download_attachment`` remains for streaming large files to disk instead.
"""

from datetime import datetime
from typing import Literal

from eggai.schemas import BaseMessage
from pydantic import Base64Bytes, BaseModel, ConfigDict, Field

# The event types carried on the CloudEvents envelope.
EMAIL_RECEIVED = "email.received"  # inbound: mail observed -> bus
EMAIL_SEND = "email.send"  # outbound: an agent's request -> mail


class EmailAddress(BaseModel):
    """A mail participant: an address with an optional display name."""

    name: str | None = None
    address: str


class Attachment(BaseModel):
    """Attachment metadata plus content when Graph provides bytes.

    ``content`` is the raw attachment bytes (base64-encoded on the wire via
    ``Base64Bytes``); it is ``None`` for item/reference attachments, which
    carry no ``contentBytes``.
    """

    filename: str
    content_type: str
    size: int  # byte length reported by Graph
    content: Base64Bytes | None = None


class Email(BaseModel):
    """An owned email model, mapped from ``outlook-helper``'s ``OutlookMessage``."""

    model_config = ConfigDict(populate_by_name=True)

    # Identity
    message_id: str  # RFC 822 internetMessageId, e.g. "<abc@host>"
    graph_id: str  # Graph internal item id, kept for follow-up Graph calls

    # Core headers
    from_: EmailAddress = Field(alias="from")  # "from" is a Python keyword
    to: list[EmailAddress] = Field(default_factory=list)
    cc: list[EmailAddress] = Field(default_factory=list)
    subject: str
    received_datetime: datetime
    extra_headers: dict[str, str] = Field(default_factory=dict)

    # Threading
    conversation_id: str | None = None
    in_reply_to: str | None = None
    references: list[str] = Field(default_factory=list)

    # Body (as delivered, no lossy conversion)
    body: str
    body_content_type: Literal["html", "text"]
    preview: str | None = None  # Graph's bodyPreview

    # Attachments (metadata only; empty for mail without attachments)
    has_attachments: bool = False
    attachments: list[Attachment] = Field(default_factory=list)


class EmailReceived(BaseModel):
    """The ``data`` payload of an ``email.received`` event.

    ``source_mailbox`` is the address of the mailbox that received the mail
    (provenance/routing). ``fetched_at`` is when the connector *observed* the
    message (the poll-run timestamp), not when the cursor advanced.
    """

    source_mailbox: str
    fetched_at: datetime
    email: Email


# Typed CloudEvents envelope: validates the payload on publish and subscribe.
EmailReceivedMessage = BaseMessage[EmailReceived]


class EmailSend(BaseModel):
    """The ``data`` payload of an ``email.send`` event — an outbound request.

    The mirror of :class:`EmailReceived`: where that reports mail the connector
    *observed*, this is a request an agent puts on the bus for the connector to
    *deliver* as mail. ``mailbox`` is the address to send **from**; the connector
    rejects any value not in its configured mailbox set (an agent cannot send as
    an arbitrary identity).

    ``reply_to_graph_id`` opts into threaded delivery: when set, the connector
    replies to that Graph message (preserving the conversation) instead of
    composing a fresh mail, so ``subject`` is taken from the original. The
    ``graph_id`` an agent uses here is the one carried on the inbound
    :class:`Email` it is replying to.

    Attachments are intentionally omitted from this first cut: ``outlook-helper``
    accepts attachments as file paths, which does not match this contract's
    inline-bytes model, and outbound attachment content is out of MVP scope.
    """

    mailbox: str  # send-from address; must be a configured mailbox
    to: list[EmailAddress] = Field(min_length=1)
    cc: list[EmailAddress] = Field(default_factory=list)
    bcc: list[EmailAddress] = Field(default_factory=list)
    subject: str = ""
    body: str
    body_content_type: Literal["html", "text"] = "html"
    # Optional threaded reply: the Graph item id of the message to reply to.
    reply_to_graph_id: str | None = None


class EmailSendMessage(BaseMessage[EmailSend]):
    """Typed CloudEvents envelope for ``email.send``.

    ``type`` carries a default (unlike the base envelope's required ``type``) so
    a typed ``data_type=EmailSendMessage`` subscription matches on it: EggAI's
    typed-subscription filter compares each incoming event against this default
    and skips anything else (e.g. ``email.received`` on the same channel).
    """

    type: str = EMAIL_SEND
