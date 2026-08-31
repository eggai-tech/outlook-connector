import base64
import datetime
from typing import Annotated

from pydantic import BaseModel, BeforeValidator, PlainSerializer


def _decode_base64(value: object) -> object:
    if isinstance(value, str):
        return base64.b64decode(value)
    return value


# Raw bytes in memory, base64 on the JSON wire. Plain `bytes` would be
# serialized as UTF-8 and fail on real (binary) attachment content.
Base64OnWire = Annotated[
    bytes,
    BeforeValidator(_decode_base64),
    PlainSerializer(
        lambda value: base64.b64encode(value).decode("ascii"),
        return_type=str,
        when_used="json",
    ),
]


class EmailAttachment(BaseModel):
    """``body`` is ``None`` when the content is withheld: the attachment
    exceeded ``max_attachment_bytes``, or it is an item/reference attachment
    (a nested message or a link), which carries no bytes. ``size`` is the
    original size in bytes as Graph reports it, kept even when the body is
    withheld so consumers can log/decide without the content."""

    file_name: str
    content_type: str | None = None
    size: int | None = None
    body: Base64OnWire | None = None


class Email(BaseModel):
    """
    Barebones email object that we send on the bus.

    Separate from detailed outlook_helper.OutlookMessage that follows Graph API.
    """

    id: str  # unique id, preferably GraphAPI immutable
    # RFC 822 Message-ID header — stable across systems, the natural dedup key
    # for consumers. None when Graph did not return internetMessageId.
    internet_message_id: str | None = None
    from_addresses: list[str] = []
    to_addresses: list[str] = []
    subject: str | None = None
    received_at: datetime.datetime | None = None
    body_html: str | None = None
    body_text: str | None = None
    has_attachments: bool = (
        False  # can be True even if attachments field is not used and left blank
    )
    attachments: list[EmailAttachment] = []
    mime_content: str | None = None
