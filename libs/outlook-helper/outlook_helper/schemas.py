"""Pydantic models returned by the library.

These decouple callers from Graph's wire format: fields are snake_case and the
nested ``emailAddress`` envelope Graph uses for senders/recipients is flattened
to :class:`EmailAddress`.
"""

from __future__ import annotations

import base64
from datetime import datetime
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, field_validator


class GraphModel(BaseModel):
    model_config = ConfigDict(populate_by_name=True, extra="ignore")


class EmailAddress(GraphModel):
    name: str | None = None
    address: str | None = None


class OutlookBody(GraphModel):
    content_type: str | None = Field(default=None, alias="contentType")
    content: str | None = None


class InternetMessageHeader(GraphModel):
    name: str | None = None
    value: str | None = None


def _unwrap_email_address(value: Any) -> Any:
    """Turn Graph's ``{"emailAddress": {...}}`` envelope into the inner object."""
    if isinstance(value, dict) and "emailAddress" in value:
        return value["emailAddress"]
    return value


class OutlookMessage(GraphModel):
    """A mail message.

    ``id`` is Graph's *immutable* id: :class:`~outlook_helper.http.GraphSession`
    sends ``Prefer: IdType="ImmutableId"`` on every request, so the id stays the
    same when the message moves between folders and is safe to persist as the
    message's unique key.
    """

    id: str
    subject: str | None = None
    from_: EmailAddress | None = Field(default=None, alias="from")
    to: list[EmailAddress] = Field(default_factory=list, alias="toRecipients")
    cc: list[EmailAddress] = Field(default_factory=list, alias="ccRecipients")
    bcc: list[EmailAddress] = Field(default_factory=list, alias="bccRecipients")
    received_at: datetime | None = Field(default=None, alias="receivedDateTime")
    sent_at: datetime | None = Field(default=None, alias="sentDateTime")
    body_preview: str | None = Field(default=None, alias="bodyPreview")
    body: OutlookBody | None = None
    is_read: bool | None = Field(default=None, alias="isRead")
    has_attachments: bool = Field(default=False, alias="hasAttachments")
    importance: str | None = None
    web_link: str | None = Field(default=None, alias="webLink")
    conversation_id: str | None = Field(default=None, alias="conversationId")
    parent_folder_id: str | None = Field(default=None, alias="parentFolderId")
    internet_message_id: str | None = Field(default=None, alias="internetMessageId")
    internet_message_headers: list[InternetMessageHeader] = Field(
        default_factory=list, alias="internetMessageHeaders"
    )
    #: The whole message in MIME format (a ``.eml``): headers, bodies and
    #: attachments in one string. Graph never returns it alongside the message
    #: JSON, so it is ``None`` unless it was asked for with ``include_mime``.
    mime_content: str | None = None

    @property
    def in_reply_to(self) -> str | None:
        for h in self.internet_message_headers:
            if h.name and h.name.lower() == "in-reply-to":
                return h.value
        return None

    @property
    def references(self) -> list[str]:
        for h in self.internet_message_headers:
            if h.name and h.name.lower() == "references":
                return h.value.split() if h.value else []
        return []

    @field_validator("from_", mode="before")
    @classmethod
    def _flatten_sender(cls, value: Any) -> Any:
        return _unwrap_email_address(value)

    @field_validator("to", "cc", "bcc", mode="before")
    @classmethod
    def _flatten_recipients(cls, value: Any) -> Any:
        if isinstance(value, list):
            return [_unwrap_email_address(item) for item in value]
        return value


class OutlookAttachmentMeta(GraphModel):
    id: str
    name: str | None = None
    content_type: str | None = Field(default=None, alias="contentType")
    size: int | None = None
    is_inline: bool = Field(default=False, alias="isInline")


class OutlookAttachment(OutlookAttachmentMeta):
    """Attachment metadata plus decoded content.

    ``content`` is the decoded bytes of a file attachment. Item attachments
    (a nested message) and reference attachments (a link) carry no
    ``contentBytes``, so ``content`` stays ``None``.
    """

    content: bytes | None = Field(default=None, alias="contentBytes")

    @field_validator("content", mode="before")
    @classmethod
    def _decode_content(cls, value: Any) -> Any:
        # Graph sends fileAttachment.contentBytes as base64; others omit it.
        if isinstance(value, str):
            return base64.b64decode(value)
        return value


class OutlookFolder(GraphModel):
    id: str
    display_name: str | None = Field(default=None, alias="displayName")
    parent_folder_id: str | None = Field(default=None, alias="parentFolderId")
    child_folder_count: int | None = Field(default=None, alias="childFolderCount")
    total_item_count: int | None = Field(default=None, alias="totalItemCount")
    unread_item_count: int | None = Field(default=None, alias="unreadItemCount")
