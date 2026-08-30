"""Boundary mapping: outlook-helper's ``OutlookMessage`` -> the owned bus model.

This is the single seam that decouples the bus contract from the dependency's
schema (per `the spec <docs/DESIGN.md#email-model>`). The helper makes
nearly every field ``Optional``; the owned :class:`~schemas.Email` makes several
required. Mapping is therefore **defensive** — a sparse-but-valid message maps
to a valid ``Email`` rather than raising, because a mapping that raised would
stall the whole mailbox (the poller processes a mailbox's batch in order).

The native ``has_attachments`` flag is free and always carried. Per-attachment
metadata *and content* cost an extra Graph call, so they are fetched by the
poller only for attachment-bearing mail and passed in here as ``attachments``;
when none is supplied the list maps to empty.
"""

import base64
from collections.abc import Sequence
from datetime import datetime

from outlook_helper import OutlookAttachment, OutlookMessage

from outlook_connector.schemas import (
    EMAIL_RECEIVED,
    Attachment,
    Email,
    EmailAddress,
    EmailReceived,
    EmailReceivedMessage,
)

# Threading headers we model as first-class fields; excluded from extra_headers
# so they don't appear twice.
_MODELED_HEADERS = {"in-reply-to", "references"}

# CloudEvents `source` for events this connector emits.
_EVENT_SOURCE = "/outlook-connector"


def _address(addr) -> EmailAddress:
    """Map a helper EmailAddress, tolerating a missing address string."""
    return EmailAddress(name=addr.name, address=addr.address or "")


def _attachment(att: OutlookAttachment) -> Attachment:
    """Map a helper attachment, tolerating its Optional name/type/size.

    The helper hands us the *decoded* attachment bytes; the owned model's
    ``content`` is a ``Base64Bytes`` field, so it must be base64-encoded here
    (otherwise the raw bytes are misread as base64 and mangled on the wire).
    Item/reference attachments carry no bytes, so ``content`` stays ``None``.
    """
    return Attachment(
        filename=att.name or "",
        content_type=att.content_type or "",
        size=att.size or 0,
        content=base64.b64encode(att.content) if att.content is not None else None,
    )


def _extra_headers(msg: OutlookMessage) -> dict[str, str]:
    """Carry non-modeled internet message headers as a flat dict (last wins)."""
    extra: dict[str, str] = {}
    for h in msg.internet_message_headers:
        if not h.name or h.name.lower() in _MODELED_HEADERS:
            continue
        extra[h.name] = h.value or ""
    return extra


def map_email(
    msg: OutlookMessage, attachments: Sequence[OutlookAttachment] = ()
) -> Email:
    """Map a helper ``OutlookMessage`` into the owned :class:`Email` model.

    ``attachments`` are the per-attachment metadata + content the poller
    fetched for attachment-bearing mail; it defaults to empty (no extra Graph
    call made).
    """
    body = msg.body.content if msg.body and msg.body.content else ""
    content_type = (msg.body.content_type or "").lower() if msg.body else ""
    # HTML is requested globally (lossless superset); default there if unknown.
    body_content_type = "text" if content_type == "text" else "html"

    return Email(
        message_id=msg.internet_message_id or msg.id,
        graph_id=msg.id,
        **{"from": _address(msg.from_) if msg.from_ else EmailAddress(address="")},
        to=[_address(a) for a in msg.to],
        cc=[_address(a) for a in msg.cc],
        subject=msg.subject or "",
        received_datetime=msg.received_at,
        extra_headers=_extra_headers(msg),
        conversation_id=msg.conversation_id,
        in_reply_to=msg.in_reply_to,
        references=msg.references,
        body=body,
        body_content_type=body_content_type,
        preview=msg.body_preview,
        # Only populated when the message was fetched with ``include_mime``;
        # the poller does not, so today this maps to None.
        mime=msg.mime_content,
        has_attachments=msg.has_attachments,
        attachments=[_attachment(a) for a in attachments],
    )


def build_event(
    msg: OutlookMessage,
    *,
    source_mailbox: str,
    fetched_at: datetime,
    attachments: Sequence[OutlookAttachment] = (),
) -> EmailReceivedMessage:
    """Wrap a mapped email in the typed CloudEvents ``email.received`` envelope."""
    payload = EmailReceived(
        source_mailbox=source_mailbox,
        fetched_at=fetched_at,
        email=map_email(msg, attachments),
    )
    return EmailReceivedMessage(source=_EVENT_SOURCE, type=EMAIL_RECEIVED, data=payload)
