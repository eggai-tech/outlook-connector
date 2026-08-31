"""Attachment handling: small inline attachments and large upload sessions.

Graph accepts attachments under ~3 MB inline (base64) on the message itself.
Larger files must be streamed to an upload session in chunks. This module is
pure data handling; the HTTP orchestration lives in the client.
"""

from __future__ import annotations

import base64
import mimetypes
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Union

#: Graph's single-request attachment ceiling (3 MB).
LARGE_ATTACHMENT_THRESHOLD = 3 * 1024 * 1024

#: Upload chunks should be a multiple of 320 KiB per Graph guidance.
DEFAULT_UPLOAD_CHUNK_SIZE = 5 * 320 * 1024  # 1,638,400 bytes


@dataclass
class Attachment:
    name: str
    content: bytes
    content_type: str = "application/octet-stream"


AttachmentSpec = Union[Attachment, str, Path]


def load_attachment(spec: AttachmentSpec) -> Attachment:
    """Coerce a path or :class:`Attachment` into an :class:`Attachment`."""
    if isinstance(spec, Attachment):
        return spec
    path = Path(spec)
    content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    return Attachment(
        name=path.name, content=path.read_bytes(), content_type=content_type
    )


def is_large(attachment: Attachment) -> bool:
    return len(attachment.content) > LARGE_ATTACHMENT_THRESHOLD


def to_inline_payload(attachment: Attachment) -> dict:
    """Build the Graph ``fileAttachment`` payload for an inline attachment."""
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": attachment.name,
        "contentType": attachment.content_type,
        "contentBytes": base64.b64encode(attachment.content).decode("ascii"),
    }


def upload_session_item(attachment: Attachment) -> dict:
    """Build the ``AttachmentItem`` body for ``createUploadSession``."""
    return {
        "attachmentType": "file",
        "name": attachment.name,
        "size": len(attachment.content),
        "contentType": attachment.content_type,
    }


def iter_upload_chunks(
    data: bytes, chunk_size: int = DEFAULT_UPLOAD_CHUNK_SIZE
) -> Iterator[tuple[bytes, str]]:
    """Yield ``(chunk_bytes, content_range_header)`` pairs covering ``data``."""
    total = len(data)
    for start in range(0, total, chunk_size):
        chunk = data[start : start + chunk_size]
        end = start + len(chunk) - 1
        yield chunk, f"bytes {start}-{end}/{total}"
