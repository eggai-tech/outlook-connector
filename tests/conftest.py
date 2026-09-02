import datetime

import pytest
from outlook_helper.schemas import (
    EmailAddress,
    OutlookAttachment,
    OutlookBody,
    OutlookMessage,
)

T0 = datetime.datetime(2026, 1, 1, 12, 0, 0, tzinfo=datetime.UTC)


def make_message(
    message_id: str = "m1",
    *,
    received_at: datetime.datetime | None = T0,
    sender: str | None = "sender@example.com",
    has_attachments: bool = False,
    body_content_type: str = "html",
    body_content: str = "<p>hi</p>",
) -> OutlookMessage:
    return OutlookMessage(
        id=message_id,
        subject=f"subject {message_id}",
        from_=EmailAddress(name="Sender", address=sender) if sender else None,
        to=[EmailAddress(address="to@example.com")],
        received_at=received_at,
        body=OutlookBody(content_type=body_content_type, content=body_content),
        has_attachments=has_attachments,
        internet_message_id=f"<{message_id}@example.com>",
    )


def make_attachment(
    name: str = "doc.pdf",
    *,
    content: bytes | None = b"%PDF-1.4",
    content_type: str = "application/pdf",
) -> OutlookAttachment:
    return OutlookAttachment(
        id=f"att-{name}",
        name=name,
        content_type=content_type,
        size=len(content) if content else 0,
        content=content,
    )


class FakeClient:
    """Stands in for outlook_helper.OutlookClient.

    ``search_email`` serves the poller's ids-only folder listing (returning the
    same message objects as stubs); ``get_email`` serves the per-message full
    fetch. ``self.messages`` is mutable so tests can simulate mail being moved
    out of (or back into) the folder between polls. Like the real client,
    ``get_email`` only fills ``mime_content`` (with ``self.mime_content``) when
    called with ``include_mime=True``.
    """

    def __init__(self, messages=(), attachments=(), mime_content="MIME-Version: 1.0\r\n"):
        self.messages = list(messages)
        self.attachments = list(attachments)
        self.mime_content = mime_content
        self.search_calls: list[dict] = []
        self.get_email_calls: list[str] = []
        self.get_email_kwargs: list[dict] = []
        self.attachment_calls: list[str] = []

    def search_email(self, **kwargs):
        self.search_calls.append(kwargs)
        top = kwargs.get("top")
        return list(self.messages) if top is None else list(self.messages)[:top]

    def get_email(self, message_id, **kwargs):
        self.get_email_calls.append(message_id)
        self.get_email_kwargs.append(kwargs)
        for message in self.messages:
            if message.id == message_id:
                if kwargs.get("include_mime"):
                    return message.model_copy(update={"mime_content": self.mime_content})
                return message
        raise AssertionError(f"get_email for unknown id {message_id!r}")

    def get_attachments(self, message_id):
        self.attachment_calls.append(message_id)
        return list(self.attachments)


class FakeChannel:
    """Stands in for eggai.Channel: records published events, can fail."""

    def __init__(self, fail_on: set[str] | None = None):
        self.published = []
        self.fail_on = fail_on or set()

    async def publish(self, event):
        if event.data.email.id in self.fail_on:
            raise RuntimeError("bus down")
        self.published.append(event)


@pytest.fixture
def t0() -> datetime.datetime:
    return T0
