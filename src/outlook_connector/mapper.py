from collections.abc import Sequence

from outlook_helper.schemas import OutlookAttachment, OutlookMessage

from .schemas import Email, EmailAttachment


def outlook_message_to_email(
    message: OutlookMessage,
    attachments: Sequence[OutlookAttachment] = (),
) -> Email:
    body = message.body
    content = body.content if body else None
    content_type = body.content_type if body else None
    sender = message.from_.address if message.from_ else None
    return Email(
        id=message.id,
        internet_message_id=message.internet_message_id,
        from_addresses=[sender] if sender else [],
        to_addresses=[email.address for email in message.to if email.address],
        subject=message.subject,
        received_at=message.received_at,
        body_text=content if content_type == "text" else None,
        body_html=content if content_type == "html" else None,
        has_attachments=message.has_attachments,
        # Item/reference attachments carry no bytes — only file attachments map.
        attachments=[
            EmailAttachment(
                file_name=attachment.name or "",
                content_type=attachment.content_type,
                body=attachment.content,
            )
            for attachment in attachments
            if attachment.content is not None
        ],
        mime_content=message.mime_content,
    )
