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
        # body is None for content withheld upstream: item/reference
        # attachments (no bytes to fetch) or over the configured size cap.
        attachments=[
            EmailAttachment(
                file_name=attachment.name or "",
                content_type=attachment.content_type,
                size=attachment.size
                if attachment.size is not None
                else (len(attachment.content) if attachment.content is not None else None),
                body=attachment.content,
            )
            for attachment in attachments
        ],
        mime_content=message.mime_content,
    )
