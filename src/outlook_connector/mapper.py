from outlook_helper.schemas import OutlookMessage

from .schemas import Email


def outlook_message_to_email(message: OutlookMessage) -> Email:
    return Email(
        id=message.id,
        from_address=message.from_.address,
        to_addresses=[email.address for email in message.to],
        subject=message.subject,
        sent_at=message.sent_at,
        body_text=message.body.content if message.body.content_type == "text" else None,
        body_html=message.body.content if message.body.content_type == "html" else None,
        has_attachments=message.has_attachments,
        mime_content=message.mime_content,
    )
