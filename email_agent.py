from datetime import datetime, timezone

from eggai import Agent, Channel

from schemas import (
    EMAIL_RECEIVED,
    Attachment,
    Email,
    EmailAddress,
    EmailReceived,
    EmailReceivedMessage,
)


email_agent = Agent("EmailAgent")
emails_channel = Channel("emails")


# Throwaway: publishes one hand-built email.received event to the bus so
# downstream consumers can build against the contract before the connector's
# polling internals (Piece 3) exist. Remove once real mail flows.
async def create_email():
    print("Publishing email.received")
    payload = EmailReceived(
        source_mailbox="invoices@egg-ai.com",
        fetched_at=datetime.now(timezone.utc),
        email=Email(
            message_id="<CADnf...@mail.gmail.com>",
            graph_id="AAMkAGI2THVSAAA=",
            from_=EmailAddress(name="Alice", address="alice@example.com"),
            to=[EmailAddress(name="Invoices", address="invoices@egg-ai.com")],
            subject="March invoice",
            received_datetime=datetime.now(timezone.utc),
            conversation_id="AAQkAGI2-conv",
            body="<p>Please find attached.</p>",
            body_content_type="html",
            preview="Please find attached.",
            has_attachments=True,
            attachments=[
                Attachment(
                    filename="invoice-2026-03.pdf",
                    content_type="application/pdf",
                    size=12345,
                )
            ],
        ),
    )
    message = EmailReceivedMessage(
        source="/outlook-connector", type=EMAIL_RECEIVED, data=payload
    )
    await emails_channel.publish(message)
