from pydantic import BaseModel


class EmailAttachment(BaseModel):
    file_name: str
    content_type: str | None = None
    body: bytes


class Email(BaseModel):
    """
    Barebones email object that we send on the bus.

    Separate from detailed outlook_helper.OutlookMessage that follows Graph API.
    """

    id: str  # unique id, preferably GraphAPI immutable
    from_addresses: list[str] = []
    to_addresses: list[str] = []
    subject: str | None = None
    body_html: str | None = None
    body_text: str | None = None
    has_attachments: bool = (
        False  # can be True even if attachments field is not used and left blank
    )
    attachments: list[EmailAttachment] = []
    mime_content: str | None = None
