"""The reference in-memory backend.

Keeps saved emails in a dict and can be emptied, so a test starts from a known
state and can assert on what was stored. It is the reference implementation of
:class:`~outlook_connector.storage.EmailStorage` and the backend the test suite
uses. Nothing survives the process, so it is **not** meant for production.
"""

from outlook_connector.schemas import Email
from outlook_connector.storage import EmailStorage


class MemoryStorage(EmailStorage):
    """Stores emails in a dict keyed by ``message_id``.

    A second save of the same email replaces the first — this backend's answer
    to duplicates, which the API leaves to each backend.
    """

    def __init__(self):
        self.emails: dict[str, Email] = {}

    async def save(self, email: Email) -> str:
        """Keep ``email`` in memory; return its ``message_id``."""
        self.emails[email.message_id] = email
        return email.message_id

    def clear(self) -> None:
        """Drop everything, so the next test starts from a known state."""
        self.emails.clear()
