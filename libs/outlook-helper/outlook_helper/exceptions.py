"""Exceptions raised by outlook-helper."""

from __future__ import annotations


class GraphError(Exception):
    """Raised for any non-2xx Microsoft Graph response (after retries) and auth failures.

    Carries the HTTP status code plus, when available, the Graph error ``code``,
    a human-readable ``message``, and the originating ``request_id``.
    """

    def __init__(
        self,
        status_code: int,
        message: str,
        *,
        code: str | None = None,
        request_id: str | None = None,
    ) -> None:
        self.status_code = status_code
        self.message = message
        self.code = code
        self.request_id = request_id
        super().__init__(f"[{status_code}] {message}")
