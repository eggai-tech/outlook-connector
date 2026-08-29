"""Resolve user-facing folder references to something Graph can address.

A reference may be a well-known name (``inbox``, ``drafts``, ...), a folder's
display name (resolved to its id and cached), or a raw folder id (passed
through). Well-known names and raw ids are accepted directly by Graph in the
``/mailFolders/{id}`` path; only display names need a lookup.
"""

from __future__ import annotations

from outlook_helper.http import GraphSession

WELL_KNOWN_FOLDERS: frozenset[str] = frozenset(
    {
        "archive",
        "clutter",
        "conflicts",
        "conversationhistory",
        "deleteditems",
        "drafts",
        "inbox",
        "junkemail",
        "localfailures",
        "msgfolderroot",
        "outbox",
        "recoverableitemsdeletions",
        "scheduled",
        "searchfolders",
        "sentitems",
        "serverfailures",
        "syncissues",
    }
)


class FolderResolver:
    def __init__(self, session: GraphSession, base_path: str) -> None:
        self._session = session
        self._base_path = base_path
        self._cache: dict[str, str] = {}

    def resolve(self, reference: str) -> str:
        """Return a folder id/well-known name usable directly in a Graph path."""
        lowered = reference.lower()
        if lowered in WELL_KNOWN_FOLDERS:
            return lowered
        if reference in self._cache:
            return self._cache[reference]

        escaped = reference.replace("'", "''")
        matches = list(
            self._session.paginate(
                f"{self._base_path}/mailFolders",
                params={"$filter": f"displayName eq '{escaped}'"},
            )
        )
        if matches:
            folder_id = matches[0]["id"]
            self._cache[reference] = folder_id
            return folder_id

        # Not a well-known name and no folder by that display name: assume it is
        # already a raw folder id.
        return reference

    def invalidate(self) -> None:
        self._cache.clear()
