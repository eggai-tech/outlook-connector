"""The high-level Outlook client exposing the README's verbs.

One :class:`OutlookClient` is bound to a single mailbox. It owns a
:class:`~outlook_helper.http.GraphSession` and a
:class:`~outlook_helper.folders.FolderResolver`, and maps Graph JSON to the
Pydantic models in :mod:`outlook_helper.models`.
"""

from __future__ import annotations

import itertools
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Iterator, Union

from outlook_helper.attachments import (
    Attachment,
    AttachmentSpec,
    is_large,
    iter_upload_chunks,
    load_attachment,
    to_inline_payload,
    upload_session_item,
)
from outlook_helper.auth import Credential
from outlook_helper.folders import FolderResolver
from outlook_helper.http import DEFAULT_BASE_URL, GraphSession
from outlook_helper.schemas import (
    OutlookAttachment,
    OutlookAttachmentMeta,
    EmailAddress,
    OutlookFolder,
    OutlookMessage,
)

RecipientSpec = Union[str, EmailAddress]
Recipients = Union[RecipientSpec, Iterable[RecipientSpec]]

# Every modeled property. ``$select`` restricts the response to *only* the listed
# properties, so to pull headers without dropping the default fields we must
# enumerate them all plus the two that aren't returned by default.
_MESSAGE_SELECT = (
    "id,subject,from,toRecipients,ccRecipients,bccRecipients,"
    "receivedDateTime,sentDateTime,bodyPreview,body,isRead,hasAttachments,"
    "importance,webLink,conversationId,parentFolderId,"
    "internetMessageId,internetMessageHeaders"
)

_PREFER_HTML = {"Prefer": 'outlook.body-content-type="html"'}


class OutlookClient:
    def __init__(
        self,
        credential: Credential,
        mailbox: str | None = None,
        *,
        session: GraphSession | None = None,
        base_url: str = DEFAULT_BASE_URL,
    ) -> None:
        self._credential = credential
        self._mailbox = mailbox
        self._session = session or GraphSession(credential, base_url=base_url)
        self._base_path = self._compute_base_path()
        self._folders = FolderResolver(self._session, self._base_path)

    def _compute_base_path(self) -> str:
        if self._mailbox:
            return f"/users/{self._mailbox}"
        if getattr(self._credential, "supports_me", False):
            return "/me"
        raise ValueError(
            "An app-only credential has no signed-in user; pass a mailbox."
        )

    @property
    def base_path(self) -> str:
        return self._base_path

    @property
    def credential(self) -> Credential:
        return self._credential

    # --- read operations ---

    def get_email(
        self,
        message_id: str,
        *,
        include_headers: bool = False,
        html_body: bool = False,
        include_mime: bool = False,
    ) -> OutlookMessage:
        """Fetch one message by id.

        ``include_mime`` additionally fetches the message's MIME content -- the
        whole ``.eml``, headers and attachments included -- into
        :attr:`~outlook_helper.schemas.OutlookMessage.mime_content`. Graph only
        serves it from a separate ``$value`` endpoint, so it costs a second
        request.
        """
        params: dict[str, Any] = {}
        if include_headers:
            params["$select"] = _MESSAGE_SELECT
        headers = dict(_PREFER_HTML) if html_body else None
        path = f"{self._base_path}/messages/{message_id}"
        data = self._session.get_json(path, params or None, headers=headers)
        message = OutlookMessage.model_validate(data)
        if include_mime:
            message.mime_content = self._session.get_text(f"{path}/$value")
        return message

    def list_messages(
        self,
        folder: str = "inbox",
        *,
        top: int | None = None,
        include_headers: bool = False,
        html_body: bool = False,
    ) -> Iterator[OutlookMessage]:
        folder_id = self._folders.resolve(folder)
        path = f"{self._base_path}/mailFolders/{folder_id}/messages"
        params: dict[str, Any] = {"$orderby": "receivedDateTime desc"}
        if include_headers:
            params["$select"] = _MESSAGE_SELECT
        headers = dict(_PREFER_HTML) if html_body else None
        return self._iter_messages(path, params, top=top, headers=headers)

    def search_email(
        self,
        *,
        sender: str | None = None,
        subject_contains: str | None = None,
        since: datetime | str | None = None,
        until: datetime | str | None = None,
        since_exclusive: datetime | str | None = None,
        until_exclusive: datetime | str | None = None,
        unread: bool | None = None,
        has_attachments: bool | None = None,
        folder: str | None = None,
        top: int | None = None,
        oldest_first: bool = False,
        include_headers: bool = False,
        html_body: bool = False,
        ids_only: bool = False,
    ) -> Iterator[OutlookMessage]:
        if ids_only and include_headers:
            raise ValueError("ids_only and include_headers are mutually exclusive")
        clauses = _build_filter_clauses(
            sender=sender,
            subject_contains=subject_contains,
            since=since,
            until=until,
            since_exclusive=since_exclusive,
            until_exclusive=until_exclusive,
            unread=unread,
            has_attachments=has_attachments,
        )
        if folder is not None:
            folder_id = self._folders.resolve(folder)
            path = f"{self._base_path}/mailFolders/{folder_id}/messages"
        else:
            path = f"{self._base_path}/messages"
        # oldest_first matters when combined with ``top``: a bounded batch then
        # drains a backlog from the oldest end instead of returning the newest
        # N and starving the tail.
        order = "receivedDateTime asc" if oldest_first else "receivedDateTime desc"
        params: dict[str, Any] = {"$orderby": order}
        if clauses:
            params["$filter"] = " and ".join(clauses)
        if include_headers:
            params["$select"] = _MESSAGE_SELECT
        if ids_only:
            # id + the sort key, nothing else: a whole-folder listing at
            # minimal payload cost, for callers that fetch bodies separately.
            params["$select"] = "id,receivedDateTime"
        headers = dict(_PREFER_HTML) if html_body else None
        return self._iter_messages(path, params, top=top, headers=headers)

    def list_attachments(self, message_id: str) -> list[OutlookAttachmentMeta]:
        items = self._session.paginate(
            f"{self._base_path}/messages/{message_id}/attachments",
            {"$select": "id,name,contentType,size,isInline"},
        )
        return [OutlookAttachmentMeta.model_validate(item) for item in items]

    def get_attachments(self, message_id: str) -> list[OutlookAttachment]:
        """Return attachments with their content.

        Unlike :meth:`list_attachments`, this omits the metadata-only
        ``$select`` so Graph returns ``contentBytes`` inline; each file
        attachment's bytes are decoded into ``OutlookAttachment.content``.
        Item/reference attachments carry no bytes, so their ``content`` is
        ``None`` (use :meth:`download_attachment` for large files you would
        rather stream to disk than hold in memory).
        """
        items = self._session.paginate(
            f"{self._base_path}/messages/{message_id}/attachments",
        )
        return [OutlookAttachment.model_validate(item) for item in items]

    # --- send / drafts ---

    def send_email(
        self,
        to: Recipients,
        subject: str,
        body: str,
        *,
        cc: Recipients | None = None,
        bcc: Recipients | None = None,
        attachments: Iterable[AttachmentSpec] | None = None,
        html: bool = False,
    ) -> None:
        loaded = [load_attachment(a) for a in (attachments or [])]
        if any(is_large(a) for a in loaded):
            # Large attachments require an upload session, which needs a draft.
            draft = self._create_draft_with_attachments(
                to, subject, body, cc, bcc, html, loaded
            )
            self.send_draft(draft.id)
            return
        message = _build_message(to, subject, body, cc, bcc, html, loaded)
        self._session.request(
            "POST",
            f"{self._base_path}/sendMail",
            json={"message": message, "saveToSentItems": True},
        )

    def create_draft(
        self,
        to: Recipients,
        subject: str,
        body: str,
        *,
        cc: Recipients | None = None,
        bcc: Recipients | None = None,
        attachments: Iterable[AttachmentSpec] | None = None,
        html: bool = False,
    ) -> OutlookMessage:
        loaded = [load_attachment(a) for a in (attachments or [])]
        return self._create_draft_with_attachments(
            to, subject, body, cc, bcc, html, loaded
        )

    def update_draft(
        self,
        message_id: str,
        *,
        subject: str | None = None,
        body: str | None = None,
        to: Recipients | None = None,
        cc: Recipients | None = None,
        bcc: Recipients | None = None,
        html: bool = False,
    ) -> OutlookMessage:
        patch: dict[str, Any] = {}
        if subject is not None:
            patch["subject"] = subject
        if body is not None:
            patch["body"] = _body_payload(body, html)
        if to is not None:
            patch["toRecipients"] = _to_recipients(to)
        if cc is not None:
            patch["ccRecipients"] = _to_recipients(cc)
        if bcc is not None:
            patch["bccRecipients"] = _to_recipients(bcc)
        data = self._session.request(
            "PATCH", f"{self._base_path}/messages/{message_id}", json=patch
        ).json()
        return OutlookMessage.model_validate(data)

    def send_draft(self, message_id: str) -> None:
        self._session.request("POST", f"{self._base_path}/messages/{message_id}/send")

    def discard_draft(self, message_id: str) -> None:
        self._session.request("DELETE", f"{self._base_path}/messages/{message_id}")

    def create_reply_draft(
        self, message_id: str, *, reply_all: bool = False
    ) -> OutlookMessage:
        """Create a reply draft that inherits the original's conversationId
        (and In-Reply-To/References headers), so it threads when sent.

        The draft is pre-addressed to the original sender; override recipients
        with :meth:`update_draft` before sending to redirect or suppress that
        (e.g. ``update_draft(draft.id, to=shared_mailbox)`` to thread a message
        into a shared mailbox without notifying the original sender).
        """
        action = "createReplyAll" if reply_all else "createReply"
        return OutlookMessage.model_validate(
            self._session.request(
                "POST", f"{self._base_path}/messages/{message_id}/{action}"
            ).json()
        )

    def reply(
        self,
        message_id: str,
        body: str,
        *,
        reply_all: bool = False,
        attachments: Iterable[AttachmentSpec] | None = None,
        html: bool = False,
    ) -> None:
        draft = self.create_reply_draft(message_id, reply_all=reply_all)
        self._session.request(
            "PATCH",
            f"{self._base_path}/messages/{draft.id}",
            json={"body": _body_payload(body, html)},
        )
        for att in (load_attachment(a) for a in (attachments or [])):
            self._add_attachment(draft.id, att)
        self.send_draft(draft.id)

    def download_attachment(
        self, message_id: str, attachment_id: str, dest_path: Path | str
    ) -> Path:
        return self._session.download(
            f"{self._base_path}/messages/{message_id}/attachments/{attachment_id}/$value",
            dest_path,
        )

    # --- folder operations ---

    def list_folders(self) -> list[OutlookFolder]:
        items = self._session.paginate(f"{self._base_path}/mailFolders")
        return [OutlookFolder.model_validate(item) for item in items]

    def create_folder(self, name: str, *, parent: str | None = None) -> OutlookFolder:
        if parent is not None:
            parent_id = self._folders.resolve(parent)
            path = f"{self._base_path}/mailFolders/{parent_id}/childFolders"
        else:
            path = f"{self._base_path}/mailFolders"
        data = self._session.request("POST", path, json={"displayName": name}).json()
        self._folders.invalidate()
        return OutlookFolder.model_validate(data)

    def delete_email(self, message_id: str, *, permanent: bool = False) -> None:
        """Delete a message. Soft (move to Deleted Items) unless ``permanent``."""
        if permanent:
            self._session.request(
                "POST", f"{self._base_path}/messages/{message_id}/permanentDelete"
            )
        else:
            self._session.request("DELETE", f"{self._base_path}/messages/{message_id}")

    def move_email(self, message_id: str, dest_folder: str) -> OutlookMessage:
        destination_id = self._folders.resolve(dest_folder)
        data = self._session.request(
            "POST",
            f"{self._base_path}/messages/{message_id}/move",
            json={"destinationId": destination_id},
        ).json()
        return OutlookMessage.model_validate(data)

    # --- helpers ---

    def _create_draft_with_attachments(
        self,
        to: Recipients,
        subject: str,
        body: str,
        cc: Recipients | None,
        bcc: Recipients | None,
        html: bool,
        loaded: list[Attachment],
    ) -> OutlookMessage:
        small = [a for a in loaded if not is_large(a)]
        large = [a for a in loaded if is_large(a)]
        message = _build_message(to, subject, body, cc, bcc, html, small)
        draft = OutlookMessage.model_validate(
            self._session.request(
                "POST", f"{self._base_path}/messages", json=message
            ).json()
        )
        for att in large:
            self._upload_large_attachment(draft.id, att)
        return draft

    def _add_attachment(self, message_id: str, attachment: Attachment) -> None:
        if is_large(attachment):
            self._upload_large_attachment(message_id, attachment)
        else:
            self._session.request(
                "POST",
                f"{self._base_path}/messages/{message_id}/attachments",
                json=to_inline_payload(attachment),
            )

    def _upload_large_attachment(self, message_id: str, attachment: Attachment) -> None:
        session = self._session.request(
            "POST",
            f"{self._base_path}/messages/{message_id}/attachments/createUploadSession",
            json={"AttachmentItem": upload_session_item(attachment)},
        ).json()
        upload_url = session["uploadUrl"]
        for chunk, content_range in iter_upload_chunks(attachment.content):
            self._session.upload_chunk(upload_url, chunk, content_range)

    def _iter_messages(
        self, path: str, params: dict, top: int | None, headers: dict | None = None
    ) -> Iterator[OutlookMessage]:
        if top is not None:
            # Ask Graph for bigger pages (its default is 10 messages per page)
            # so a bounded read is one round trip, not top/10. $top is a page
            # size, not a total — islice below stays the exact bound. Capped at
            # 100, which mail endpoints accept everywhere.
            params = {**params, "$top": min(top, 100)}
        raw = self._session.paginate(path, params, headers=headers)
        if top is not None:
            raw = itertools.islice(raw, top)
        for item in raw:
            yield OutlookMessage.model_validate(item)


def _to_recipients(value: Recipients) -> list[dict]:
    if isinstance(value, (str, EmailAddress)):
        items: Iterable[RecipientSpec] = [value]
    else:
        items = value
    recipients = []
    for item in items:
        if isinstance(item, EmailAddress):
            email: dict[str, Any] = {"address": item.address}
            if item.name:
                email["name"] = item.name
        else:
            email = {"address": item}
        recipients.append({"emailAddress": email})
    return recipients


def _body_payload(body: str, html: bool) -> dict:
    return {"contentType": "html" if html else "text", "content": body}


def _build_message(
    to: Recipients,
    subject: str,
    body: str,
    cc: Recipients | None,
    bcc: Recipients | None,
    html: bool,
    attachments: list[Attachment],
) -> dict:
    message: dict[str, Any] = {
        "subject": subject,
        "body": _body_payload(body, html),
        "toRecipients": _to_recipients(to),
    }
    if cc is not None:
        message["ccRecipients"] = _to_recipients(cc)
    if bcc is not None:
        message["bccRecipients"] = _to_recipients(bcc)
    if attachments:
        message["attachments"] = [to_inline_payload(a) for a in attachments]
    return message


def _fmt_dt(value: datetime | str) -> str:
    if isinstance(value, datetime):
        # Preserve sub-second precision; Graph wants ISO-8601 UTC. Aware
        # datetimes are converted; naive ones are taken to already be UTC.
        if value.tzinfo is not None:
            value = value.astimezone(timezone.utc)
        return value.strftime("%Y-%m-%dT%H:%M:%S.%fZ")
    return value


def _build_filter_clauses(
    *,
    sender: str | None,
    subject_contains: str | None,
    since: datetime | str | None,
    until: datetime | str | None,
    since_exclusive: datetime | str | None,
    until_exclusive: datetime | str | None,
    unread: bool | None,
    has_attachments: bool | None,
) -> list[str]:
    clauses: list[str] = []
    if sender is not None:
        clauses.append(f"from/emailAddress/address eq '{sender}'")
    if subject_contains is not None:
        clauses.append(f"contains(subject,'{subject_contains}')")
    if since is not None:
        clauses.append(f"receivedDateTime ge {_fmt_dt(since)}")
    if until is not None:
        clauses.append(f"receivedDateTime le {_fmt_dt(until)}")
    if since_exclusive is not None:
        clauses.append(f"receivedDateTime gt {_fmt_dt(since_exclusive)}")
    if until_exclusive is not None:
        clauses.append(f"receivedDateTime lt {_fmt_dt(until_exclusive)}")
    if unread is not None:
        clauses.append(f"isRead eq {'false' if unread else 'true'}")
    if has_attachments is not None:
        clauses.append(f"hasAttachments eq {'true' if has_attachments else 'false'}")
    return clauses
