# Required `outlook-helper` changes (consumed by the connector MVP)

Derived from [the dependency findings](./00-dependency-verification-findings.md).
These let the connector keep using the typed `OutlookClient` API while getting
RFC-822 identity, precise threading, a forced-HTML body, strict `>` polling, and
cheap attachment metadata. After landing, cut a new tag and re-pin
`pyproject.toml`.

Graph facts that drive the design:
- `internetMessageId` **is** returned by default → only needs to be on the model.
- `internetMessageHeaders` is **not** returned by default → must be `$select`ed,
  and `$select` restricts the response to *only* the listed properties, so the
  helper must apply a **curated full select** (all modeled fields + the two new
  ones), not just `$select=internetMessageHeaders`.

---

## 1. `schemas.py` — model the identity + threading fields

```python
class InternetMessageHeader(GraphModel):
    name: str | None = None
    value: str | None = None


class OutlookMessage(GraphModel):
    ...  # existing fields
    internet_message_id: str | None = Field(default=None, alias="internetMessageId")
    internet_message_headers: list[InternetMessageHeader] = Field(
        default_factory=list, alias="internetMessageHeaders"
    )

    @property
    def in_reply_to(self) -> str | None:
        for h in self.internet_message_headers:
            if h.name and h.name.lower() == "in-reply-to":
                return h.value
        return None

    @property
    def references(self) -> list[str]:
        for h in self.internet_message_headers:
            if h.name and h.name.lower() == "references":
                return h.value.split() if h.value else []
        return []
```

`internet_message_id` will start populating immediately (default-returned).
`internet_message_headers` only populates when `$select`ed (change 4).

## 2. `http.py` — thread a `headers` arg through reads

`paginate`/`get_json` currently can't pass headers, so the `Prefer` header can't
reach paginated reads. Add an optional `headers` param:

```python
def get_json(self, path, params=None, *, headers=None) -> dict:
    return self.request("GET", path, params=params, headers=headers).json()

def paginate(self, path, params=None, *, headers=None) -> Iterator[dict]:
    page = self.get_json(path, params, headers=headers)
    while True:
        yield from page.get("value", [])
        next_link = page.get("@odata.nextLink")
        if not next_link:
            return
        page = self.get_json(next_link, headers=headers)  # nextLink carries its own query
```

## 3 & 4. `client.py` — strict operators, sub-second precision, HTML body, full select

```python
# module-level: every modeled property, so $select doesn't drop defaults
_MESSAGE_SELECT = (
    "id,subject,from,toRecipients,ccRecipients,bccRecipients,"
    "receivedDateTime,sentDateTime,bodyPreview,body,isRead,hasAttachments,"
    "importance,webLink,conversationId,parentFolderId,"
    "internetMessageId,internetMessageHeaders"
)

def _fmt_dt(value: datetime | str) -> str:
    if isinstance(value, datetime):
        # preserve sub-second precision; Graph wants ISO-8601 UTC
        return value.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ")
    return value
```

Add strict-comparison + fidelity options to `search_email` (and pass them to
`_iter_messages` / `_build_filter_clauses`):

```python
def search_email(
    self, *,
    sender=None, subject_contains=None,
    since=None, until=None,
    since_exclusive=None,          # NEW -> receivedDateTime gt
    until_exclusive=None,          # NEW -> receivedDateTime lt
    unread=None, has_attachments=None,
    folder=None, top=None,
    include_headers: bool = False, # NEW -> apply _MESSAGE_SELECT
    html_body: bool = False,       # NEW -> Prefer HTML
) -> Iterator[OutlookMessage]:
    ...
    params = {"$orderby": "receivedDateTime desc"}  # connector overrides to asc via its own sort
    if clauses:
        params["$filter"] = " and ".join(clauses)
    if include_headers:
        params["$select"] = _MESSAGE_SELECT
    headers = {"Prefer": 'outlook.body-content-type="html"'} if html_body else None
    return self._iter_messages(path, params, top=top, headers=headers)
```

In `_build_filter_clauses`, add the strict operators:

```python
if since_exclusive is not None:
    clauses.append(f"receivedDateTime gt {_fmt_dt(since_exclusive)}")
if until_exclusive is not None:
    clauses.append(f"receivedDateTime lt {_fmt_dt(until_exclusive)}")
```

`_iter_messages` gains `headers=None` and forwards it to `self._session.paginate`.
Apply the same `include_headers` / `html_body` options to `get_email` and
`list_messages` for consistency.

## 5. `client.py` — cheap attachment metadata

`list_attachments` currently pulls full `contentBytes`. Restrict it:

```python
def list_attachments(self, message_id: str) -> list[OutlookAttachmentMeta]:
    items = self._session.paginate(
        f"{self._base_path}/messages/{message_id}/attachments",
        {"$select": "id,name,contentType,size,isInline"},
    )
    return [OutlookAttachmentMeta.model_validate(item) for item in items]
```

## Not changing in the helper

- **Async.** The helper stays synchronous; the connector wraps every call in
  `asyncio.to_thread(...)` so the event loop (and `Retry-After` sleeps) don't
  block it. Making `outlook-helper` async is a much larger lift and out of scope.
- **Retry scope.** 429/503-only retry is fine; the connector catches both
  `GraphError` and `httpx` exceptions per-mailbox, as the spec prescribes.

## Connector usage after these land

```python
msgs = await asyncio.to_thread(
    lambda: list(client.search_email(
        since_exclusive=cursor,     # strict > , sub-second precise
        include_headers=True,       # internet_message_id + in_reply_to/references
        html_body=True,             # body guaranteed HTML
    ))
)
msgs.sort(key=lambda m: m.received_at)   # ascending, per design
```
