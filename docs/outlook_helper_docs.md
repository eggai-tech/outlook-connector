# outlook-helper — Library Reference

A Python library for working with Microsoft 365 email through the Microsoft Graph API.
This single file is the complete reference for **using `outlook-helper` from another
service** (e.g. a connector, agent, or daemon). Everything you need — auth setup,
the full client API, data models, error handling, and worked examples — is here.

- **Package:** `outlook_helper`
- **Python:** 3.13+
- **Transport:** Microsoft Graph `v1.0` over HTTPS (via `httpx`)
- **Auth:** MSAL (delegated device-code flow, or app-only client credentials)
- **Models:** Pydantic v2

---

## 1. Installation & dependencies

The library is distributed as the `outlook-helper` package. In the consuming
service, add it as a dependency and import from the top-level `outlook_helper`
package.

```bash
uv sync          # or: pip install outlook-helper
```

Runtime dependencies (pulled in automatically): `httpx`, `msal`,
`msal-extensions`, `pydantic`, `click`.

Everything you normally need is exported from the package root:

```python
from outlook_helper import (
    OutlookClient,
    # config
    DelegatedConfig, AppOnlyConfig, CertificateConfig,
    # credentials
    DeviceCodeCredential, ClientSecretCredential, Credential,
    # models
    OutlookMessage, OutlookBody, EmailAddress,
    OutlookAttachmentMeta, OutlookFolder, Attachment,
    # errors
    GraphError,
)
```

---

## 2. Mental model

```
Config  ──►  Credential  ──►  OutlookClient  ──►  Graph (HTTPS)
(dataclass)  (mints tokens)   (one per mailbox)
```

1. Build a **config** dataclass describing your Azure AD app and auth model.
2. Wrap it in a **credential** that mints bearer tokens.
3. Construct an **`OutlookClient`** bound to **one mailbox**.
4. Call verbs (`get_email`, `send_email`, `search_email`, …). The client maps
   Graph JSON to Pydantic models and raises `GraphError` on failures.

The client is **synchronous** and bound to a single mailbox. For multiple
mailboxes, create one client per mailbox.

---

## 3. Authentication

Two auth models are supported behind a common `Credential` protocol. The rest of
the library never branches on which one you use.

### 3.1 Delegated (signed-in user) — device-code flow

Use when acting **as a user**. On first token request, a device-code message is
printed to stderr (or routed to your callback); the user visits the URL and
enters the code. Tokens are then cached and refreshed silently.

```python
from pathlib import Path
from outlook_helper import OutlookClient, DelegatedConfig, DeviceCodeCredential

credential = DeviceCodeCredential(
    DelegatedConfig(
        client_id="<app-client-id>",
        tenant_id="<tenant-id-or-'common'>",   # default: "common"
        cache_path=Path("~/.cache/outlook/token.bin").expanduser(),  # optional
    )
)
client = OutlookClient(credential)   # no mailbox needed → uses the signed-in user (/me)
```

**`DelegatedConfig` fields**

| Field        | Type                 | Default                                          | Notes                                              |
|--------------|----------------------|--------------------------------------------------|----------------------------------------------------|
| `client_id`  | `str`                | required                                         | Azure AD app (public client) id.                   |
| `tenant_id`  | `str`                | `"common"`                                       | Tenant id, or `"common"`/`"organizations"`.        |
| `scopes`     | `tuple[str, ...]`    | `("Mail.ReadWrite", "Mail.Send", "User.Read")`   | Delegated Graph scopes.                            |
| `cache_path` | `Path \| None`       | `None`                                           | Where to persist the token cache (see below).      |

**Token cache behaviour** (`cache_path`):
- `None` → in-memory only; the user must re-authenticate every process run.
- A path → an **encrypted** persistent cache (keyring-backed via
  `msal-extensions`) when available; otherwise it falls back to a **plaintext**
  file cache and emits a `RuntimeWarning`.

**Custom device-code prompt.** By default the prompt message is printed to
stderr. To surface it elsewhere (logs, a UI, a chat message), pass a callback:

```python
def on_prompt(flow: dict) -> None:
    # flow contains: "message", "user_code", "verification_uri", "expires_in", ...
    send_to_user(flow["message"])

credential = DeviceCodeCredential(DelegatedConfig(client_id="..."),
                                  prompt_callback=on_prompt)
```

### 3.2 App-only (daemon) — client credentials

Use when acting **as the application** with no signed-in user (background
services, shared mailboxes). There is **no `/me`**, so you **must** pass a
`mailbox` to the client.

```python
from outlook_helper import OutlookClient, AppOnlyConfig, ClientSecretCredential

credential = ClientSecretCredential(
    AppOnlyConfig(
        client_id="<app-client-id>",
        tenant_id="<tenant-id>",
        client_secret="<secret>",
    )
)
client = OutlookClient(credential, mailbox="shared@example.com")  # mailbox REQUIRED
```

**`AppOnlyConfig` fields**

| Field           | Type                       | Default                                              | Notes                                          |
|-----------------|----------------------------|------------------------------------------------------|------------------------------------------------|
| `client_id`     | `str`                      | required                                             |                                                |
| `tenant_id`     | `str`                      | required                                             |                                                |
| `client_secret` | `str \| None`              | `None`                                               | Provide **this or** `certificate`, not both.   |
| `certificate`   | `CertificateConfig \| None`| `None`                                               | Provide **this or** `client_secret`, not both. |
| `scopes`        | `tuple[str, ...]`          | `("https://graph.microsoft.com/.default",)`          | App-only uses `.default`.                      |

> `AppOnlyConfig` raises `ValueError` at construction if you supply neither or
> both of `client_secret` / `certificate`.

**Certificate credential** (instead of a secret):

```python
from outlook_helper import AppOnlyConfig, CertificateConfig, ClientSecretCredential

config = AppOnlyConfig(
    client_id="...",
    tenant_id="...",
    certificate=CertificateConfig(
        private_key_path="/secrets/key.pem",
        thumbprint="ABCD1234...",
        public_certificate_path="/secrets/cert.pem",  # optional (for SNI / x5c)
    ),
)
credential = ClientSecretCredential(config)
```

### 3.3 The `Credential` protocol

Both credential classes satisfy this protocol. You can supply your own.

```python
@runtime_checkable
class Credential(Protocol):
    supports_me: bool          # True only for delegated (enables the /me shortcut)
    def get_token(self) -> str: ...   # returns a valid bearer token
```

- `ClientSecretCredential.supports_me == False`
- `DeviceCodeCredential.supports_me == True`

Token acquisition is **lazy**: the underlying MSAL app and any network/disk I/O
happen on the first `get_token()` call, not at construction. A failed token
acquisition raises `GraphError(status_code=401, ...)`.

To pre-authenticate (e.g. trigger the device-code prompt up front):

```python
client.credential.get_token()   # forces sign-in / cache priming now
```

---

## 4. Constructing the client

```python
OutlookClient(
    credential: Credential,
    mailbox: str | None = None,
    *,
    session: GraphSession | None = None,
    base_url: str = "https://graph.microsoft.com/v1.0",
)
```

| Param        | Meaning                                                                                             |
|--------------|-----------------------------------------------------------------------------------------------------|
| `credential` | A `Credential` (delegated or app-only).                                                             |
| `mailbox`    | Target mailbox (UPN/email). If omitted, requires a delegated credential and targets `/me`.          |
| `session`    | Inject a pre-built `GraphSession` (testing / shared `httpx.Client` / custom retry config).          |
| `base_url`   | Graph base URL. Override for sovereign clouds (e.g. GCC High, China).                               |

**Mailbox resolution rules:**
- `mailbox="user@org.com"` → all calls target `/users/user@org.com`.
- `mailbox=None` + delegated credential → targets `/me` (the signed-in user).
- `mailbox=None` + app-only credential → **raises `ValueError`** (no signed-in
  user exists).

Read-only properties: `client.base_path` (e.g. `/me` or `/users/...`) and
`client.credential`.

---

## 5. Client API reference

All methods raise `GraphError` on any non-2xx Graph response (after retries).
Folder arguments accept a well-known name, a folder display name, or a raw
folder id (see §7).

### 5.1 Reading

#### `get_email(message_id: str) -> OutlookMessage`
Fetch a single message by id.

```python
msg = client.get_email("AAMk...")
print(msg.subject, msg.from_.address, msg.body.content)
```

#### `list_messages(folder="inbox", *, top=None) -> Iterator[OutlookMessage]`
List messages in a folder, **newest first** (`receivedDateTime desc`).

- Returns a **lazy iterator** that follows Graph pagination automatically.
- `top` caps the total number of messages yielded across all pages (`None` = all).

```python
for msg in client.list_messages(folder="inbox", top=50):
    print(msg.id, msg.subject)
```

#### `search_email(*, sender=None, subject_contains=None, since=None, until=None, unread=None, has_attachments=None, folder=None, top=None) -> Iterator[OutlookMessage]`
Search with precise server-side filters. **All arguments are keyword-only.**
Returns a lazy iterator, newest first.

| Filter             | Type                      | Graph clause produced                                  |
|--------------------|---------------------------|--------------------------------------------------------|
| `sender`           | `str` (email)             | `from/emailAddress/address eq '<sender>'` (exact match)|
| `subject_contains` | `str`                     | `contains(subject,'<text>')`                           |
| `since`            | `datetime \| str`         | `receivedDateTime ge <iso>`                            |
| `until`            | `datetime \| str`         | `receivedDateTime le <iso>`                            |
| `unread`           | `bool`                    | `isRead eq false` (True) / `isRead eq true` (False)    |
| `has_attachments`  | `bool`                    | `hasAttachments eq true/false`                         |
| `folder`           | `str`                     | scopes to that folder; `None` searches the whole mailbox |
| `top`              | `int`                     | caps total results                                     |

- `datetime` values are formatted as `%Y-%m-%dT%H:%M:%SZ` (UTC, no offset). Pass
  timezone-aware/UTC datetimes for correctness, or pass a pre-formatted ISO 8601
  `str` to control the format yourself.
- `sender` is an **exact** address match, not a substring.
- Omitted filters are simply not constrained. With no filters, this lists the
  mailbox (or `folder`) newest-first.

```python
from datetime import datetime, timezone

results = client.search_email(
    subject_contains="invoice",
    sender="billing@vendor.com",
    since=datetime(2026, 1, 1, tzinfo=timezone.utc),
    has_attachments=True,
    unread=True,
    folder="inbox",
    top=100,
)
for msg in results:
    ...
```

#### `list_attachments(message_id: str) -> list[OutlookAttachmentMeta]`
Return **metadata** for all attachments on a message (not the bytes). Use
`download_attachment` to fetch content.

```python
for att in client.list_attachments(msg_id):
    print(att.id, att.name, att.content_type, att.size, att.is_inline)
```

#### `download_attachment(message_id, attachment_id, dest_path) -> Path`
Stream an attachment's bytes to `dest_path` (a `str` or `Path`) and return the
written `Path`. Streams to disk in chunks — safe for large files.

```python
for att in client.list_attachments(msg_id):
    path = client.download_attachment(msg_id, att.id, f"./downloads/{att.name}")
    print("saved", path)
```

### 5.2 Sending

#### `send_email(to, subject, body, *, cc=None, bcc=None, attachments=None, html=False) -> None`
Send a new message immediately (also saved to Sent Items). Returns `None`.

- `to` / `cc` / `bcc`: a **recipient spec** — see §6. (A single string, an
  `EmailAddress`, or an iterable of either.)
- `body`: plain text by default; set `html=True` to send `body` as HTML.
- `attachments`: an iterable of **attachment specs** — see §8. Large attachments
  (> 3 MB) are automatically sent via a draft + chunked upload session.

```python
client.send_email(
    to=["a@x.com", "b@x.com"],
    subject="Report",
    body="<h1>Q1</h1><p>See attached.</p>",
    cc="manager@x.com",
    html=True,
    attachments=["/data/q1.pdf", "/data/big-video.mp4"],  # >3MB handled automatically
)
```

#### `reply(message_id, body, *, reply_all=False, attachments=None, html=False) -> None`
Reply to a message (preserving the original via Graph's `createReply` /
`createReplyAll`), then send. Returns `None`.

```python
client.reply(msg_id, "Thanks — received.", reply_all=True)
client.reply(msg_id, "<p>See attached.</p>", html=True, attachments=["/tmp/a.pdf"])
```

### 5.3 Drafts

#### `create_draft(to, subject, body, *, cc=None, bcc=None, attachments=None, html=False) -> OutlookMessage`
Create a draft **without sending**. Returns the created `OutlookMessage` — use
its `.id` to update or send later. Large attachments are uploaded via a session.

#### `update_draft(message_id, *, subject=None, body=None, to=None, cc=None, bcc=None, html=False) -> OutlookMessage`
Patch fields on an existing draft. Only the provided fields are changed. Returns
the updated `OutlookMessage`.

> ⚠️ `html` defaults to `False`. If you pass `body=` to update a draft that should
> remain HTML, also pass `html=True`, otherwise the body is stored as plain text.

#### `send_draft(message_id: str) -> None`
Send a previously created/updated draft.

#### `discard_draft(message_id: str) -> None`
Delete a draft (hard delete via `DELETE`).

```python
draft = client.create_draft("a@x.com", "Subject", "Initial body")
draft = client.update_draft(draft.id, subject="Revised", body="<p>Final</p>", html=True)
client.send_draft(draft.id)
# or: client.discard_draft(draft.id)
```

### 5.4 Folders & message lifecycle

#### `list_folders() -> list[OutlookFolder]`
List the mailbox's mail folders.

#### `create_folder(name, *, parent=None) -> OutlookFolder`
Create a folder, optionally nested under `parent` (a folder reference). Returns
the new `OutlookFolder`. The internal folder-name cache is invalidated so the
new folder is resolvable by name immediately.

#### `move_email(message_id, dest_folder) -> OutlookMessage`
Move a message to another folder. `dest_folder` is a folder reference. Returns
the moved message (note: Graph assigns it a **new id** in the destination).

#### `delete_email(message_id, *, permanent=False) -> None`
Delete a message. By default this is a **soft delete** (moves to Deleted Items).
With `permanent=True` it is irrecoverable (Graph `permanentDelete`).

```python
folders = client.list_folders()
archive = client.create_folder("Processed", parent="Archive")
moved = client.move_email(msg_id, "Processed")
client.delete_email(old_id, permanent=False)
```

---

## 6. Recipient specs

Anywhere a recipient is accepted (`to`, `cc`, `bcc`), you may pass:

- a **string** email address: `"a@x.com"`
- an **`EmailAddress`**: `EmailAddress(address="a@x.com", name="Alice")`
- an **iterable** of either: `["a@x.com", EmailAddress(address="b@x.com")]`

```python
from outlook_helper import EmailAddress

client.send_email(
    to=[EmailAddress(address="a@x.com", name="Alice"), "b@x.com"],
    subject="Hi", body="...",
)
```

---

## 7. Folder references

Every folder argument (`folder=`, `dest_folder`, `parent=`) accepts **three
forms**, resolved in this order:

1. **Well-known name** (case-insensitive) — passed straight to Graph:
   `archive`, `clutter`, `conflicts`, `conversationhistory`, `deleteditems`,
   `drafts`, `inbox`, `junkemail`, `localfailures`, `msgfolderroot`, `outbox`,
   `recoverableitemsdeletions`, `scheduled`, `searchfolders`, `sentitems`,
   `serverfailures`, `syncissues`.
2. **Display name** — looked up via a filtered query and **cached** for the life
   of the client (e.g. `"Archive"`, `"My Project"`). First match wins.
3. **Raw folder id** — if it is neither a well-known name nor a matching display
   name, it is assumed to already be a folder id and passed through.

```python
client.list_messages(folder="inbox")            # well-known
client.list_messages(folder="Project Alpha")    # display name → resolved + cached
client.list_messages(folder="AAMkAGI1...")       # raw id
```

---

## 8. Attachments

### Specs accepted by `attachments=`
An attachment spec (`AttachmentSpec`) is one of:

- a **path** (`str` or `pathlib.Path`) — read from disk; name and content type
  are inferred (content type via `mimetypes`, defaulting to
  `application/octet-stream`).
- an **`Attachment`** — in-memory bytes you build yourself:

```python
from outlook_helper import Attachment

att = Attachment(
    name="report.csv",
    content=b"a,b,c\n1,2,3\n",
    content_type="text/csv",   # default: "application/octet-stream"
)
client.send_email("a@x.com", "Data", "see attached", attachments=[att])
```

### Small vs. large (handled automatically)
- **≤ 3 MB**: sent inline (base64) on the message in a single request.
- **> 3 MB** (`LARGE_ATTACHMENT_THRESHOLD = 3 * 1024 * 1024`): the library creates
  a draft, opens a Graph **upload session**, and streams the file in ~1.6 MB
  chunks, then sends. You don't need to do anything special — just pass the file.

---

## 9. Data models

All models are Pydantic v2 (`from outlook_helper import ...`). Field names are
snake_case; Graph's `{"emailAddress": {...}}` envelope for senders/recipients is
flattened to `EmailAddress`. Unknown Graph fields are ignored.

### `OutlookMessage`
| Field              | Type                  | Notes                                            |
|--------------------|-----------------------|--------------------------------------------------|
| `id`               | `str`                 | Message id (changes when moved).                 |
| `subject`          | `str \| None`         |                                                  |
| `from_`            | `EmailAddress \| None`| Sender. **Note the trailing underscore.**        |
| `to`               | `list[EmailAddress]`  | Defaults to `[]`.                                |
| `cc`               | `list[EmailAddress]`  | Defaults to `[]`.                                |
| `bcc`              | `list[EmailAddress]`  | Defaults to `[]`.                                |
| `received_at`      | `datetime \| None`    | Graph `receivedDateTime`.                        |
| `sent_at`          | `datetime \| None`    | Graph `sentDateTime`.                            |
| `body_preview`     | `str \| None`         | Short text snippet.                              |
| `body`             | `OutlookBody \| None` | Full body (content + content type).              |
| `is_read`          | `bool \| None`        |                                                  |
| `has_attachments`  | `bool`                | Defaults to `False`.                             |
| `importance`       | `str \| None`         | `"low"`/`"normal"`/`"high"`.                     |
| `web_link`         | `str \| None`         | Open-in-Outlook URL.                             |
| `conversation_id`  | `str \| None`         |                                                  |
| `parent_folder_id` | `str \| None`         |                                                  |

### `EmailAddress`
| Field     | Type           |
|-----------|----------------|
| `name`    | `str \| None`  |
| `address` | `str \| None`  |

### `OutlookBody`
| Field          | Type           | Notes                          |
|----------------|----------------|--------------------------------|
| `content_type` | `str \| None`  | `"text"` or `"html"`.          |
| `content`      | `str \| None`  | The body content.              |

### `OutlookAttachmentMeta`
| Field          | Type           | Notes                  |
|----------------|----------------|------------------------|
| `id`           | `str`          | Use with `download_attachment`. |
| `name`         | `str \| None`  |                        |
| `content_type` | `str \| None`  |                        |
| `size`         | `int \| None`  | Bytes.                 |
| `is_inline`    | `bool`         | Defaults to `False`.   |

### `OutlookFolder`
| Field               | Type           |
|---------------------|----------------|
| `id`                | `str`          |
| `display_name`      | `str \| None`  |
| `parent_folder_id`  | `str \| None`  |
| `child_folder_count`| `int \| None`  |
| `total_item_count`  | `int \| None`  |
| `unread_item_count` | `int \| None`  |

### `Attachment` (input model)
| Field          | Type    | Default                       |
|----------------|---------|-------------------------------|
| `name`         | `str`   | required                      |
| `content`      | `bytes` | required                      |
| `content_type` | `str`   | `"application/octet-stream"`  |

---

## 10. Error handling

Every failure surfaces as a single exception type.

### `GraphError(Exception)`
Raised for any non-2xx Graph response (after retries) and for auth failures.

| Attribute     | Type          | Meaning                                            |
|---------------|---------------|----------------------------------------------------|
| `status_code` | `int`         | HTTP status (e.g. 401, 403, 404, 429, 500).        |
| `message`     | `str`         | Human-readable message (from Graph when available).|
| `code`        | `str \| None` | Graph error code (e.g. `ErrorItemNotFound`).       |
| `request_id`  | `str \| None` | Originating request id — quote it in support cases.|

`str(err)` renders as `"[<status_code>] <message>"`.

```python
from outlook_helper import GraphError

try:
    msg = client.get_email(message_id)
except GraphError as err:
    if err.status_code == 404:
        ...  # not found
    elif err.status_code in (401, 403):
        ...  # auth / permission problem
    else:
        log.error("Graph %s %s (req %s)", err.status_code, err.code, err.request_id)
        raise
```

### Automatic retries
The HTTP layer retries `429` (throttling) and `503` responses up to **3 times**,
honouring the `Retry-After` header (falling back to exponential backoff:
`2**attempt` seconds). If retries are exhausted, the final response is raised as
`GraphError`. To customise, inject your own `GraphSession` (see §11).

---

## 11. Advanced: customising the HTTP session

The client owns a `GraphSession` (the single seam to Graph). Inject your own to
share an `httpx.Client`, change retry behaviour, or target a non-default base URL.

```python
import httpx
from outlook_helper import OutlookClient
from outlook_helper.http import GraphSession

session = GraphSession(
    credential,
    base_url="https://graph.microsoft.com/v1.0",
    max_retries=5,
    client=httpx.Client(timeout=30.0),
)
client = OutlookClient(credential, mailbox="x@org.com", session=session)
...
session.close()   # close the underlying httpx client when done
```

`GraphSession` exposes low-level helpers if you need raw Graph access not covered
by the client verbs: `request(method, path, ...)`, `get_json(path, params)`,
`paginate(path, params)` (lazy, follows `@odata.nextLink`), `download(path, dest)`,
and `upload_chunk(url, data, content_range)`.

---

## 12. Behavioural notes & gotchas

- **Lazy iterators.** `list_messages` and `search_email` return generators that
  page on demand. Nothing is fetched until you iterate. Wrap in `list(...)` if you
  need the full result eagerly, and use `top=` to bound large mailboxes.
- **`from_` not `from`.** The sender field is `msg.from_` (Python keyword clash).
  It may be `None`; guard with `msg.from_.address if msg.from_ else None`.
- **`mailbox` is required for app-only.** Without a signed-in user there is no
  `/me`; constructing the client without a mailbox raises `ValueError`.
- **HTML vs text.** `html=False` is the default everywhere. For HTML bodies pass
  `html=True` — including on `update_draft`, or the body downgrades to plain text.
- **Move changes the id.** After `move_email`, use the returned message's `.id`.
- **`send_email` / `reply` / `send_draft` / `discard_draft` / `delete_email` return `None`.**
  Only `get_email` / `create_draft` / `update_draft` / `move_email` /
  `create_folder` return models; `list_*` return lists or iterators.
- **Datetimes.** Pass UTC/aware `datetime`s to `search_email`; naive ones are
  formatted as-is with a trailing `Z`.
- **Thread-safety.** A client/session wraps one `httpx.Client`; use one client per
  thread, or inject a shared, thread-safe `httpx.Client`.
- **Required Graph permissions.** Delegated: `Mail.ReadWrite`, `Mail.Send`,
  `User.Read`. App-only: application permissions `Mail.ReadWrite` and `Mail.Send`
  (admin-consented), accessed via the `.default` scope.

---

## 13. End-to-end examples

### 13.1 Delegated: triage unread mail with attachments

```python
from outlook_helper import OutlookClient, DelegatedConfig, DeviceCodeCredential, GraphError

client = OutlookClient(DeviceCodeCredential(DelegatedConfig(client_id="...")))

try:
    for msg in client.search_email(unread=True, has_attachments=True, top=25):
        print(msg.subject, "from", msg.from_.address if msg.from_ else "?")
        for att in client.list_attachments(msg.id):
            client.download_attachment(msg.id, att.id, f"./inbox/{att.name}")
        client.move_email(msg.id, "Processed")   # display name; create it first if needed
except GraphError as e:
    print("Graph failure:", e)
```

### 13.2 App-only daemon: send a report from a shared mailbox

```python
from outlook_helper import OutlookClient, AppOnlyConfig, ClientSecretCredential, Attachment

client = OutlookClient(
    ClientSecretCredential(AppOnlyConfig(client_id="...", tenant_id="...", client_secret="...")),
    mailbox="reports@example.com",
)

client.send_email(
    to=["team@example.com"],
    cc="lead@example.com",
    subject="Nightly report",
    body="<p>Attached.</p>",
    html=True,
    attachments=[Attachment(name="report.csv", content=csv_bytes, content_type="text/csv")],
)
```

### 13.3 Draft → review → send

```python
draft = client.create_draft("client@acme.com", "Proposal", "Draft text")
# ...human or agent reviews draft.id...
client.update_draft(draft.id, body="<p>Final proposal</p>", html=True)
client.send_draft(draft.id)
```

---

## 14. CLI (optional)

The library also ships a thin CLI (`outlook-helper`) over the same client, useful
for smoke-testing credentials without writing code. Config comes from options or
the env vars `OUTLOOK_CLIENT_ID`, `OUTLOOK_TENANT_ID`, `OUTLOOK_MAILBOX`,
`OUTLOOK_CLIENT_SECRET`, `OUTLOOK_CACHE_PATH`.

```bash
outlook-helper --client-id <id> login
outlook-helper --client-id <id> list --folder inbox --top 20
outlook-helper --client-id <id> search --subject-contains invoice --has-attachments
outlook-helper --client-id <id> send --to a@x.com --subject Hi --body "Hello" --attach ./f.pdf
outlook-helper --client-id <id> reply <message-id> --body "Thanks" --reply-all
outlook-helper --client-id <id> draft --to a@x.com --subject Hi --body "..."
outlook-helper --client-id <id> send-draft <message-id>
outlook-helper --client-id <id> download <message-id> <attachment-id> ./out.bin
outlook-helper --client-id <id> folders
outlook-helper --client-id <id> mkdir "Processed" --parent Archive
outlook-helper --client-id <id> move <message-id> Archive
outlook-helper --client-id <id> delete <message-id> --permanent
```

For app-only auth add `--auth app-only --tenant-id <t> --client-secret <s>`. Run
`outlook-helper --help` for the full command list.

---

## 15. Quick API index

| Verb | Signature (return) |
|------|--------------------|
| Read one | `get_email(message_id) -> OutlookMessage` |
| List | `list_messages(folder="inbox", *, top=None) -> Iterator[OutlookMessage]` |
| Search | `search_email(*, sender, subject_contains, since, until, unread, has_attachments, folder, top) -> Iterator[OutlookMessage]` |
| List attachments | `list_attachments(message_id) -> list[OutlookAttachmentMeta]` |
| Download attachment | `download_attachment(message_id, attachment_id, dest_path) -> Path` |
| Send | `send_email(to, subject, body, *, cc, bcc, attachments, html) -> None` |
| Reply | `reply(message_id, body, *, reply_all, attachments, html) -> None` |
| Create draft | `create_draft(to, subject, body, *, cc, bcc, attachments, html) -> OutlookMessage` |
| Update draft | `update_draft(message_id, *, subject, body, to, cc, bcc, html) -> OutlookMessage` |
| Send draft | `send_draft(message_id) -> None` |
| Discard draft | `discard_draft(message_id) -> None` |
| List folders | `list_folders() -> list[OutlookFolder]` |
| Create folder | `create_folder(name, *, parent=None) -> OutlookFolder` |
| Move | `move_email(message_id, dest_folder) -> OutlookMessage` |
| Delete | `delete_email(message_id, *, permanent=False) -> None` |
