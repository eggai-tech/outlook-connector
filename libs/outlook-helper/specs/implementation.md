# outlook-helper — Implementation Plan

A synchronous Python library (plus a thin CLI) for working with Microsoft 365
email through the Microsoft Graph API. Built lean on `msal` (auth + token
caching) and `httpx` (HTTP), with Pydantic models as the return types.

---

## 1. Locked decisions

| Area | Decision |
| --- | --- |
| Auth model | Both **delegated** (user sign-in) and **app-only** (client credentials), configurable via a pluggable credential |
| Delegated flow | **Device code** flow (headless-friendly) |
| Token cache | Encrypted file cache via `msal-extensions` (OS keyring-backed), fallback to plain file when keyring is unavailable |
| Graph client | `msal` + `httpx`, hand-written request/response handling |
| Concurrency | **Sync only** |
| Return types | **Pydantic** models |
| Target mailbox | Fixed **per-client** at construction (delegated defaults to `/me`; app-only requires a mailbox) |
| Config input | **Explicit config object** passed in code (library stays env-agnostic) |
| Errors | Single **`GraphError`** exception (status code + Graph error code + message + request id); auto-retry on 429/503 |
| Delete | **Soft** delete by default, `permanent=True` flag for hard delete |
| Drafts | **Public** draft lifecycle (create / update / send / discard) |
| Folders | Addressable by **well-known name, display name (resolved + cached), or raw ID** |
| Attachments | Small inline (<3 MB) + **chunked upload sessions** for large files |
| Pagination | **Lazy auto-paging iterator** following `@odata.nextLink` |
| Search | **Precise structured filters (`$filter`)** initially — subject, date range, sender, has-attachments, folder. Free-text `$search` **deferred** to a future iteration |
| Read/unread, flags, categories, calendar, contacts, batching | **Out of scope** (not in README) |
| Scope | Library + thin CLI |
| Testing | Unit tests with mocked HTTP (`pytest` + `respx`), recorded Graph fixtures |
| Runtime | Python 3.13, managed with `uv` |

---

## 2. Dependencies

Runtime:
- `httpx` — HTTP client
- `msal` — token acquisition (device code + client credentials)
- `msal-extensions` — encrypted, persistent token cache
- `pydantic` (v2) — data models
- `click` — CLI framework

Dev:
- `pytest` — test runner
- `respx` — httpx mock/transport for unit tests

---

## 3. Module layout

```
outlook_helper/
  __init__.py        # public exports: OutlookClient, configs, models, GraphError
  config.py          # DelegatedConfig, AppOnlyConfig
  auth.py            # Credential protocol + DeviceCodeCredential, ClientSecretCredential; token cache
  http.py            # GraphSession: token injection, error mapping, retry, pagination
  folders.py         # FolderResolver: well-known / display-name / id resolution + cache
  attachments.py     # attachment input handling; inline encode + chunked upload sessions
  models.py          # Pydantic models
  exceptions.py      # GraphError
  client.py          # OutlookClient — the public verbs
  cli.py             # click CLI over OutlookClient
docs/
  DESIGN.md  # this file
tests/
  conftest.py        # fixtures, respx setup, recorded Graph payloads
  fixtures/          # recorded Graph JSON responses
  test_*.py          # per-module tests
```

`click` is used for the CLI.

---

## 4. Configuration (`config.py`)

```python
@dataclass(frozen=True)
class DelegatedConfig:
    client_id: str
    tenant_id: str = "common"           # or a specific tenant GUID/domain
    scopes: tuple[str, ...] = ("Mail.ReadWrite", "Mail.Send", "User.Read")
    cache_path: Path | None = None      # where the encrypted token cache lives

@dataclass(frozen=True)
class AppOnlyConfig:
    client_id: str
    tenant_id: str
    client_secret: str | None = None    # secret OR certificate
    certificate: CertificateConfig | None = None
    scopes: tuple[str, ...] = ("https://graph.microsoft.com/.default",)
```

`offline_access` is added implicitly for delegated refresh tokens. Validation:
exactly one of `client_secret` / `certificate` for app-only.

---

## 5. Authentication (`auth.py`)

A minimal credential protocol so the rest of the library never branches on auth
mode:

```python
class Credential(Protocol):
    def get_token(self) -> str: ...          # valid bearer token (refreshes as needed)
    @property
    def default_mailbox_path(self) -> str: ...  # "/me" or "/users/{mailbox}"
```

- **`DeviceCodeCredential(DelegatedConfig)`** — MSAL `PublicClientApplication`.
  First call triggers the device-code prompt (user code + verification URL
  surfaced via a callback, default prints to stderr). Tokens cached and silently
  refreshed on subsequent calls. `default_mailbox_path == "/me"`.
- **`ClientSecretCredential(AppOnlyConfig)`** — MSAL `ConfidentialClientApplication`
  client-credentials grant. App-only has no `/me`, so a mailbox is **required**
  on the client; `default_mailbox_path` is unused.

**Token cache:** `msal-extensions` builds an encrypted, OS-keyring-backed
persistence at `cache_path`; if the platform keyring is unavailable, fall back
to unencrypted `FilePersistence` with a logged warning. In-memory only when
`cache_path is None`.

---

## 6. HTTP session (`http.py`)

`GraphSession` wraps a single `httpx.Client` and is the only place that talks to
Graph.

Responsibilities:
- Prepend base URL (`https://graph.microsoft.com/v1.0`) and inject
  `Authorization: Bearer <token>` from the credential on every request.
- **Error mapping:** any non-2xx → `GraphError` carrying `status_code`, Graph
  `code`, `message`, and `request_id` (from the `x-ms-request-id` /
  `request-id` header).
- **Retry:** on `429` and `503`, honor `Retry-After`; otherwise exponential
  backoff. Configurable max attempts (default 3).
- **Pagination:** `paginate(path, params) -> Iterator[dict]` lazily follows
  `@odata.nextLink`, yielding raw items one page at a time. The client layer
  maps each item to a model, so callers iterate models without loading the full
  result set.
- **Streaming:** `stream_get(path) -> Iterator[bytes]` for large attachment
  downloads.

---

## 7. Folder resolution (`folders.py`)

`FolderResolver` turns a user-facing folder reference into something Graph can
address, improving on raw-ID-only approaches:

1. **Well-known names** (`inbox`, `drafts`, `sentitems`, `deleteditems`,
   `junkemail`, `archive`, …) → used directly in the Graph path, no lookup.
2. **Display names** → `GET /mailFolders?$filter=displayName eq '...'`
   (including child folders), result **cached** on the client.
3. **Raw IDs** → passed through unchanged.

Used by `list_messages`, `search_email`, `move_email`, and `create_folder`'s
`parent`.

---

## 8. Data models (`models.py`)

Pydantic v2 models, populated from Graph JSON via aliases:

- `EmailAddress` — `name: str | None`, `address: str`
- `Recipient` — wraps `EmailAddress` (Graph nests as `emailAddress`)
- `Body` — `content_type: Literal["text", "html"]`, `content: str`
- `Message` — `id`, `subject`, `from_`, `to`, `cc`, `bcc`, `received_at`,
  `sent_at`, `body_preview`, `body`, `is_read`, `has_attachments`,
  `importance`, `web_link`, `conversation_id`, `parent_folder_id`
- `AttachmentMeta` — `id`, `name`, `content_type`, `size`, `is_inline`
- `Folder` — `id`, `display_name`, `parent_folder_id`, `child_folder_count`,
  `total_item_count`, `unread_item_count`

Input helpers (not Graph-shaped): recipients accept `str | EmailAddress`;
attachments accept `str | Path | AttachmentInput`.

---

## 9. Public API (`client.py`)

```python
client = OutlookClient(credential=cred, mailbox="user@org.com")  # mailbox optional for delegated

# Read
client.get_email(message_id) -> Message
client.list_messages(folder="inbox", *, top=None) -> Iterator[Message]
client.search_email(*, sender=None, subject_contains=None,
                    since=None, until=None, unread=None,
                    has_attachments=None, folder=None) -> Iterator[Message]
                    # precise $filter only; free-text `query=` deferred to a future iteration
client.list_attachments(message_id) -> list[AttachmentMeta]
client.download_attachment(message_id, attachment_id, dest_path) -> Path

# Send
client.send_email(to, subject, body, *, cc=None, bcc=None,
                  attachments=None, html=False) -> None
client.reply(message_id, body, *, reply_all=False, attachments=None, html=False) -> None

# Drafts
client.create_draft(to, subject, body, *, cc=None, bcc=None,
                    attachments=None, html=False) -> Message
client.update_draft(message_id, *, subject=None, body=None, to=None, ...) -> Message
client.send_draft(message_id) -> None
client.discard_draft(message_id) -> None     # delete an unsent draft

# Folders
client.list_folders() -> list[Folder]
client.create_folder(name, *, parent=None) -> Folder

# Organize
client.delete_email(message_id, *, permanent=False) -> None
client.move_email(message_id, dest_folder) -> Message
```

All folder arguments accept well-known name / display name / id. All errors
raise `GraphError`.

---

## 10. Attachments (`attachments.py`)

- **Threshold:** 3 MB (Graph's single-request limit).
- **Send / draft / reply:**
  - No attachment, or all small → single `sendMail` (or inline on the draft).
  - Any attachment > threshold → ensure a **draft** exists, attach small ones
    inline, stream large ones via `POST .../attachments/createUploadSession`
    then chunked `PUT` to the returned `uploadUrl`, then send the draft.
- **Download:** `GET /messages/{id}/attachments/{aid}/$value` streamed to
  `dest_path`.

This is why `send_email` is built on the same draft machinery the public draft
methods expose.

---

## 11. Search mapping

Initial implementation is **precise structured filtering only**, mapped to
`$filter` (combined with `and`), with `$orderby receivedDateTime desc`. This
covers the motivating case — e.g. "subject contains ABC, received after
$date, with an attachment":

- `sender` → `from/emailAddress/address eq '...'`
- `subject_contains` → `contains(subject, '...')`
- `since` / `until` → `receivedDateTime ge/le ...`
- `unread` → `isRead eq false`
- `has_attachments` → `hasAttachments eq true`
- `folder` → scopes the request path via `FolderResolver`

**Deferred (future iteration):** free-text `$search`. Because Graph forbids
combining `$search` with `$filter`/`$orderby` in several cases, keeping the
initial surface to `$filter` avoids that complexity entirely. A `query=`
parameter can be added later as a separate code path.

---

## 12. CLI (`cli.py`)

Thin `click` wrapper over `OutlookClient` for manual use and auth bootstrap.
Reads config from CLI flags / a config file path. Subcommands mirror the API:
`login`, `get`, `list`, `search`, `send`, `reply`, `draft`, `download`,
`folders`, `mkdir`, `move`, `delete`. Output as readable text or `--json`.

---

## 13. Error model (`exceptions.py`)

```python
class GraphError(Exception):
    status_code: int
    code: str | None        # Graph error code, e.g. "ErrorItemNotFound"
    message: str
    request_id: str | None
```

Raised for all non-2xx Graph responses after retries are exhausted, and for auth
failures surfaced by MSAL.

---

## 14. Testing strategy

TDD with `pytest` + `respx`. No live tenant required; CI-friendly.

`respx` is chosen over `pytest-vcr`/`vcrpy` deliberately: VCR-style cassettes
must be recorded against a real tenant first (incompatible with the no-tenant
decision), capture tokens and message PII that would need scrubbing, and make
our trickiest cases — 429/503 retry, the 3 MB upload-session boundary, malformed
errors, multi-page pagination — hard to provoke. Those are all easier to author
by hand with `respx`, which is also the native httpx transport mock.

- Record representative Graph JSON payloads into `tests/fixtures/`.
- Mock the token endpoint and Graph endpoints with `respx`.
- Dedicated tests for the tricky paths: `GraphError` mapping, 429/503 retry with
  `Retry-After`, lazy pagination across `@odata.nextLink`, folder-name
  resolution + cache, upload-session chunking boundary (just under / just over
  3 MB), and the small-vs-large send branch.
- Each module tested in isolation behind its interface.

An optional, opt-in live integration suite is explicitly **not** part of this
plan (can be added later).

---

## 15. Implementation phases

Each phase is test-first and independently reviewable.

**Phase 0 — Scaffolding**
`pyproject.toml` deps, package skeleton, `exceptions.GraphError`, config
dataclasses with validation, `pytest`/`respx` harness and fixture loading.

**Phase 1 — Auth**
`Credential` protocol; `ClientSecretCredential` (simplest, mock token endpoint),
then `DeviceCodeCredential`; encrypted token cache via `msal-extensions` with
fallback. Tests mock MSAL/token responses.

**Phase 2 — GraphSession**
httpx wrapper, token injection, error→`GraphError` mapping, retry, lazy
`paginate`, `stream_get`. Heavily unit-tested with `respx`.

**Phase 3 — Models**
Pydantic models + parsing from recorded Graph JSON; round-trip/aliasing tests.

**Phase 4 — Read operations**
`get_email`, `list_messages`, `search_email`, `list_attachments` (search mapping
+ pagination integrated).

**Phase 5 — Folders**
`FolderResolver` (well-known / display-name / id + cache), `list_folders`,
`create_folder`, `move_email`.

**Phase 6 — Send, drafts, attachments**
`attachments` module (inline + upload sessions), `create_draft` / `update_draft`
/ `send_draft` / `discard_draft`, `send_email`, `reply`, `download_attachment`.

**Phase 7 — Delete**
`delete_email` with soft default and `permanent=True`.

**Phase 8 — CLI**
`click` subcommands over the finished client.

**Phase 9 — Docs**
Update `README.md` with install + usage examples; finalize `AGENTS.md`.

---

## 16. Out of scope

Read/unread toggling, importance/flag mutation, categories, batch requests,
calendar, and contacts. Free-text `$search` (precise `$filter` ships first).
Async API. Env-var config loading. Live integration tests. Each can be a
follow-up if needed.
