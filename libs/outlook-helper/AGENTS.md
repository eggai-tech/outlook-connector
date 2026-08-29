# AGENTS.md

Guidance for AI agents and developers working in this repo.

## What this is

`outlook-helper` is a synchronous Python library (plus a thin `click` CLI) for
M365 email via the Microsoft Graph API. Design lives in
[specs/DESIGN.md](specs/DESIGN.md).

## Layout

```
outlook_helper/
  config.py       # DelegatedConfig / AppOnlyConfig / CertificateConfig (dataclasses)
  auth.py         # Credential protocol; DeviceCodeCredential, ClientSecretCredential
  http.py         # GraphSession: the single HTTP seam (auth, retry, pagination, streaming)
  folders.py      # FolderResolver: well-known name / display name / id -> Graph id
  models.py       # Pydantic models (Message, Folder, AttachmentMeta, ...)
  attachments.py  # attachment loading, inline payloads, upload-session chunking
  client.py       # OutlookClient: the public verbs
  cli.py          # click CLI
tests/            # pytest + respx; fixtures in tests/fixtures/
```

## Conventions

- **TDD.** Write a failing test first, watch it fail, then implement. All tests
  mock Graph with `respx`; never call a live tenant in tests.
- **One HTTP seam.** All Graph traffic goes through `GraphSession`. Don't create
  `httpx` calls elsewhere.
- **No I/O in constructors.** Credentials build their MSAL app lazily on first
  `get_token()`; keep it that way so objects construct offline.
- **Models are output-only.** Graph's camelCase + `emailAddress` envelope is
  mapped to snake_case flattened models via aliases/validators.
- **Errors.** Non-2xx responses become `GraphError` (status, code, message,
  request id). Don't leak raw `httpx` errors.

## Commands

```bash
uv sync                 # install (Python 3.13)
uv run pytest           # run tests
uv run pytest -q tests/test_client_send.py   # one module
uv run outlook-helper --help
```

## Scope notes

In scope: the verbs in the README. Out of scope (deliberately): read/unread,
flags, categories, calendar, contacts, batch requests, free-text `$search`
(precise `$filter` ships first), and async. See DESIGN.md §16.
