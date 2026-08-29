# outlook-helper

A Python library for working with M365 email through the Microsoft Graph API.

Functionality:
- authenticate users (delegated device-code flow, or app-only client credentials)
- get email
- search email (precise filters: subject, date range, sender, attachments)
- send email (with attachments) — new email or reply
- create and send drafts
- download attachments
- create folders
- delete email (soft or permanent)
- move email between folders

See [docs/reference.md](docs/reference.md) for the full library reference,
and [specs/DESIGN.md](specs/DESIGN.md) for the design.

## Install

Requires Python 3.13+.

The library lives inside the
[outlook-connector](https://github.com/eggai-tech/outlook-connector) repository,
under `libs/outlook-helper`. It has no dependency on the connector, and installs
on its own: git fetches the whole repository, but only this subdirectory is
built, so you get the library and nothing else.

With uv:

```toml
[project]
dependencies = ["outlook-helper>=0.4.0"]

[tool.uv.sources]
outlook-helper = { git = "ssh://git@github.com/eggai-tech/outlook-connector.git",
                   tag = "outlook-helper-v0.4.0", subdirectory = "libs/outlook-helper" }
```

With pip:

```bash
pip install "outlook-helper @ git+ssh://git@github.com/eggai-tech/outlook-connector.git@outlook-helper-v0.4.0#subdirectory=libs/outlook-helper"
```

Releases of this library are tagged `outlook-helper-v*`, separately from the
connector's own `v*` image tags. Pin a tag rather than a branch — the
`subdirectory` fragment is required either way.

## Library usage

```python
from outlook_helper import OutlookClient, DelegatedConfig, DeviceCodeCredential

# Delegated (signed-in user) — a device-code prompt is printed on first call.
credential = DeviceCodeCredential(
    DelegatedConfig(client_id="<app-client-id>", tenant_id="<tenant>")
)
client = OutlookClient(credential)  # defaults to the signed-in user's mailbox

# Read
msg = client.get_email("<message-id>")
for m in client.list_messages(folder="inbox", top=20):
    print(m.subject, m.from_.address)

# Precise search: subject contains "ABC", received after a date, with attachments
from datetime import datetime
results = client.search_email(
    subject_contains="ABC",
    since=datetime(2026, 1, 1),
    has_attachments=True,
)

# Send (large attachments are streamed via an upload session automatically)
client.send_email(
    to="someone@example.com",
    subject="Hello",
    body="<p>Hi there</p>",
    html=True,
    attachments=["/path/to/file.pdf"],
)

# Drafts
draft = client.create_draft("someone@example.com", "Draft", "body")
client.update_draft(draft.id, subject="Updated subject")
client.send_draft(draft.id)

# Reply, move, delete
client.reply("<message-id>", "Thanks!", reply_all=True)
client.move_email("<message-id>", "Archive")        # well-known name, display name, or id
client.delete_email("<message-id>", permanent=False)

# Attachments
for att in client.list_attachments("<message-id>"):
    client.download_attachment("<message-id>", att.id, f"./{att.name}")
```

### App-only (daemon) authentication

```python
from outlook_helper import OutlookClient, AppOnlyConfig, ClientSecretCredential

credential = ClientSecretCredential(
    AppOnlyConfig(client_id="...", tenant_id="...", client_secret="...")
)
# App-only has no signed-in user, so a mailbox is required.
client = OutlookClient(credential, mailbox="shared@example.com")
```

## CLI

A thin CLI ships as `outlook-helper`:

```bash
outlook-helper --client-id <id> login
outlook-helper --client-id <id> list --folder inbox --top 20
outlook-helper --client-id <id> search --subject-contains ABC --has-attachments
outlook-helper --client-id <id> send --to a@x.com --subject Hi --body "Hello" --attach ./f.pdf
outlook-helper --client-id <id> move <message-id> Archive
outlook-helper --client-id <id> delete <message-id> --permanent
```

Config can also come from environment variables: `OUTLOOK_CLIENT_ID`,
`OUTLOOK_TENANT_ID`, `OUTLOOK_MAILBOX`, `OUTLOOK_CLIENT_SECRET`,
`OUTLOOK_CACHE_PATH`. Run `outlook-helper --help` for all commands.

## Development

Standalone, from a clone of this directory:

```bash
uv sync
uv run pytest
```

From inside the outlook-connector workspace, the suite needs its own rootdir, so
use the recipe that supplies it:

```bash
just test-helper
```

Tests mock the Graph HTTP layer with `respx`; no live tenant is required.
