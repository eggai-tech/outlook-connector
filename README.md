# outlook-connector

A bridge between Microsoft 365 email and the [EggAI](https://github.com/eggai-tech) bus.

The EggAI bus is the message bus that EggAI agents communicate over. This
connector gives those agents an **email channel**: mail arriving in a connected
M365 mailbox is published onto the bus as `email.received` events, attachment
content included, so agents can consume email without knowing anything about
Microsoft 365 or the Graph API.

> **Status:** beta. The inbound path works end to end; bus message shapes and
> configuration may still change. Outbound (sending mail from bus events) is on
> the roadmap but not part of the current codebase.

## What it does

A long-running service that polls one configured M365 mailbox and, for each new
message:

- maps it to an owned, Graph-independent `Email` model — sender/recipients,
  subject, HTML/text body, `internetMessageId`, received timestamp;
- fetches **attachment content** inline (file attachments; item/reference
  attachments carry no bytes and are skipped);
- wraps it in a typed CloudEvents envelope (`type: email.received`) and
  publishes it to a configurable channel over kafka, redis, or an in-memory
  transport.

Delivery bias is **never duplicate, occasionally drop**: the poll cursor only
advances past a message once it is published, and the first failure stops the
batch until the next cycle. Mail arriving while the service is down is not
replayed (seed `initial_cursor` to backfill from a known instant).

## The bus contract

Each event is an [eggai](https://pypi.org/project/eggai/) `BaseMessage`
(CloudEvents 1.0) with `source: /outlook-connector`, `type: email.received`,
and a `data` payload of:

| field            | meaning                                             |
| ---------------- | --------------------------------------------------- |
| `source_mailbox` | address of the mailbox that received the mail       |
| `fetched_at`     | when the connector observed it (poll-run timestamp) |
| `email`          | the `Email` model below                             |

`Email`: `id` (Graph immutable id), `internet_message_id` (RFC 822 — the
natural dedup key for consumers), `from_addresses`, `to_addresses`, `subject`,
`received_at`, `body_html`/`body_text`, `has_attachments`,
`attachments[{file_name, content_type, body}]`, `mime_content`. Models live in
[`src/outlook_connector/schemas.py`](src/outlook_connector/schemas.py) and
[`bus.py`](src/outlook_connector/bus.py).

The channel name is set by `bus.channel` in `config.yaml`; the eggai SDK
prefixes it with the `EGGAI_NAMESPACE` environment variable (default `eggai`),
so events land on `<namespace>.<channel>` — set the namespace to match your
consumers.

## Quickstart

Docker Compose builds the connector and starts it alongside postgres and redis.
[`just`](https://just.systems) drives it:

```sh
just up          # seed config, build the image, start everything
just logs        # follow the connector
just down        # stop (keeps the volumes)
just reset       # wipe volumes + image and start clean
just             # list every recipe
```

`just up` copies `config.yaml.example` -> `config.yaml` and `env.example` ->
`.env` on first run. Fill both in before the connector can reach M365:
`mailbox` in `config.yaml`, and `AZURE_TENANT_ID` / `AZURE_CLIENT_ID` /
`AZURE_CLIENT_SECRET` in `.env`. Every Azure connection parameter comes from
the environment — put any of them in `config.yaml` and startup fails with a
pointed error. Neither file is baked into the image: `config.yaml` is
bind-mounted read-only and the credentials arrive through the environment.

One prerequisite: **host ports 35432 (postgres) and 36379 (redis)**, overridable
via `POSTGRES_HOST_PORT` / `REDIS_HOST_PORT` in `.env`. Inside the compose
network the services are always `postgres:5432` and `redis:6379`. The build needs
no credentials of any kind.

Redis doubles as the EggAI bus for the compose stack — `BUS__TRANSPORT` and
`BUS__BROKER_URL` default to it, overriding `config.yaml`. Set them in `.env` to
point at a kafka broker instead. Postgres is started as a supporting service;
the connector has no persistence layer yet and does not read from it.

## The outlook-helper library

The connector's Microsoft 365 access is a self-contained library,
[`outlook-helper`](libs/outlook-helper), that happens to live in this
repository. It knows nothing about the connector, the bus or EggAI — it is a
plain wrapper over the Graph API for reading, searching, sending and replying to
mail, handling attachments and managing folders.

**Other projects can use it on its own.** See
[libs/outlook-helper/README.md](libs/outlook-helper/README.md) for what it does,
how to install it as a dependency, and its API.

## Development

Repository layout, the uv workspace, the test suites and release tagging are all
covered in [DEVELOPMENT.md](DEVELOPMENT.md).

## Built on

- [EggAI SDK](https://pypi.org/project/eggai/) for bus connectivity.
- [`outlook-helper`](libs/outlook-helper) for Microsoft 365 Graph API access.

## Implementation

Architecture, the bus message contract, configuration, and all implementation
decisions live in [development/DESIGN.md](development/DESIGN.md).

## License

[MIT](LICENSE).
