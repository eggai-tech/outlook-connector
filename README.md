# outlook-connector

A bridge between Microsoft 365 email and the [EggAI](https://github.com/eggai-tech) bus.

The EggAI bus is the message bus that EggAI agents communicate over. This
connector gives those agents an **email channel**: inbound mail is published onto
the bus as events, and outbound bus messages are delivered as email. Agents can
receive and reply to email without knowing anything about Microsoft 365 or the
Graph API.

> **Status:** usable (beta). The connector works end to end; bus message shapes
> and configuration may still change.

## What it does

It is a **bidirectional bridge** between one or more M365 mailboxes and the bus:

- **Inbound** — email arriving in a connected mailbox is published onto the bus
  for agents to consume.
- **Outbound** — messages an agent puts on the bus are sent as email from the
  appropriate mailbox.

Supported on both directions:

- **Read and send** plain email (subject, body, recipients).
- **Attachments**, inbound and outbound.
- **Threading** — replies are linked to the original conversation so agents can
  hold a back-and-forth rather than fire one-off messages.

## Mailboxes

The connector serves a **configured set of M365 mailboxes**, identified by
address. Each bus message is tied to a mailbox address, so an inbound event
records which mailbox received the mail and an outbound message states which
mailbox to send from. The set of mailboxes is fixed at deploy time.

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
`mailboxes` in `config.yaml`, and `AZURE_TENANT_ID` / `AZURE_CLIENT_ID` /
`AZURE_CLIENT_SECRET` in `.env`. Every Azure connection parameter comes from
the environment — put any of them in `config.yaml` and startup fails with a
pointed error. Neither file is baked into the image: `config.yaml` is
bind-mounted read-only and the credentials arrive through the environment.

One prerequisite: **host ports 15432 (postgres) and 16379 (redis)**, overridable
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
decisions live in [docs/implementation.md](docs/implementation.md).
