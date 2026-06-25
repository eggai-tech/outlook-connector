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

## Built on

- [EggAI SDK](https://pypi.org/project/eggai/) for bus connectivity.
- `eggai-tech/outlook-helper` (private) for Microsoft 365 Graph API access.

## Implementation

Architecture, the bus message contract, configuration, and all implementation
decisions live in [docs/implementation.md](docs/implementation.md).
