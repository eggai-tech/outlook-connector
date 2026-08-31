# Piece 1 — Bus contract (owned models)

**Ships:** the owned Pydantic bus contract + a script that publishes one event.
**Depends on:** Piece 0 (threading-field availability shapes the email model).

## Why

Locking the bus contract early lets downstream consumers start building against
it while the connector internals are still in progress. The contract is
**owned** — `outlook-helper`'s own model is mapped into ours at the boundary, so
the bus is decoupled from that dependency's schema.

This piece **replaces** the prototype `schemas.py`, whose shape does not match
the spec (it carries attachment *content* as base64, a different envelope, and
`body_text`/`body_html` splits).

## Scope

### Envelope

A Pydantic model per [the spec](../../docs/DESIGN.md#envelope):

| Field            | Meaning                                                |
| ---------------- | ------------------------------------------------------ |
| `type`           | `email.received`                                       |
| `source_mailbox` | Address of the mailbox that received the mail          |
| `fetched_at`     | When the connector observed the message (poll-run ts)  |
| `email`          | The email model below                                  |

`source_mailbox` is first-class (provenance/routing). `fetched_at` means "the
connector saw this at T", not "the cursor was T".

### Email model

Owned Pydantic model per [the spec](../../docs/DESIGN.md#email-model):

- **Identity:** `message_id` (RFC 822 `internetMessageId`), `graph_id` (Graph
  internal item id, kept for follow-up Graph calls).
- **Core headers:** `from` / `to` / `cc` as `{name, address}` submodels,
  `subject`, `received_datetime`, `extra_headers: dict[str, str]`.
- **Threading:** `conversation_id`, `in_reply_to`, `references: list`.
- **Body:** `body` (as delivered, no lossy conversion), `body_content_type`
  (`"html"` | `"text"`), `preview` (optional `bodyPreview`).
- **Attachments (metadata only):** `has_attachments: bool`,
  `attachments: list[{filename, content_type, size}]` (empty until Piece 4).
  No attachment *content*.

## Done when

- Models validate and round-trip through the bus serializer.
- Unit tests cover required vs optional fields and the `from`/`to`/`cc`
  submodels.
- A throwaway script publishes one hand-built `email.received` event to the bus
  (replacing the prototype `create_email`).
