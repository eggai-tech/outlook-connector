# Piece 4 — Attachment metadata enrichment

**Ships:** attachment-bearing emails carry attachment metadata on the bus.
**Depends on:** Piece 3 (the polling core + mapping).

## Why

Additive and cleanly isolated — it lands last without touching the core loop.
Attachment metadata costs an **extra Graph call per attachment-bearing email**
(attachments are a separate navigation property), so it is gated behind the free
`has_attachments` boolean.

## Scope

Per [the spec](../docs/implementation.md#email-model):

- For each fetched message where `has_attachments` is true, make the extra Graph
  call (via `outlook-helper`) to fetch per-attachment metadata.
- Populate `attachments: list[{filename, content_type, size}]` on the email
  model.
- Messages without attachments make **no** extra call.
- Attachment **content** stays out of scope — metadata only.

## Done when

- An email with attachments is published with a populated `attachments` list
  (`filename`, `content_type`, `size` per attachment).
- An email without attachments triggers no extra Graph call and carries an empty
  `attachments` list.
