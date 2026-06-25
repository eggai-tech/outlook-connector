# Piece 3 — Inbound polling core

**Ships:** the MVP — real inbound mail flows to the bus end to end.
**Depends on:** Piece 1 (models), Piece 2 (config + lifecycle), Piece 0 (mapping).

## Why

This is the heart of the spec. Build single-mailbox first to nail the cursor and
mapping logic, then generalise to multi-mailbox with isolation. You could ship
after this piece and have a working connector.

## Scope

### Poll loop

Single `asyncio` loop per [the spec](../docs/implementation.md#flow):
poll configured mailboxes **sequentially** each cycle, then **sleep the
configured interval** (fixed-delay — the loop never overlaps itself). Default
~60s, configurable.

### Per-mailbox cursor

Per [the spec](../docs/implementation.md#per-mailbox-cursor):

- Each mailbox has its own **last-seen `receivedDateTime` cursor**.
- Cursor is **volatile** (in-memory), initialised to **"now"** on startup — only
  mail received strictly after process start is bridged.
- Poll queries `receivedDateTime > cursor` (**strict `>`**).
- Cursor advances to the **maximum `receivedDateTime` observed** in the batch
  (server-relative, avoids clock skew).

### Per-cycle, per-mailbox processing

Per [the spec](../docs/implementation.md#per-cycle-per-mailbox-processing),
inside `try/except` so one mailbox's failure never crashes the loop or blocks
others:

1. Fetch new messages (`receivedDateTime > cursor`).
2. Sort the batch **ascending** by `receivedDateTime`.
3. For each message in order: map to the email model → publish `email.received`
   → on success, advance cursor to this message's `receivedDateTime`.
4. On the **first publish failure, stop the batch**, leaving the cursor at the
   last successfully-published message. Next cycle resumes from there.
5. On any **Graph API error** (`429`, `5xx`, network): log it, leave the cursor
   untouched, continue to the next mailbox. No custom backoff (honor
   `Retry-After` only if `outlook-helper` surfaces it — see Piece 0).

### Mapping

Map `outlook-helper`'s message model → the owned email model at the boundary.
Request body as HTML and the `internetMessageHeaders` needed for threading
(per Piece 0 findings).

### Multi-mailbox + observability

Poll N mailboxes sequentially with a per-mailbox cursor map and per-mailbox
`try/except` isolation. Emit a per-cycle, per-mailbox summary log: fetched
count, published count, errors.

## Design principle

**Never duplicate, occasionally drop.** Every cursor/failure decision biases
toward never publishing a duplicate `email.received`, at the accepted cost of
occasionally dropping an email. See
[known limitations](../docs/implementation.md#known-limitations).

## Done when

- A real email arriving in a connected mailbox after startup is published to the
  bus exactly once.
- Publish failure mid-batch leaves the cursor correct (no duplicates, resumes
  next cycle).
- A Graph error on one mailbox doesn't affect the others or crash the loop.
- Each cycle logs a per-mailbox summary.
