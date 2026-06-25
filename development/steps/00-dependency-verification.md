# Piece 0 — Dependency verification spike

**Status:** prerequisite — do this first.
**Ships:** a findings note, not a feature. De-risks every piece that follows.

## Why

The implementation spec [explicitly gates the design](../docs/implementation.md#dependency-verification)
on capabilities of `outlook-helper` v0.1.0 that must be confirmed against its
real API before building. Several design decisions collapse to fallbacks if a
capability is missing.

The dependency is checked out locally at `/Volumes/My/work/eggai/outlook-helper`
(and pinned in `pyproject.toml` at tag `v0.1.0`).

## What to confirm

- **Filter by `receivedDateTime`** — can we query a mailbox for messages with
  `receivedDateTime > cursor` (strict `>`)?
- **Body as HTML globally** — can the body be requested as HTML
  (lossless superset) for all messages?
- **Identity & threading fields:**
  - `internetMessageId` (RFC 822 message id) — required for `message_id`.
  - `conversationId` (native Graph property) — required for coarse threading.
  - `internetMessageHeaders` (`In-Reply-To` / `References`) — these are **not
    returned by default** and must be explicitly requested. Confirm the helper
    can surface them.
  - **Fallback:** if the SMTP threading headers are not cheaply available but
    `conversationId` is, the design falls back to coarse thread grouping by
    conversation without precise reply-chain linkage.
- **Attachments** — read `hasAttachments` (free, native boolean) and fetch
  per-attachment metadata (`filename`, `content_type`, `size`). Confirm the cost
  (expected: one extra Graph call per attachment-bearing email).
- **Retry / backoff** — does the helper perform its own retry/backoff and
  surface `Retry-After`? This determines whether the connector needs any retry
  beyond the natural fixed-delay interval gap.

## Done when

A short findings note exists that, for each item above, records "confirmed" or
"missing → fallback X", and the implementation spec is amended if any assumption
turns out to be wrong.
