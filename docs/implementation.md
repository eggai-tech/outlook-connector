# Implementation

All major architecture and implementation details and decisions are referenced
here.

## Scope

This document describes the **inbound MVP**: a long-running service that polls
one or more M365 mailboxes and publishes received email onto the EggAI bus.

The [README](../README.md) describes a fuller, bidirectional bridge (outbound
send, attachment content both ways, reply threading). That is **aspirational and
ahead of the implementation** — it is *not* covered here. The MVP is inbound
only. Outbound, attachment *content*, and reply *sending* are explicitly out of
scope for this phase.

## Foundations

- Python long-running service.
- Uses `asyncio`.
- Uses `uv` for venv and dependency management.
- Uses the `eggai-tech/outlook-helper` private repository for all interaction
  with the M365 Graph API. **`v0.1.0` is insufficient** (see
  [verification](#dependency-verification)); the connector pins the tag that
  ships the extensions listed there (`since_exclusive`, `include_headers`,
  `html_body`, attachment-metadata `$select`).
- The helper is **synchronous**; every call into it runs via
  `asyncio.to_thread(...)` so a blocking request or `Retry-After` sleep never
  stalls the event loop.
- Email is exposed on the bus via an **owned** Pydantic model (see
  [Email model](#email-model)); outlook-helper's own model is mapped into it at
  the boundary, so the bus contract is decoupled from that dependency's schema.
- Bus messages are Pydantic models.

## Design principle: never duplicate, occasionally drop

Every cursor and failure-handling decision below biases toward **never
publishing a duplicate** `email.received` event, at the accepted cost of
**occasionally dropping** an email. Downstream agents may act on email (e.g.
reply), so a duplicate is worse than a gap for this MVP. The
[known limitations](#known-limitations) section lists exactly where drops can
occur.

## Flow

1. On startup, load and validate configuration (see [Configuration](#configuration)).
   Connect to the EggAI bus eagerly. **Fail fast** — see [Startup](#startup--shutdown).
2. Enter a single `asyncio` loop.
3. Each cycle, poll the configured mailboxes **sequentially**.
4. After the cycle completes, **sleep the configured interval** (fixed-delay
   scheduling — the loop never overlaps itself). Default interval ~60s,
   configurable.

### Per-mailbox cursor

- Each mailbox has its own **last-seen `receivedDateTime` cursor**.
- The cursor is **volatile** (in-memory only). On startup it is initialised to
  **"now"**, so only mail received strictly after process start is bridged.
- A poll queries the mailbox for messages with
  `receivedDateTime > cursor` (**strict `>`**).
- After fetching, the cursor advances to the **maximum `receivedDateTime`
  observed** in the batch (server-relative — avoids host/server clock skew).

### Per-cycle, per-mailbox processing

For each mailbox, within `try/except` so one mailbox's failure never crashes the
loop or blocks the others:

1. Fetch new messages (`receivedDateTime > cursor`).
2. Sort the batch **ascending** by `receivedDateTime`.
3. For each message, in order:
   - Map it to the bus [email model](#email-model).
   - Publish an `email.received` event to the bus.
   - On success, advance the cursor to this message's `receivedDateTime`.
   - On the **first publish failure, stop the batch** and leave the cursor at
     the last successfully-published message. The next cycle resumes from there.
4. On any **Graph API error**: log it, leave the cursor untouched, and continue
   to the next mailbox. The fixed-delay loop retries next cycle. The helper
   already retries `429`/`503` honoring `Retry-After`; everything else (other
   `5xx`, network failures) surfaces as an exception. The per-mailbox
   `try/except` therefore catches **both** `outlook_helper.GraphError` **and**
   `httpx` transport exceptions. No custom backoff in the MVP.

## EggAI bus

### Envelope

A bus message is a Pydantic model:

| Field            | Meaning                                                |
| ---------------- | ------------------------------------------------------ |
| `type`           | `email.received`                                       |
| `source_mailbox` | Address of the mailbox that received the mail          |
| `fetched_at`     | When the connector **observed** this message (poll run timestamp) |
| `email`          | The email model (below), extracted by consumers        |

`source_mailbox` is a first-class envelope field (provenance/routing depend on
it). `fetched_at` replaces the original vague "last run / read on date" — it
means "the connector saw this at T", not "the cursor was T".

### Email model

An owned Pydantic model. outlook-helper's model is mapped into it.

**Identity**

- `message_id` — the RFC 822 `internetMessageId` (e.g. `<abc@host>`), **not**
  Graph's internal item id. Threading links by `internetMessageId`.
- `graph_id` — Graph's internal item id, kept separately for follow-up Graph
  calls (e.g. fetching attachment metadata).

**Core headers**

- `from` — `{name, address}` submodel.
- `to` — list of `{name, address}`.
- `cc` — list of `{name, address}`.
- `subject`
- `received_datetime`
- `extra_headers` — `dict[str, str]` catch-all for non-modeled headers.

**Threading headers**

- `conversation_id` — Graph's `conversationId` (native message property, always
  present — cheap, coarse thread grouping).
- `in_reply_to` — SMTP `In-Reply-To` header.
- `references` — list, SMTP `References` header.

  `in_reply_to` and `references` (and `internetMessageId`) live in Graph's
  `internetMessageHeaders`, which is **not returned by default** and must be
  explicitly requested. See [verification](#dependency-verification).

**Body**

- `body` — the body string, as Graph delivers it (no lossy conversion).
- `body_content_type` — `"html"` or `"text"`. **HTML is requested globally**
  (lossless superset; consumers can strip to text).
- `preview` — optional short snippet (Graph's `bodyPreview`).

**Attachments** (metadata only — content is out of scope)

- `has_attachments` — native Graph boolean (free).
- `attachments` — list of `{filename, content_type, size}`. Populating this
  costs an **extra Graph call per attachment-bearing email** (attachments are a
  separate navigation property). No attachment *content* is bridged.

## Configuration

`pydantic-settings` (`BaseSettings`) with layering:

- **Structural config — YAML file:** mailbox address list, poll interval, bus
  connection (transport/broker URL, channel/topic), Azure `tenant_id`,
  `client_id`.
- **Secrets — environment variables only:** `client_secret`. **Never** written
  to the config file.

## Startup & shutdown

**Startup — fail fast.** Validate the Pydantic config (including that the
`client_secret` env var is present) and connect to the bus eagerly. On any
failure, log a clear error and **exit non-zero**; the orchestrator restarts with
backoff. Mid-run bus loss is handled by the eggai client's reconnect plus the
"publish fails → stop batch, retry next cycle" path — no bespoke reconnection
logic in the MVP.

**Shutdown — graceful drain.** On `SIGTERM`/`SIGINT`, stop starting new cycles,
let the current in-flight publish finish (cancel the sleep, not the publish),
close the bus connection, and exit 0. Because the cursor advances per successful
publish, interrupting *between* publishes is always safe.

## Observability

Stdlib `logging`, configurable level, to stdout (for container log capture).
Each cycle emits a per-mailbox summary: fetched count, published count, errors.
No metrics stack in the MVP.

## Known limitations

These are intentional consequences of the MVP design:

- **Downtime gap.** The cursor is volatile and starts at "now", so email
  arriving while the service is down is **silently dropped** — never published.
  Pre-existing inbox mail at startup is ignored. There is no replay/backfill;
  recovery is manual.
- **Boundary drop.** With max-seen cursor + strict `>`, a second email sharing
  the exact `receivedDateTime` second of the batch maximum, arriving after that
  poll, can be dropped.
- **Attachment content.** Only attachment *metadata* is bridged; content is not.

## Dependency verification

Done. The full results are in
[the findings note](../development/steps/00-dependency-verification-findings.md);
the required helper changes are specified in
[the helper-changes note](../development/steps/00-outlook-helper-changes.md).
Outcome against `v0.1.0`:

- ✅ **Confirmed:** `conversationId`; `hasAttachments` + per-attachment metadata
  (`name`→`filename`, `content_type`, `size`) at one extra Graph call per
  attachment-bearing email; retry/backoff on `429`/`503` honoring `Retry-After`.
- ❌ **Was missing → fixed in the helper:** `internetMessageId` and
  `internetMessageHeaders` (`In-Reply-To`/`References`) are not modeled or
  `$select`ed by `v0.1.0`. Rather than fall back to coarse-only threading, the
  helper is extended to surface both (via `include_headers`), so the design's
  `message_id = internetMessageId` identity and precise reply-chain threading
  are preserved.
- ⚠️ **Was partial → fixed in the helper:** no strict `>` filter (only
  `ge`/`le`, truncated to whole seconds) and HTML body only by Graph default.
  The helper adds `since_exclusive` (native `gt`, sub-second precise) and
  `html_body` (`Prefer: outlook.body-content-type="html"`).

The **coarse-threading-by-`conversationId` fallback** remains the documented
contingency if the helper changes are ever reverted, but the MVP targets the
extended helper.
