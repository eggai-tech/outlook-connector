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
- Uses `outlook-helper` for all interaction with the M365 Graph API. It is
  vendored at [libs/outlook-helper](../libs/outlook-helper) as a uv workspace
  member; it was previously a private git dependency pinned by tag. The
  connector needs the extensions listed under
  [verification](#dependency-verification) (`since_exclusive`,
  `include_headers`, `html_body`, attachment-metadata `$select`) — all present
  in the vendored copy, none of them in the original `v0.1.0`.
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
   - If the message's free `has_attachments` flag is set, make **one extra Graph
     call** (`list_attachments`, async-wrapped) to fetch per-attachment metadata;
     messages without attachments make no extra call. See the
     [email model](#email-model)'s attachments note.
   - Map it to the bus [email model](#email-model).
   - **Save it to storage** (see [Storage](#storage)) — storage is the durable
     record, so it is written *before* the bus. A failed save means the message
     never reaches the bus: log it, advance the cursor anyway, and continue with
     the next message. Retrying a failed save is out of scope, and holding the
     cursor back would re-deliver every *later* message next cycle, so this
     drops the one email rather than duplicating the rest.
   - Publish an `email.received` event to the bus.
   - On success, advance the cursor to this message's `receivedDateTime`.
   - On the **first publish failure, stop the batch** and leave the cursor at
     the last successfully-published message. The next cycle resumes from there.
     The attachment-metadata fetch runs inside this same per-message guard, so a
     Graph error fetching metadata likewise stops the batch with the cursor at
     the last success — never a duplicate, retried next cycle.
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
- `mime` — the whole message as RFC 822 MIME. Graph serves it only on an
  extra per-message call, which the poller does not make, so it is empty
  unless it was populated. [Storage](#storage) keeps it verbatim.

**Attachments** (metadata only — content is out of scope)

- `has_attachments` — native Graph boolean (free).
- `attachments` — list of `{filename, content_type, size}`. Populating this
  costs an **extra Graph call per attachment-bearing email** (attachments are a
  separate navigation property), so it is gated behind the free `has_attachments`
  flag. No attachment *content* is bridged.

  The extra call lives in the **poller**, not the boundary mapping: the mapping
  ([`map_email`](../src/outlook_connector/mapping.py)) stays a pure, synchronous, side-effect-free
  function that receives the already-fetched metadata, while the poller owns all
  Graph I/O (`list_attachments` via `asyncio.to_thread`). Mapping is defensive
  about the helper's `Optional` attachment fields — a missing `name` /
  `content_type` / `size` maps to `""` / `""` / `0` rather than raising.

## Configuration

`pydantic-settings` (`BaseSettings`) with layering:

- **Structural config — YAML file:** mailbox address list, poll interval, bus
  connection (transport/broker URL, channel/topic).
- **Azure connection parameters — environment variables only:**
  `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`. The secret and
  the identifiers that travel with it are one credential, so they come from one
  place. **Never** written to the config file: a config file carrying any of
  them (or a pre-migration `azure:` block) is rejected at startup rather than
  silently ignored.
- **Storage — `DATABASE_URL`, environment variable only:** the PostgreSQL
  connection URL. Like the Azure credential, it is a secret and never written to
  the config file. It is required; there is nothing to select and nothing to turn
  off.

## Storage

Every received email is saved to **PostgreSQL** through a single call,
`save_to_storage(email)`, which takes the connector's own `Email` object. There
is one store, chosen up front; the connector does not select, layer, or abstract
over alternatives. `Email` carries a `mime` field, empty unless it was populated,
which is stored verbatim.

Failure is reported by raising `StorageError`. The connector never retries a
failed save, but an email *saved* and then *not published* is re-fetched and
saved again next cycle (see [Flow](#per-cycle-per-mailbox-processing)), so the
same email can arrive twice: **duplicate handling is the store's business**, and
the call makes no promise about what a second save of the same email does.

The database connection is established **once, at startup**, from `DATABASE_URL`
and lives as long as the process; if it cannot be established the connector does
not start (see [Startup](#startup--shutdown)).

Logging never carries the subject, the body, or the `mime` field — only the email
id.

## Startup & shutdown

**Startup — fail fast.** Validate the Pydantic config (including that the
`AZURE_*` env vars and `DATABASE_URL` are present), connect to PostgreSQL, and
connect to the bus eagerly. On any failure, log a clear error and **exit
non-zero**; the orchestrator restarts with backoff. Mid-run bus loss is handled
by the eggai client's reconnect plus the "publish fails → stop batch, retry next
cycle" path — no bespoke reconnection logic in the MVP.

**Shutdown — graceful drain.** On `SIGTERM`/`SIGINT`, stop starting new cycles,
let the current in-flight publish finish (cancel the sleep, not the publish),
close the bus connection, and exit 0. Because the cursor advances per successful
publish, interrupting *between* publishes is always safe.

## Observability

Stdlib `logging`, configurable level, to stdout (for container log capture).
Each cycle emits a per-mailbox summary: fetched count, published count, dropped
count (emails whose save failed), errors.
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
- **Failed save drops the email.** An email the database would not take is not
  published and is not retried; the cursor moves past it. See
  [Storage](#storage).

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
