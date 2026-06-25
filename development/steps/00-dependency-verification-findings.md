# Piece 0 — Dependency verification findings

**Dependency:** `outlook-helper`, tag `v0.1.0` (`3be412b`; local checkout HEAD ==
`v0.1.0`). Verified by reading the source, not by live Graph calls.

**Source of truth for the claims below:**
- `outlook_helper/client.py` — `search_email`, `list_messages`, `list_attachments`
- `outlook_helper/schemas.py` — `OutlookMessage`, `OutlookAttachmentMeta` (note
  `GraphModel.model_config = extra="ignore"` — unmodeled Graph fields are dropped)
- `outlook_helper/http.py` — `GraphSession` (retry/backoff, pagination, headers/params)

## Summary

| # | Capability | Verdict |
|---|------------|---------|
| 1 | Filter by `receivedDateTime` **strict `>`** | ⚠️ **Missing** → only `ge`/`le`; emulate `>` in connector |
| 2 | Body as HTML globally | ⚠️ **Partial** → not requested; relies on Graph default, but type is reported (lossless) |
| 3 | `internetMessageId` for `message_id` | ❌ **Missing** → not modeled, not requested |
| 4 | `conversationId` | ✅ **Confirmed** |
| 5 | `internetMessageHeaders` (`In-Reply-To`/`References`) | ❌ **Missing** → fall back to coarse threading by `conversationId` |
| 6 | `hasAttachments` + per-attachment metadata | ✅ **Confirmed** (with a bandwidth-cost caveat) |
| 7 | Retry/backoff + `Retry-After` | ✅ **Confirmed** (429/503 only) |

Two findings outside the original checklist are material and appear at the end:
**(A) the helper is fully synchronous** and **(B) there is a clean escape hatch**
(`GraphSession`) that resolves most of the gaps below cheaply.

---

## 1. Filter by `receivedDateTime` (strict `>`) — ⚠️ Missing strict operator

`search_email(since=...)` ([client.py:368-390](../../../outlook-helper/outlook_helper/client.py#L368))
builds `receivedDateTime ge <dt>` — that is `>=`, **not** the strict `>` the
design requires. There is no `gt` option. `until` maps to `le`.

Also: `_fmt_dt` formats to **whole-second** precision (`%Y-%m-%dT%H:%M:%SZ`),
discarding the sub-second precision Graph stores on `receivedDateTime`.

**Why it matters:** with a max-seen cursor and `since=cursor` (`ge`), the message
sitting *exactly at* the cursor is re-fetched every poll → a **duplicate
`email.received`**, which directly violates the "never duplicate" principle.

**Fallback (recommended):** keep calling `search_email(since=cursor)` and drop
any message whose `received_at <= cursor` in the connector, i.e. emulate strict
`>` client-side. Cheap, no escape hatch needed. Because of the second-precision
truncation, query with `since=cursor` truncated to the second (so nothing is
missed) and rely on the client-side `> cursor` (full precision) filter to remove
the re-fetched boundary rows.

## 2. Body as HTML globally — ⚠️ Partial (works, but not by request)

The helper sets **no** `Prefer: outlook.body-content-type="html"` header and no
`$select` (confirmed: `grep` for `Prefer`/`$select` in `outlook_helper/` →
none). It accepts whatever representation Graph returns by default — which *is*
HTML for `body`, so in practice HTML arrives.

Crucially, `OutlookMessage.body` is an `OutlookBody{content_type, content}`
([schemas.py:25-47](../../../outlook-helper/outlook_helper/schemas.py#L25)), so
the **actual** content type is reported. The connector can map `body` →
`body` and `body.content_type` → `body_content_type` losslessly regardless of
which representation Graph chose.

**Verdict:** the design's *intent* (lossless body + known type) is satisfied.
The literal "request HTML globally" is not done, but is unnecessary given the
type is surfaced. If a tenant ever defaulted to text, force HTML via the escape
hatch (custom `Prefer` header on a `GraphSession`, finding B).

## 3. `internetMessageId` → `message_id` — ❌ Missing

Not a field on `OutlookMessage`, and `extra="ignore"` means it is dropped even
if returned. Graph does not return it by default anyway — it requires
`$select=internetMessageId`. (`grep internetmessage outlook_helper/` → none.)

**Why it matters:** the design makes `message_id = internetMessageId` the
**primary identity** and threads by it. That assumption is broken by the
high-level API. The only stable identifier the helper surfaces is `graph_id`
(`OutlookMessage.id`).

**Fallbacks (decision needed — see "Proposed spec amendments"):**
- **(a) Escape hatch:** issue the list query through a `GraphSession` with
  `$select=...,internetMessageId,...`. Same call, no extra round-trip. Preserves
  the design as written.
- **(b) Amend design:** use `graph_id` as the bus `message_id` and rely on
  `conversation_id` for threading; drop the RFC-822 identity for the MVP.

## 4. `conversationId` — ✅ Confirmed

`OutlookMessage.conversation_id` (alias `conversationId`,
[schemas.py:52](../../../outlook-helper/outlook_helper/schemas.py#L52)). Native
Graph property, returned by default, always present. Maps directly to the bus
model's `conversation_id`.

## 5. `internetMessageHeaders` (`In-Reply-To` / `References`) — ❌ Missing

Not modeled, dropped by `extra="ignore"`, and not requested (needs
`$select=internetMessageHeaders`, which Graph omits by default).

**Fallback (already anticipated by the design):** coarse thread grouping by
`conversation_id` (finding 4) without precise reply-chain linkage. Confirmed
available. If precise linkage is wanted later, the escape hatch (finding B) can
`$select` the headers and parse them, since the helper's model would ignore them.

## 6. `hasAttachments` + per-attachment metadata — ✅ Confirmed (cost caveat)

- `OutlookMessage.has_attachments` — native bool, free, default `False`
  ([schemas.py:49](../../../outlook-helper/outlook_helper/schemas.py#L49)). ✅
- `client.list_attachments(message_id)` returns `OutlookAttachmentMeta` with
  `name`, `content_type`, `size` (plus `id`, `is_inline`)
  ([schemas.py:68-73](../../../outlook-helper/outlook_helper/schemas.py#L68)).
  The bus model wants `filename` — trivially `name → filename`. ✅
- **Cost:** one extra Graph call per attachment-bearing email (`list_attachments`
  → `GraphSession.paginate(.../attachments)`), matching the design's expectation.

**Caveat:** the helper adds **no `$select`** to the attachments query, so Graph
returns the **full** attachment resource — including `contentBytes` (base64) for
file attachments — which the model then *discards* via `extra="ignore"`. So
"metadata only" still transfers the entire attachment over the wire. Functionally
correct, but bandwidth-wasteful for large attachments. For true metadata-only,
use the escape hatch with `$select=id,name,contentType,size,isInline`. Acceptable
for the MVP; flag for later.

## 7. Retry / backoff / `Retry-After` — ✅ Confirmed

`GraphSession.request` ([http.py:54-81](../../../outlook-helper/outlook_helper/http.py#L54))
retries on **429 and 503** up to `max_retries=3`, sleeping for `Retry-After`
when present and falling back to exponential `2**attempt`
([http.py:131-138](../../../outlook-helper/outlook_helper/http.py#L131)).

**The connector needs no retry beyond the natural interval gap**, with two
caveats:
- Only **429/503** are retried. Other 5xx (500/502/504) and **network errors**
  (httpx exceptions) propagate immediately. The connector's per-mailbox
  `try/except` must therefore catch **both** `outlook_helper.GraphError` **and**
  `httpx`-level exceptions, then leave the cursor and move on (as the design
  already prescribes for "any Graph API error").
- All non-2xx responses (after retries) surface as a single `GraphError`
  carrying `status_code`, `code`, `message`, `request_id`
  ([exceptions.py](../../../outlook-helper/outlook_helper/exceptions.py)).

---

## Additional finding A — the helper is fully synchronous ⚠️ (not in checklist, important)

`GraphSession` uses `httpx.Client` (sync) and `time.sleep` for backoff
([http.py:33-39](../../../outlook-helper/outlook_helper/http.py#L33)).
`OutlookClient` exposes only sync methods. The connector runs an `asyncio` loop
(per the spec's Foundations). **Calling `OutlookClient` directly from the event
loop will block it** — and a throttled call's `Retry-After` sleep freezes the
entire loop for seconds.

**Required:** wrap every `OutlookClient` call in `asyncio.to_thread(...)` (or a
`run_in_executor`). This belongs in the implementation spec.

## Additional finding B — escape hatch via `GraphSession` ✅ (resolves 1/2/3/5/6)

`OutlookClient` owns a private `_session: GraphSession` and exposes public
`.base_path` and `.credential` properties ([client.py:63-69](../../../outlook-helper/outlook_helper/client.py#L63)).
`GraphSession.paginate(path, params)` accepts arbitrary OData params and
`request(..., headers=...)` accepts arbitrary headers. So the connector can
construct its own `GraphSession(client.credential)` and issue a tailored query —
e.g. `params={"$filter": "receivedDateTime gt <iso>", "$select": "id,internetMessageId,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,internetMessageHeaders", "$orderby": "receivedDateTime"}`
and `headers={"Prefer": 'outlook.body-content-type="html"'}` — getting strict
`>`, `internetMessageId`, the SMTP threading headers, and a forced-HTML body in a
**single** call, then map the raw JSON itself.

This is the highest-leverage option: it preserves the design as written for
findings 1, 2, 3, and 5 at no extra round-trips, at the cost of dropping down
from the typed high-level API to raw Graph JSON for the poll query.

---

## Proposed spec amendments (decision needed before building)

These touch design decisions, so I'm surfacing them rather than editing
`docs/implementation.md` unilaterally:

1. **Threading/identity (findings 3 & 5):** use the escape hatch (B) to `$select`
   `internetMessageId` + `internetMessageHeaders` and keep the design's
   message-id identity and precise threading — **vs** accept the helper's typed
   API as-is, set `message_id = graph_id`, and thread coarsely by
   `conversation_id` only.
2. **Poll query (finding 1):** confirm the client-side strict-`>` emulation over
   `search_email(since=…)` — **vs** also move the poll query to the escape hatch
   (B) to get native `gt`. (If we adopt B for #1 anyway, both are solved at once.)
3. **Async (finding A):** add "all `outlook-helper` calls run via
   `asyncio.to_thread`" to the spec — no real alternative; the library is sync.
4. **Error handling (finding 7):** spec should name both `GraphError` **and**
   `httpx` exceptions as the "Graph API error" the per-mailbox `try/except`
   catches.

**Done-when status:** every checklist item is recorded above as confirmed or
missing→fallback. The spec amendments are drafted but **await your call on
decisions 1 and 2** before I edit `docs/implementation.md`.
