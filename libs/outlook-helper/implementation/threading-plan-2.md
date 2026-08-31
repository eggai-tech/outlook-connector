# Thread-safe reply into a shared mailbox (without notifying the original sender)

## Context

An app-only agent (its own app credentials) has access to shared mailbox
`sandbox@example.com`. It reads an email from Alice → the shared mailbox, and needs
to **add more information into that same conversation** so that everyone in the org
with access to the shared mailbox can find it in-thread — *without* emailing Alice back.

Why the naive approach fails: Outlook/Exchange thread messages by the server-assigned
`conversationId` plus the RFC-2822 `In-Reply-To`/`References` headers — **not** by subject.
A fresh `send_email(...)` (even same subject) gets a new `conversationId` and no
`References` header, so it shows up as a *separate* conversation — the "floats separately,
hard to find" problem. `conversationId` is read-only and the threading headers can't be
set on an arbitrary send, so **the only robust way to thread is Graph's `createReply`
action** on the original message.

The library already exposes `reply()` ([client.py:258-280](../outlook_helper/client.py#L258-L280)),
which uses `createReply`/`createReplyAll` correctly — but it always addresses the reply to
the original sender (Alice) and only lets the caller override body + attachments. There is
**no way to reply while redirecting recipients**, which is exactly what "thread-only, don't
notify Alice" requires.

Decision (from discussion): the threaded message will be **addressed to the shared mailbox
itself** (`sandbox@example.com`) so a copy is delivered into the shared Inbox and is visible
even in flat list view. Alice is not a recipient.

## Recommended approach

Fill the one gap by exposing reply-draft creation, then compose with the **existing**
draft verbs (`update_draft`, `send_draft`) — no duplicated send/patch logic.

### Change: add `create_reply_draft` to `OutlookClient`

In [outlook_helper/client.py](../outlook_helper/client.py), alongside the draft family
(`create_draft`/`update_draft`/`send_draft`/`discard_draft`), add:

```python
def create_reply_draft(
    self, message_id: str, *, reply_all: bool = False
) -> OutlookMessage:
    """Create a reply draft that inherits the original's conversationId
    (and In-Reply-To/References headers), so it threads when sent.
    The draft is pre-addressed to the original sender; override recipients
    with update_draft() before sending to redirect or suppress that."""
    action = "createReplyAll" if reply_all else "createReply"
    return OutlookMessage.model_validate(
        self._session.request(
            "POST", f"{self._base_path}/messages/{message_id}/{action}"
        ).json()
    )
```

This reuses the exact `createReply` call already in `reply()`; the only new surface is
*returning the draft* instead of immediately patching + sending it. Recipient override,
body, and attachments are then all handled by the existing `update_draft()`
([client.py:225-250](../outlook_helper/client.py#L225-L250)) and `send_draft()`
([client.py:252-253](../outlook_helper/client.py#L252-L253)) — `update_draft(to=...)` PATCHes
`toRecipients`, which *replaces* Alice with the shared mailbox.

Optionally refactor `reply()` to call `create_reply_draft()` internally to remove the
duplicated `createReply` block (low-risk, keeps behavior identical).

### Resulting agent workflow

```python
client = OutlookClient(credential, mailbox="sandbox@example.com")  # app-only

draft = client.create_reply_draft(original_message_id)   # inherits conversationId
client.update_draft(
    draft.id,
    to="sandbox@example.com",          # into the shared mailbox, NOT Alice
    body="<p>Extra information …</p>",
    html=True,
)
client.send_draft(draft.id)
```

Both the Sent copy and the delivered Inbox copy carry the original's `conversationId`, so
they thread with Alice's email for everyone viewing the shared mailbox.

## Loop-safety note (important, given delivery to the shared mailbox)

Because the message is delivered *into* `sandbox@example.com`'s Inbox, any agent loop that
scans the Inbox for mail to process will re-encounter its **own** message. The agent's
processing logic must **skip messages whose `from` address is the shared mailbox itself**
(self-authored). `OutlookMessage` already exposes `from_` and `conversationId`
([client.py:45-46](../outlook_helper/client.py#L45-L46)), so this is a caller-side filter, not
a library change. (Alternative guards: move the delivered copy to a subfolder, or tag with a
custom `x-*` header — not needed for the MVP.)

## Permissions note

App-only `createReply` + send against a shared mailbox requires the app to hold
**`Mail.ReadWrite` + `Mail.Send`** application permissions, ideally scoped to that mailbox
(Application Access Policy). No code depends on this, but the reply/send will 403 without it.

## Verification

- **Unit test** (respx, mirroring the existing `reply()` test): mock
  `POST /users/sandbox@example.com/messages/{id}/createReply` → draft JSON carrying a
  `conversationId`; assert `create_reply_draft()` returns an `OutlookMessage` with that
  `conversationId`. Then mock the `PATCH .../messages/{draftId}` and
  `POST .../messages/{draftId}/send` and assert the composed flow sends `toRecipients =
  [sandbox@example.com]` and does **not** include Alice. Run with `uv run pytest`.
- **Live smoke** (optional, real tenant): reply into a real conversation via the workflow
  above; confirm in OWA that (a) the new message sits inside Alice's conversation, (b) Alice
  is not a recipient, and (c) a second shared-mailbox user sees it in-thread.

## Out of scope

- Preserving the quoted original body in the reply (current `reply()` already replaces the
  body wholesale — matched here for consistency).
- Any "exclude sender" search helper for the loop guard (caller-side filter suffices).
