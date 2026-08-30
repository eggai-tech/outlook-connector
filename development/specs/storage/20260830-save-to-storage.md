---
type: feature
title: Saves emails to storage
---
# Save Emails to Storage

## Overview

The outlook-connector saves every received email to storage.

The feature has two parts:

- A generic internal API that the connector calls once per received email.
- One or more storage backends that implement this API.

The configuration lists the active backends. If several backends are active,
the connector saves the email to each of them, one after the other. The order
is not part of the API.

## Scope

In scope:

- The generic save-to-storage API.
- The interface that a backend must implement.
- The configuration that selects the active backends.
- A reference in-memory backend.

Out of scope:

- All other backend implementations.
- Retrying an email after a failed save.

## The Email Object

The connector passes its own `Email` object, which holds a full email message.

This object is separate from the `outlook-helper` schemas on purpose. The
helper wraps the Graph API, while the connector is free to build richer
abstractions on top of it.

The object includes a `mime` field, which may be empty if it was not
populated.

## Behaviour

The connector saves an email through one call. The call passes the email to
each active backend and returns the value each backend returned, keyed by
backend name. Each backend defines its own return value, for example an id or
a URL. The connector only logs these values.

A backend reports a failure by raising. The call stops at the first failure,
so the remaining backends are not called, and the error reaches the caller.
The error names the backend that failed and the email that was being saved.

The connector calls save before it puts the message on the bus. Storage is the
durable record, so it is written first. If the save fails, the message does
not reach the bus. The caller logs the error and continues with the next
email.

A backend is created once, when the connector starts, and lives as long as the
process. A backend that cannot start raises, and the connector does not start
either. Backends are not optional.

If no backend is active, nothing is stored. This is a valid configuration.

## Configuration

`STORAGE_BACKENDS` holds the list of active backend names, for example
`["memory"]`. A name can appear only once, so there is at most one instance of
each backend.

Each backend has its own flat, prefixed configuration fields, named
`STORAGE_<BACKEND>_<FIELD>`, for example `STORAGE_S3_BUCKET`. A backend reads
its own fields. The connector ignores the fields of a backend that is not
active.

If `STORAGE_BACKENDS` holds a name that no backend answers to, the connector
fails at configuration load.

## Logging

The connector logs the email id, the backend names, and the values the
backends returned. It never logs the subject, the body, or the `mime` field.

## Technical Considerations

Each backend decides how and where it stores an email. For example, a database
backend may split the email into fields and store them in its own schema,
while an S3 backend may store one blob under the email id. A backend may use
the `mime` field, or ignore it.

The connector does not know these details. It only passes the email object to
each active backend.

Duplicates are handled by the backend. The API makes no promise about what a
second save of the same email does. Because a failed save stops the remaining
backends, the same email may be saved again later by a backend that already
holds it.

## Reference Backend

An in-memory backend is part of this feature. It keeps the emails in memory
and can be emptied, so that tests start from a known state. It serves as the
reference implementation and as the backend used in tests. It is not meant for
production use.
