# Design / Architecture document for this solution

This is a live document that records major code decisions and directions. It does not describe what can be read in code.
Context and explanations belong in each module and function's docstring.

If this document and the code approach disagree, this document is right. 
This document needs to be updated with a decision _before_ starting to make code changes.

## Foundations

- Python long-running service.
- Uses `asyncio`.
- The helper is **synchronous**; every call into it runs via
  `asyncio.to_thread(...)` so a blocking request or `Retry-After` sleep never
  stalls the event loop.
- Uses `uv` for venv and dependency management.

## Configuration

Use Pydantic Settings, read from a config.yaml file.
Secrets are never allowed in the configuration file, and must come from env variables.

## Bus

This service communicates with other services using a message bus, through the EggAI SDK.
A supported bus type and its parameters are given in the config file.
All interaction happens through the SDK, so that the bus implmentation may be changed. 
No interaction goes directly to the underlying message queue. 

## Data model

All input and output objects are Pydantic models.

The Outlook Helper is considered to be part of this project.
Data models from Outlook Helper can be used as-is in interaction with other services. 
There is no need to map them to new models.


## Outlook Graph API interaction

- Uses `outlook-helper` for all interaction with the M365 Graph API. It is
  vendored at [libs/outlook-helper](../libs/outlook-helper) as a uv workspace
  member; it was previously a private git dependency pinned by tag. The
  connector needs the extensions listed under
  [verification](#dependency-verification) (`since_exclusive`,
  `include_headers`, `html_body`, attachment-metadata `$select`) — all present
  in the vendored copy, none of them in the original `v0.1.0`.

## Storage system

The current solution uses PostgreSQL for storage. 
To keep it possible to add other storage options later, the code does not refer to PostgreSQL by name but instead 
calls a generic `save_to_storage()` wrapper function.

## Ingestion

- Emails are guaranteed to be saved and published **at-least-once**.
- Emails are saved to storage first, then a message is published to the bus.
- Consuming emails should be idempotent and accept re-publishing an email multiple times.
- Emails are processed and saved individually. If one fails, the batch continues.

## Observability

Stdlib `logging`, configurable level, to stdout (for container log capture).

## Per-mailbox cursor

- The mailbox has a last-seen `received_date_time` cursor.
- The cursor lives in-memory only. On startup it is initialised to
  **"now"**, so only mail received strictly after process start is bridged.
- The cursor can be initialized to another date at startup through a configuration field.
- A poll queries the mailbox for messages with `receivedDateTime > cursor` (**strict `>`**).
- After fetching, the cursor advances to the **maximum `receivedDateTime`
  observed** in the batch (server-relative, avoids host/server clock skew).

## Identity

- Emails are identified using Graph API unique (immutable) ids.
