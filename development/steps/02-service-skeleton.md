# Piece 2 — Service skeleton (config + lifecycle)

**Ships:** a deployable long-running service with no email logic.
**Depends on:** nothing hard (can run parallel to Piece 1).

## Why

Get the operational shell right — config, fail-fast startup, clean shutdown,
logging — before adding polling. After this piece the process boots, validates,
connects to the bus, idles, and shuts down cleanly.

## Scope

### Configuration

`pydantic-settings` (`BaseSettings`) with layering per
[the spec](../docs/implementation.md#configuration):

- **Structural config — YAML file:** mailbox address list, poll interval, bus
  connection (transport/broker URL, channel/topic), Azure `tenant_id`,
  `client_id`.
- **Secrets — environment variables only:** `client_secret`. **Never** written
  to the config file.

### Startup — fail fast

Per [the spec](../docs/implementation.md#startup--shutdown): validate the
Pydantic config (including that `client_secret` is present) and connect to the
bus eagerly. On any failure, log a clear error and **exit non-zero**; the
orchestrator restarts with backoff. No bespoke bus-reconnection logic.

### Shutdown — graceful drain

On `SIGTERM`/`SIGINT`: stop starting new cycles, let any in-flight publish
finish (cancel the sleep, not the publish), close the bus connection, exit 0.

### Observability

Stdlib `logging`, configurable level, to stdout (for container log capture).
No metrics stack.

## Done when

- `main.py` runs as a real long-running service: loads + validates config,
  connects to the bus, idles in an `asyncio` loop, and drains cleanly on
  `SIGTERM`/`SIGINT`.
- Missing/invalid config (e.g. absent `client_secret`) exits non-zero with a
  clear message.
