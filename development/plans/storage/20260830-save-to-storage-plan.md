---
type: plan
title: Saves emails to storage
spec: development/specs/storage/20260830-save-to-storage.md
---
# Save Emails to Storage — Plan

## Code Layout

    src/outlook_connector/storage/__init__.py   EmailStorage, StorageError,
                                                BACKENDS, save()
    src/outlook_connector/storage/memory.py     MemoryStorage

Each further backend gets its own module in the same package. The expected
backends are PostgreSQL and S3.

## The Interface

`EmailStorage` is an abstract base class with one abstract method:

    class EmailStorage(ABC):
        @abstractmethod
        async def save(self, email: Email) -> Any: ...

The constructor takes no arguments. A backend that needs settings reads them
from `config.py` itself. A backend that needs a connection opens it in its
constructor, so that a broken backend fails at startup.

## The Registry

`BACKENDS` is a plain dict in `storage/__init__.py`:

    BACKENDS: dict[str, type[EmailStorage]] = {"memory": MemoryStorage}

No entry points, no dynamic imports. A new backend adds one line here.

## Startup

At startup the connector reads `config.STORAGE_BACKENDS`, and for each name:

1. Looks the name up in `BACKENDS`. An unknown name raises, and the connector
   does not start.
2. Calls the class with no arguments and keeps the instance.

The instances are held in one module-level list of `(name, instance)` pairs,
in the configured order, and reused for the life of the process.

## The Save Call

    async def save(email: Email) -> dict[str, Any]

It walks the active backends in order and awaits `backend.save(email)`. It
collects the return values in a dict keyed by backend name and returns it.

If a backend raises, `save` stops and raises `StorageError` from the original
exception. `StorageError` carries the backend name and the email id, so that
the message says which backend failed on which email.

## The In-Memory Backend

`MemoryStorage` keeps a dict of email id to `Email`. Its `save` stores the
email and returns the email id. It has a `clear()` method for tests.

## Call Site

In the receive path, the connector awaits `save(email)` before it publishes
the message on the bus. It logs the returned dict. It catches `StorageError`,
logs it, and continues with the next email.

## Tests

- `save` with no active backends returns an empty dict and does nothing.
- `save` with one backend stores the email and returns `{"memory": <id>}`.
- `save` with two backends returns both values.
- A backend that raises makes `save` raise `StorageError`, with the backend
  name in the message and the original exception as its cause.
- A failing first backend leaves the second backend untouched.
- An unknown name in `STORAGE_BACKENDS` fails at startup.
- The receive path does not publish on the bus when `save` raises.
- The log line holds the email id, and holds no subject, body, or mime.

Tests use `MemoryStorage` and a small failing backend defined in the test
module.

## Order of Work

1. `EmailStorage`, `StorageError`, and `BACKENDS` in `storage/__init__.py`.
2. `MemoryStorage`, and its entry in `BACKENDS`.
3. `STORAGE_BACKENDS` in `config.py`, and the startup step that builds the
   instances.
4. `save()`.
5. The call site in the receive path.
