"""Tests for the save-to-storage API, its registry, and the reference backend.

These cover the guarantees the spec makes to a caller: every active backend is
called and its value returned under its name; the first backend to raise stops
the walk and surfaces as a :class:`StorageError` naming it; and listing no
backend at all is a valid configuration that stores nothing.

The extra backends live in this module and are registered in ``BACKENDS`` for
the duration of a test by :func:`_active`, so the production registry keeps its
single real entry.

Runnable standalone (`python -m tests.test_storage`) or under pytest.
"""

import asyncio
from contextlib import contextmanager
from datetime import datetime, timezone

import pytest

from outlook_connector import storage
from outlook_connector.schemas import Email, EmailAddress
from outlook_connector.storage import EmailStorage, StorageError
from outlook_connector.storage.memory import MemoryStorage

_RECEIVED_AT = datetime(2026, 8, 30, 9, 0, 0, tzinfo=timezone.utc)


def _email(message_id: str = "<m1@example.com>") -> Email:
    return Email(
        message_id=message_id,
        graph_id="graph-1",
        **{"from": EmailAddress(address="alice@example.com")},
        to=[EmailAddress(address="bob@egg-ai.com")],
        subject="March invoice",
        received_datetime=_RECEIVED_AT,
        body="<p>Please find attached.</p>",
        body_content_type="html",
    )


class _Echo(EmailStorage):
    """A second backend, to prove the values come back keyed by name."""

    async def save(self, email: Email) -> str:
        return f"echo:{email.message_id}"


class _Broken(EmailStorage):
    """Reports failure the only way a backend can: by raising."""

    def __init__(self):
        self.calls = 0

    async def save(self, email: Email):
        self.calls += 1
        raise RuntimeError("disk on fire")


class _Untouched(EmailStorage):
    """Records whether it was called at all."""

    def __init__(self):
        self.calls = 0

    async def save(self, email: Email) -> str:
        self.calls += 1
        return "ok"


@contextmanager
def _active(*backends: tuple[str, type[EmailStorage]]):
    """Register ``backends``, activate them in order, then restore the registry.

    Yields the active ``(name, instance)`` pairs so a test can assert on the
    instances the connector actually holds.
    """
    saved = dict(storage.BACKENDS)
    storage.BACKENDS.update(dict(backends))
    try:
        yield storage.init_backends([name for name, _ in backends])
    finally:
        storage.BACKENDS.clear()
        storage.BACKENDS.update(saved)
        storage.init_backends([])


# --- no backends -----------------------------------------------------------


def test_no_active_backends_stores_nothing_and_returns_empty():
    async def go():
        with _active() as pairs:
            assert pairs == []
            assert await storage.save(_email()) == {}

    asyncio.run(go())


# --- one and many backends -------------------------------------------------


def test_one_backend_stores_the_email_and_returns_its_value():
    async def go():
        with _active(("memory", MemoryStorage)) as pairs:
            email = _email()
            assert await storage.save(email) == {"memory": "<m1@example.com>"}

            (_, backend) = pairs[0]
            assert backend.emails == {"<m1@example.com>": email}

    asyncio.run(go())


def test_two_backends_both_run_and_both_values_come_back():
    async def go():
        with _active(("memory", MemoryStorage), ("echo", _Echo)) as pairs:
            results = await storage.save(_email())

            assert results == {
                "memory": "<m1@example.com>",
                "echo": "echo:<m1@example.com>",
            }
            (_, memory) = pairs[0]
            assert "<m1@example.com>" in memory.emails

    asyncio.run(go())


# --- failures --------------------------------------------------------------


def test_a_raising_backend_surfaces_as_storage_error_naming_it():
    async def go():
        with _active(("broken", _Broken)):
            email = _email()
            with pytest.raises(StorageError) as caught:
                await storage.save(email)

            error = caught.value
            assert error.backend == "broken"
            assert error.email_id == email.message_id
            # The message says which backend failed on which email...
            assert "broken" in str(error)
            assert email.message_id in str(error)
            # ...and the backend's own exception is kept as the cause.
            assert isinstance(error.__cause__, RuntimeError)
            assert str(error.__cause__) == "disk on fire"

    asyncio.run(go())


def test_a_failing_backend_stops_the_walk_before_the_next_one():
    async def go():
        with _active(("broken", _Broken), ("later", _Untouched)) as pairs:
            with pytest.raises(StorageError):
                await storage.save(_email())

            (_, broken), (_, later) = pairs
            assert broken.calls == 1
            assert later.calls == 0  # never reached

    asyncio.run(go())


# --- the registry and startup ----------------------------------------------


def test_unknown_backend_name_fails_at_startup():
    with pytest.raises(ValueError, match="Unknown storage backend: 'nope'"):
        storage.init_backends(["nope"])
    assert storage.active_backends() == []


def test_a_backend_that_cannot_start_stops_the_connector():
    class _WontStart(EmailStorage):
        def __init__(self):
            raise ConnectionError("no route to storage")

        async def save(self, email: Email):  # pragma: no cover - never reached
            raise AssertionError

    saved = dict(storage.BACKENDS)
    storage.BACKENDS["wont-start"] = _WontStart
    try:
        with pytest.raises(ConnectionError):
            storage.init_backends(["wont-start"])
    finally:
        storage.BACKENDS.clear()
        storage.BACKENDS.update(saved)
        storage.init_backends([])


def test_backends_are_built_once_and_kept_in_configured_order():
    with _active(("echo", _Echo), ("memory", MemoryStorage)) as pairs:
        assert [name for name, _ in pairs] == ["echo", "memory"]
        # active_backends() hands back the same instances, call after call.
        assert [b for _, b in storage.active_backends()] == [b for _, b in pairs]


def test_memory_is_the_registered_reference_backend():
    assert storage.BACKENDS["memory"] is MemoryStorage


# --- the reference backend -------------------------------------------------


def test_memory_backend_clear_returns_it_to_a_known_state():
    async def go():
        backend = MemoryStorage()
        await backend.save(_email())
        assert backend.emails

        backend.clear()
        assert backend.emails == {}

    asyncio.run(go())


if __name__ == "__main__":
    test_no_active_backends_stores_nothing_and_returns_empty()
    test_one_backend_stores_the_email_and_returns_its_value()
    test_two_backends_both_run_and_both_values_come_back()
    test_a_raising_backend_surfaces_as_storage_error_naming_it()
    test_a_failing_backend_stops_the_walk_before_the_next_one()
    test_unknown_backend_name_fails_at_startup()
    test_a_backend_that_cannot_start_stops_the_connector()
    test_backends_are_built_once_and_kept_in_configured_order()
    test_memory_is_the_registered_reference_backend()
    test_memory_backend_clear_returns_it_to_a_known_state()
    print("All storage tests passed.")
