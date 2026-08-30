"""Save every received email to storage: the generic API and the backends.

Implements `the spec <development/specs/storage/20260830-save-to-storage.md>`.

The connector saves an email through one call, :func:`save`. It passes the
email to each active backend in turn and returns what each one returned, keyed
by backend name. Each backend defines its own return value (an id, a URL, …)
and decides how and where it stores an email — one blob per message, a row per
field, whatever suits it. The connector knows none of those details; it only
hands over the :class:`~outlook_connector.schemas.Email` object.

A backend reports failure by raising. :func:`save` stops at the **first**
failure, so the remaining backends are not called, and re-raises as
:class:`StorageError`, which names the backend and the email. Retrying is not
this module's job, so the same email may later be saved again by a backend that
already holds it: **duplicates are the backend's business**, and the API makes
no promise about what a second save does.

Backends are constructed once, by :func:`init_backends` at startup, and live as
long as the process. A backend that cannot start raises there, so the connector
does not start either — backends are not optional. Listing none is still a
valid configuration: :func:`save` then stores nothing and returns an empty dict.
"""

import logging
from abc import ABC, abstractmethod
from collections.abc import Sequence
from typing import Any

from outlook_connector.schemas import Email

logger = logging.getLogger(__name__)


class EmailStorage(ABC):
    """The interface a storage backend must implement.

    The constructor takes **no arguments**: a backend that needs settings reads
    its own ``STORAGE_<BACKEND>_<FIELD>`` values from
    :mod:`outlook_connector.config` itself, and one that needs a connection
    opens it in the constructor — so a backend that cannot start fails at
    startup rather than on the first email.
    """

    @abstractmethod
    async def save(self, email: Email) -> Any:
        """Store ``email``; return whatever identifies it in this backend."""


class StorageError(Exception):
    """A backend failed to save an email.

    Carries the backend name and the email id so the message says *which*
    backend failed on *which* email. The backend's own exception is kept as
    ``__cause__``.
    """

    def __init__(self, backend: str, email_id: str, cause: Exception):
        self.backend = backend
        self.email_id = email_id
        super().__init__(
            f"storage backend {backend!r} failed to save email {email_id}: {cause}"
        )


# Imported last: `memory` imports EmailStorage from this module, so the class
# above must already be bound. Every further backend gets its own module here
# and one line in BACKENDS below — no entry points, no dynamic imports.
from outlook_connector.storage.memory import MemoryStorage

#: Every backend the connector answers to, by the name used in
#: ``STORAGE_BACKENDS``. :func:`~outlook_connector.config.load_settings`
#: validates configured names against this, so an unknown name fails at
#: configuration load.
BACKENDS: dict[str, type[EmailStorage]] = {"memory": MemoryStorage}

# The active backends as (name, instance) pairs, in configured order. Mutated
# in place by init_backends so `from ... import` of this module's functions
# keeps seeing the current list.
_active: list[tuple[str, EmailStorage]] = []


def init_backends(names: Sequence[str]) -> list[tuple[str, EmailStorage]]:
    """Build the active backends, in ``names`` order, for the life of the process.

    Raises on an unknown name, or out of a backend's constructor, so the
    connector never starts with storage half-configured. ``init_backends([])``
    deactivates everything, which is also how a test returns to a clean slate.
    """
    active: list[tuple[str, EmailStorage]] = []
    for name in names:
        backend_cls = BACKENDS.get(name)
        if backend_cls is None:
            raise ValueError(
                f"Unknown storage backend: {name!r}. "
                f"Known: {', '.join(sorted(BACKENDS)) or '(none)'}"
            )
        active.append((name, backend_cls()))
    _active[:] = active
    return active


def active_backends() -> list[tuple[str, EmailStorage]]:
    """The active ``(name, instance)`` pairs, in the order :func:`save` calls them."""
    return list(_active)


async def save(email: Email) -> dict[str, Any]:
    """Save ``email`` to every active backend; return their values by name.

    Backends are called one after another — the order is not part of the API.
    The first one to raise stops the walk and surfaces as :class:`StorageError`;
    the backends after it are not called, and the ones before it keep what they
    already stored.
    """
    results: dict[str, Any] = {}
    for name, backend in _active:
        try:
            results[name] = await backend.save(email)
        except Exception as exc:
            raise StorageError(name, email.message_id, exc) from exc
    return results
