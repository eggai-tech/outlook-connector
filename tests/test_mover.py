import asyncio
import itertools

import pytest
from conftest import T0, make_message
from eggai import Channel, InMemoryTransport
from outlook_helper import GraphError

from outlook_connector.bus import build_bus_event
from outlook_connector.mover import (
    EMAIL_MOVE,
    MoveEmail,
    MoveEmailMessage,
    Mover,
    build_mover_agent,
)

MAILBOX = "inbox@example.com"
_CHANNEL_SEQ = itertools.count()


class FakeMoveClient:
    """Stands in for outlook_helper.OutlookClient, move side only."""

    def __init__(self, fails_on=()):
        self.fails_on = set(fails_on)
        self.move_calls: list[tuple[str, str]] = []

    def move_email(self, message_id, dest_folder):
        if dest_folder in self.fails_on:
            raise GraphError(404, "folder not found", code="ErrorItemNotFound")
        self.move_calls.append((message_id, dest_folder))
        # Graph reissues the id in the destination folder.
        return make_message(f"{message_id}-moved")


def make_mover(**client_kwargs) -> Mover:
    return Mover(mailbox=MAILBOX, client=FakeMoveClient(**client_kwargs))


def test_apply_moves_the_message():
    mover = make_mover()

    mover.apply(MoveEmail(message_id="m1", destination_folder="archive"))

    assert mover.client.move_calls == [("m1", "archive")]


def test_apply_accepts_a_matching_mailbox():
    mover = make_mover()

    mover.apply(
        MoveEmail(message_id="m1", destination_folder="Processed", mailbox=MAILBOX)
    )

    assert mover.client.move_calls == [("m1", "Processed")]


def test_apply_ignores_another_mailbox():
    mover = make_mover()

    mover.apply(
        MoveEmail(message_id="m1", destination_folder="archive", mailbox="other@example.com")
    )

    assert mover.client.move_calls == []


def test_apply_swallows_graph_errors():
    mover = make_mover(fails_on={"nope"})

    mover.apply(MoveEmail(message_id="m1", destination_folder="nope"))  # must not raise

    assert mover.client.move_calls == []


def test_command_round_trips_through_json():
    command = MoveEmailMessage(
        source="/consumer", data=MoveEmail(message_id="m1", destination_folder="archive")
    )

    parsed = MoveEmailMessage.model_validate_json(command.model_dump_json())

    assert parsed.type == EMAIL_MOVE
    assert parsed.data.message_id == "m1"
    assert parsed.data.destination_folder == "archive"


@pytest.mark.parametrize(
    "destination_folder, message_id", [("", "m1"), ("archive", "")]
)
def test_command_rejects_empty_fields(destination_folder, message_id):
    with pytest.raises(ValueError):
        MoveEmail(message_id=message_id, destination_folder=destination_folder)


def run_on_the_bus(*events) -> Mover:
    """Publish ``events`` on one channel and let a live mover agent consume them."""
    mover = make_mover()
    # InMemoryTransport keeps its queues in class-level state, so they outlive
    # the event loop each scenario runs in. A channel of its own per scenario
    # keeps one test from publishing into the dead loop of the last.
    name = f"emails-{next(_CHANNEL_SEQ)}"

    async def scenario() -> None:
        transport = InMemoryTransport()
        agent = build_mover_agent(mover, channel=name, transport=transport)
        await agent.start()

        channel = Channel(name, transport=transport)
        for event in events:
            await channel.publish(event)
        await asyncio.sleep(0.05)  # let the consume loop drain the queue
        await agent.stop()

    asyncio.run(scenario())
    return mover


def test_agent_applies_a_command_off_the_bus():
    """End to end over a real (in-memory) transport: publish, consume, move."""
    mover = run_on_the_bus(
        MoveEmailMessage(
            source="/consumer",
            data=MoveEmail(message_id="m1", destination_folder="archive"),
        )
    )

    assert mover.client.move_calls == [("m1", "archive")]


def test_agent_ignores_the_events_sharing_its_channel():
    """Commands and outbound events ride one channel: only the former is handled."""
    received = build_bus_event(make_message("m1"), source_mailbox=MAILBOX, fetched_at=T0)
    command = MoveEmailMessage(
        source="/consumer", data=MoveEmail(message_id="m2", destination_folder="archive")
    )

    mover = run_on_the_bus(received, command)

    assert mover.client.move_calls == [("m2", "archive")]
