import datetime

from conftest import T0, FakeClient, make_attachment, make_message

from outlook_connector.poller import Poller


def _at(seconds: int) -> datetime.datetime:
    return T0 + datetime.timedelta(seconds=seconds)


def test_rescan_publishes_unseen_oldest_first():
    client = FakeClient(
        messages=[
            make_message("newer", received_at=_at(20)),
            make_message("older", received_at=_at(10)),
        ]
    )
    poller = Poller(client=client)

    messages = poller.poll_mailbox()

    assert [m.id for m in messages] == ["older", "newer"]
    listing = client.search_calls[0]
    assert listing["ids_only"] is True
    assert listing["oldest_first"] is True
    assert listing["since"] is None
    # full content is fetched per message, after the cheap listing
    assert client.get_email_calls == ["older", "newer"]


def test_seen_set_suppresses_republish_within_process():
    message = make_message("m1", received_at=_at(10))
    client = FakeClient(messages=[message])
    poller = Poller(client=client)

    assert [m.id for m in poller.poll_mailbox()] == ["m1"]
    poller.mark_published(message)

    # the mail is still in the folder; the seen-set keeps the cycle quiet
    assert poller.poll_mailbox() == []
    assert client.get_email_calls == ["m1"]  # no second full fetch either


def test_restart_republishes_whatever_is_still_in_the_folder():
    """At-least-once: the seen-set is process-local by design; a fresh poller
    re-emits everything still present and consumers dedupe."""
    message = make_message("m1", received_at=_at(10))
    client = FakeClient(messages=[message])

    first = Poller(client=client)
    first.mark_published(message)
    assert first.poll_mailbox() == []

    restarted = Poller(client=client)
    assert [m.id for m in restarted.poll_mailbox()] == ["m1"]


def test_pruning_forgets_mail_that_left_the_folder():
    """The listing is the truth: moved-out mail leaves the set (bounding it by
    folder size), and mail moved back in is re-published — consumers dedupe."""
    message = make_message("m1", received_at=_at(10))
    client = FakeClient(messages=[message])
    poller = Poller(client=client)
    poller.mark_published(message)

    client.messages = []  # the mover filed it away
    assert poller.poll_mailbox() == []
    assert poller._published_ids == set()

    client.messages = [message]  # a human dragged it back
    assert [m.id for m in poller.poll_mailbox()] == ["m1"]


def test_batch_bound_takes_oldest_unseen():
    cohort = [make_message(f"m{i}", received_at=_at(10 + i)) for i in range(3)]
    client = FakeClient(messages=cohort)
    poller = Poller(client=client, batch_max_messages=2)
    poller.mark_published(cohort[0])  # m0 already published

    messages = poller.poll_mailbox()

    # bound applies to the *unseen* mail, oldest first
    assert [m.id for m in messages] == ["m1", "m2"]
    assert client.get_email_calls == ["m1", "m2"]  # no full fetch for seen mail


def test_ignore_received_before_is_passed_as_since():
    client = FakeClient()
    poller = Poller(client=client, ignore_received_before=_at(0))

    poller.poll_mailbox()

    assert client.search_calls[0]["since"] == _at(0)


def test_source_folder_is_passed_through():
    client = FakeClient()
    poller = Poller(client=client, source_folder="Bankbestätigungen")

    poller.poll_mailbox()

    assert client.search_calls[0]["folder"] == "Bankbestätigungen"


def test_heartbeat_fires_per_listed_and_fetched_message():
    beats = []
    client = FakeClient(messages=[make_message(f"m{i}") for i in range(2)])
    poller = Poller(client=client, heartbeat=lambda: beats.append(1))

    poller.poll_mailbox()

    assert len(beats) == 4  # 2 listed + 2 fetched


def test_fetch_attachments_gated_on_flag():
    client = FakeClient(attachments=[make_attachment()])
    poller = Poller(client=client)

    assert poller.fetch_attachments(make_message("plain")) == []
    assert client.attachment_calls == []

    attachments = poller.fetch_attachments(make_message("rich", has_attachments=True))
    assert len(attachments) == 1
    assert client.attachment_calls == ["rich"]


def test_oversized_attachment_stripped_to_metadata():
    client = FakeClient(
        attachments=[
            make_attachment("small.pdf", content=b"ok"),
            make_attachment("huge.pdf", content=b"x" * 100),
        ]
    )
    poller = Poller(client=client, max_attachment_bytes=10)

    small, huge = poller.fetch_attachments(make_message("m1", has_attachments=True))

    assert small.content == b"ok"
    assert huge.content is None
    assert huge.size == 100  # original size survives for the metadata-only entry

def test_listed_message_gone_before_fetch_is_skipped():
    """The downstream mover files mail out of the folder concurrently; a 404
    between the listing and the full fetch must skip that message, not abort
    the whole cycle."""
    from outlook_helper import GraphError

    kept = make_message("kept")
    gone = make_message("gone")

    class RacyClient(FakeClient):
        def get_email(self, message_id, **kwargs):
            if message_id == "gone":
                raise GraphError(status_code=404, message="ErrorItemNotFound")
            return super().get_email(message_id, **kwargs)

    client = RacyClient(messages=[gone, kept])
    poller = Poller(client=client)

    fetched = poller.poll_mailbox()

    assert [m.id for m in fetched] == ["kept"]
