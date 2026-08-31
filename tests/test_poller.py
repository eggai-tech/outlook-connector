import datetime

from conftest import T0, FakeClient, make_attachment, make_message

from outlook_connector.poller import Poller


def _at(seconds: int) -> datetime.datetime:
    return T0 + datetime.timedelta(seconds=seconds)


def test_default_cursor_is_now_value():
    poller = Poller(client=FakeClient(), now=lambda: T0)
    assert poller.cursor == T0


def test_poll_mailbox_sorts_oldest_first():
    client = FakeClient(
        messages=[
            make_message("newer", received_at=_at(20)),
            make_message("older", received_at=_at(10)),
        ]
    )
    poller = Poller(client=client, cursor=T0)

    messages = poller.poll_mailbox()

    assert [m.id for m in messages] == ["older", "newer"]
    assert client.search_calls[0]["since_exclusive"] == T0


def test_advance_is_monotonic():
    poller = Poller(client=FakeClient(), cursor=_at(10))

    poller.advance(_at(30))
    assert poller.cursor == _at(30)

    poller.advance(_at(20))  # out-of-order timestamp must not pull the cursor back
    assert poller.cursor == _at(30)

    poller.advance(None)
    assert poller.cursor == _at(30)


def test_fetch_attachments_gated_on_flag():
    client = FakeClient(attachments=[make_attachment()])
    poller = Poller(client=client, cursor=T0)

    assert poller.fetch_attachments(make_message("plain")) == []
    assert client.attachment_calls == []

    attachments = poller.fetch_attachments(make_message("rich", has_attachments=True))
    assert len(attachments) == 1
    assert client.attachment_calls == ["rich"]
