from click.testing import CliRunner

from outlook_helper.auth import ClientSecretCredential, DeviceCodeCredential
from outlook_helper.cli import build_client, main
from outlook_helper.schemas import OutlookFolder, OutlookMessage


class FakeClient:
    def __init__(self):
        self.calls = []

    def get_email(self, message_id):
        self.calls.append(("get_email", message_id))
        return OutlookMessage(id=message_id, subject="Hello")

    def list_messages(self, folder="inbox", *, top=None):
        self.calls.append(("list_messages", folder, top))
        return iter(
            [
                OutlookMessage(id="1", subject="One"),
                OutlookMessage(id="2", subject="Two"),
            ]
        )

    def search_email(self, **kwargs):
        self.calls.append(("search_email", kwargs))
        return iter([OutlookMessage(id="9", subject="Found")])

    def send_email(self, to, subject, body, **kwargs):
        self.calls.append(("send_email", to, subject, body, kwargs))

    def delete_email(self, message_id, *, permanent=False):
        self.calls.append(("delete_email", message_id, permanent))

    def move_email(self, message_id, dest_folder):
        self.calls.append(("move_email", message_id, dest_folder))
        return OutlookMessage(id=message_id, subject="moved")

    def create_folder(self, name, *, parent=None):
        self.calls.append(("create_folder", name, parent))
        return OutlookFolder(id="F1", displayName=name)

    def list_folders(self):
        self.calls.append(("list_folders",))
        return [OutlookFolder(id="F1", displayName="Inbox")]

    def create_draft(self, to, subject, body, **kwargs):
        self.calls.append(("create_draft", to, subject, body, kwargs))
        return OutlookMessage(id="D1", subject=subject)

    def download_attachment(self, message_id, attachment_id, dest_path):
        self.calls.append(("download_attachment", message_id, attachment_id, dest_path))
        return dest_path


def run(args, fake):
    return CliRunner().invoke(main, args, obj={"client": fake})


# --- client construction ---


def test_build_client_delegated():
    client = build_client(
        {
            "auth": "delegated",
            "client_id": "cid",
            "tenant_id": "tid",
            "mailbox": None,
            "client_secret": None,
            "cache_path": None,
        }
    )
    assert isinstance(client.credential, DeviceCodeCredential)


def test_build_client_app_only():
    client = build_client(
        {
            "auth": "app-only",
            "client_id": "cid",
            "tenant_id": "tid",
            "mailbox": "u@x.com",
            "client_secret": "shh",
            "cache_path": None,
        }
    )
    assert isinstance(client.credential, ClientSecretCredential)
    assert client.base_path == "/users/u@x.com"


# --- command wiring ---


def test_get_command_prints_subject():
    fake = FakeClient()
    result = run(["get", "M1"], fake)
    assert result.exit_code == 0
    assert "Hello" in result.output
    assert ("get_email", "M1") in fake.calls


def test_list_command_prints_messages():
    fake = FakeClient()
    result = run(["list"], fake)
    assert result.exit_code == 0
    assert "One" in result.output
    assert "Two" in result.output


def test_search_command_passes_filters():
    fake = FakeClient()
    result = run(["search", "--subject-contains", "ABC", "--has-attachments"], fake)
    assert result.exit_code == 0
    name, kwargs = fake.calls[0][0], fake.calls[0][1]
    assert name == "search_email"
    assert kwargs["subject_contains"] == "ABC"
    assert kwargs["has_attachments"] is True


def test_send_command_calls_send_email():
    fake = FakeClient()
    result = run(["send", "--to", "bob@x.com", "--subject", "S", "--body", "B"], fake)
    assert result.exit_code == 0
    call = fake.calls[0]
    assert call[0] == "send_email"
    assert call[1] == ["bob@x.com"]
    assert call[2] == "S"


def test_delete_command_permanent_flag():
    fake = FakeClient()
    result = run(["delete", "M1", "--permanent"], fake)
    assert result.exit_code == 0
    assert ("delete_email", "M1", True) in fake.calls


def test_folders_command_lists():
    fake = FakeClient()
    result = run(["folders"], fake)
    assert result.exit_code == 0
    assert "Inbox" in result.output


def test_move_command():
    fake = FakeClient()
    result = run(["move", "M1", "archive"], fake)
    assert result.exit_code == 0
    assert ("move_email", "M1", "archive") in fake.calls


def test_mkdir_command():
    fake = FakeClient()
    result = run(["mkdir", "Reports"], fake)
    assert result.exit_code == 0
    assert ("create_folder", "Reports", None) in fake.calls
