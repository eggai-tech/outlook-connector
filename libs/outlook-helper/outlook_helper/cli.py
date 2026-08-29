"""A thin click CLI over :class:`~outlook_helper.client.OutlookClient`.

Config comes from options (with environment-variable fallbacks). The client is
built lazily on first use so ``--help`` and tests that inject a client never
require credentials.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import click

from outlook_helper.auth import ClientSecretCredential, DeviceCodeCredential
from outlook_helper.client import OutlookClient
from outlook_helper.config import AppOnlyConfig, DelegatedConfig
from outlook_helper.schemas import OutlookMessage


def build_client(params: dict[str, Any]) -> OutlookClient:
    if not params.get("client_id"):
        raise click.UsageError("--client-id is required (or set OUTLOOK_CLIENT_ID)")
    if params["auth"] == "app-only":
        if not params.get("client_secret"):
            raise click.UsageError("app-only auth requires --client-secret")
        if not params.get("tenant_id"):
            raise click.UsageError("app-only auth requires --tenant-id")
        credential = ClientSecretCredential(
            AppOnlyConfig(
                client_id=params["client_id"],
                tenant_id=params["tenant_id"],
                client_secret=params["client_secret"],
            )
        )
    else:
        cache_path = params.get("cache_path")
        credential = DeviceCodeCredential(
            DelegatedConfig(
                client_id=params["client_id"],
                tenant_id=params.get("tenant_id") or "common",
                cache_path=Path(cache_path) if cache_path else None,
            )
        )
    return OutlookClient(credential, mailbox=params.get("mailbox"))


def get_client(ctx: click.Context) -> OutlookClient:
    obj = ctx.ensure_object(dict)
    if "client" not in obj:
        obj["client"] = build_client(obj["params"])
    return obj["client"]


def _format_message(msg: OutlookMessage) -> str:
    sender = msg.from_.address if msg.from_ else "?"
    return f"{msg.id}\t{sender}\t{msg.subject or '(no subject)'}"


@click.group()
@click.option("--client-id", envvar="OUTLOOK_CLIENT_ID")
@click.option("--tenant-id", envvar="OUTLOOK_TENANT_ID")
@click.option("--mailbox", envvar="OUTLOOK_MAILBOX")
@click.option(
    "--auth",
    type=click.Choice(["delegated", "app-only"]),
    default="delegated",
    show_default=True,
)
@click.option("--client-secret", envvar="OUTLOOK_CLIENT_SECRET")
@click.option(
    "--cache-path",
    envvar="OUTLOOK_CACHE_PATH",
    help="Where to persist the delegated token cache.",
)
@click.pass_context
def main(ctx, client_id, tenant_id, mailbox, auth, client_secret, cache_path):
    """Work with M365 email through the Microsoft Graph API."""
    obj = ctx.ensure_object(dict)
    obj["params"] = {
        "client_id": client_id,
        "tenant_id": tenant_id,
        "mailbox": mailbox,
        "auth": auth,
        "client_secret": client_secret,
        "cache_path": cache_path,
    }


@main.command()
@click.pass_context
def login(ctx):
    """Sign in (delegated) and prime the token cache."""
    get_client(ctx).credential.get_token()
    click.echo("Authenticated.")


@main.command()
@click.argument("message_id")
@click.pass_context
def get(ctx, message_id):
    """Fetch a single message by id."""
    msg = get_client(ctx).get_email(message_id)
    click.echo(_format_message(msg))
    if msg.body and msg.body.content:
        click.echo()
        click.echo(msg.body.content)


@main.command(name="list")
@click.option("--folder", default="inbox", show_default=True)
@click.option("--top", type=int, default=None)
@click.pass_context
def list_cmd(ctx, folder, top):
    """List messages in a folder (newest first)."""
    for msg in get_client(ctx).list_messages(folder=folder, top=top):
        click.echo(_format_message(msg))


@main.command()
@click.option("--sender")
@click.option("--subject-contains")
@click.option("--since")
@click.option("--until")
@click.option("--unread/--no-unread", default=None)
@click.option("--has-attachments", is_flag=True, default=None)
@click.option("--folder")
@click.option("--top", type=int, default=None)
@click.pass_context
def search(
    ctx, sender, subject_contains, since, until, unread, has_attachments, folder, top
):
    """Search messages with precise filters."""
    for msg in get_client(ctx).search_email(
        sender=sender,
        subject_contains=subject_contains,
        since=since,
        until=until,
        unread=unread,
        has_attachments=has_attachments,
        folder=folder,
        top=top,
    ):
        click.echo(_format_message(msg))


@main.command()
@click.option("--to", "to", multiple=True, required=True)
@click.option("--subject", required=True)
@click.option("--body", required=True)
@click.option("--cc", multiple=True)
@click.option("--bcc", multiple=True)
@click.option("--attach", "attachments", multiple=True, type=click.Path(exists=True))
@click.option("--html", is_flag=True)
@click.pass_context
def send(ctx, to, subject, body, cc, bcc, attachments, html):
    """Send a new message."""
    get_client(ctx).send_email(
        list(to),
        subject,
        body,
        cc=list(cc) or None,
        bcc=list(bcc) or None,
        attachments=list(attachments) or None,
        html=html,
    )
    click.echo("Sent.")


@main.command()
@click.argument("message_id")
@click.option("--body", required=True)
@click.option("--reply-all", is_flag=True)
@click.option("--attach", "attachments", multiple=True, type=click.Path(exists=True))
@click.option("--html", is_flag=True)
@click.pass_context
def reply(ctx, message_id, body, reply_all, attachments, html):
    """Reply to a message."""
    get_client(ctx).reply(
        message_id,
        body,
        reply_all=reply_all,
        attachments=list(attachments) or None,
        html=html,
    )
    click.echo("Replied.")


@main.command(name="draft")
@click.option("--to", "to", multiple=True, required=True)
@click.option("--subject", required=True)
@click.option("--body", required=True)
@click.option("--cc", multiple=True)
@click.option("--attach", "attachments", multiple=True, type=click.Path(exists=True))
@click.option("--html", is_flag=True)
@click.pass_context
def draft(ctx, to, subject, body, cc, attachments, html):
    """Create a draft (does not send)."""
    msg = get_client(ctx).create_draft(
        list(to),
        subject,
        body,
        cc=list(cc) or None,
        attachments=list(attachments) or None,
        html=html,
    )
    click.echo(f"Draft created: {msg.id}")


@main.command(name="send-draft")
@click.argument("message_id")
@click.pass_context
def send_draft(ctx, message_id):
    """Send a previously created draft."""
    get_client(ctx).send_draft(message_id)
    click.echo("Sent.")


@main.command()
@click.argument("message_id")
@click.argument("attachment_id")
@click.argument("dest", type=click.Path())
@click.pass_context
def download(ctx, message_id, attachment_id, dest):
    """Download an attachment to DEST."""
    path = get_client(ctx).download_attachment(message_id, attachment_id, dest)
    click.echo(f"Saved {path}")


@main.command()
@click.pass_context
def folders(ctx):
    """List mail folders."""
    for folder in get_client(ctx).list_folders():
        click.echo(f"{folder.id}\t{folder.display_name}")


@main.command()
@click.argument("name")
@click.option("--parent")
@click.pass_context
def mkdir(ctx, name, parent):
    """Create a mail folder."""
    folder = get_client(ctx).create_folder(name, parent=parent)
    click.echo(f"Created folder: {folder.id}")


@main.command()
@click.argument("message_id")
@click.argument("dest_folder")
@click.pass_context
def move(ctx, message_id, dest_folder):
    """Move a message to another folder."""
    get_client(ctx).move_email(message_id, dest_folder)
    click.echo("Moved.")


@main.command()
@click.argument("message_id")
@click.option("--permanent", is_flag=True)
@click.pass_context
def delete(ctx, message_id, permanent):
    """Delete a message (soft delete unless --permanent)."""
    get_client(ctx).delete_email(message_id, permanent=permanent)
    click.echo("Deleted.")


if __name__ == "__main__":
    main()
