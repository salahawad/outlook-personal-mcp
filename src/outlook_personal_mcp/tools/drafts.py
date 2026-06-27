from __future__ import annotations

import base64
import os

from mcp.types import ToolAnnotations

from ..safety import graph_id, resolve_read_file


def _recips(addrs):
    return [{"emailAddress": {"address": a}} for a in addrs]


def register(mcp, client, settings):
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create a draft message",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def create_draft(
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        html: bool = False,
    ) -> dict:
        """Create a draft message (not sent)."""
        msg = {
            "subject": subject,
            "body": {"contentType": "HTML" if html else "Text", "content": body},
            "toRecipients": _recips(to),
        }
        if cc:
            msg["ccRecipients"] = _recips(cc)
        d = await client.post("/me/messages", json=msg)
        return {"id": d["id"], "subject": subject}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Update a draft message",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=True,
        )
    )
    async def update_draft(
        draft_id: str,
        subject: str | None = None,
        body: str | None = None,
        html: bool = False,
    ) -> dict:
        """Update a draft's subject and/or body."""
        patch: dict = {}
        if subject is not None:
            patch["subject"] = subject
        if body is not None:
            patch["body"] = {"contentType": "HTML" if html else "Text", "content": body}
        did = graph_id(draft_id, name="draft_id")
        d = await client.patch(f"/me/messages/{did}", json=patch)
        return {"id": d["id"]}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add an attachment to a draft",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def add_attachment(draft_id: str, file_path: str) -> dict:
        """Attach a local file to a draft."""
        path = resolve_read_file(settings, file_path)
        with open(path, "rb") as fh:
            content = base64.b64encode(fh.read()).decode("ascii")
        did = graph_id(draft_id, name="draft_id")
        payload = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": os.path.basename(path),
            "contentBytes": content,
        }
        a = await client.post(f"/me/messages/{did}/attachments", json=payload)
        return {"attachment_id": a.get("id"), "name": a.get("name")}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Send a draft message",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def send_draft(draft_id: str) -> dict:
        """Send an existing draft."""
        did = graph_id(draft_id, name="draft_id")
        await client.post(f"/me/messages/{did}/send")
        return {"sent": draft_id}
