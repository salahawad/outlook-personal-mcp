from __future__ import annotations

import base64
import os


def _recips(addrs):
    return [{"emailAddress": {"address": a}} for a in addrs]


def register(mcp, client):
    @mcp.tool()
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

    @mcp.tool()
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
        d = await client.patch(f"/me/messages/{draft_id}", json=patch)
        return {"id": d["id"]}

    @mcp.tool()
    async def add_attachment(draft_id: str, file_path: str) -> dict:
        """Attach a local file to a draft."""
        with open(file_path, "rb") as fh:
            content = base64.b64encode(fh.read()).decode("ascii")
        payload = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": os.path.basename(file_path),
            "contentBytes": content,
        }
        a = await client.post(f"/me/messages/{draft_id}/attachments", json=payload)
        return {"attachment_id": a.get("id"), "name": a.get("name")}

    @mcp.tool()
    async def send_draft(draft_id: str) -> dict:
        """Send an existing draft."""
        await client.post(f"/me/messages/{draft_id}/send")
        return {"sent": draft_id}
