from __future__ import annotations

from mcp.types import ToolAnnotations

from ..models import SendMailInput, shape_message
from ..safety import ensure_size_within_limit, graph_id, resolve_write_file

_LIST_SELECT = "id,subject,from,isRead,receivedDateTime,bodyPreview,hasAttachments"


def register(mcp, client, settings):
    # ---------- read ----------
    @mcp.tool(
        annotations=ToolAnnotations(
            title="List mail messages",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def list_messages(
        folder_id: str | None = None,
        top: int = 25,
        skip: int = 0,
        unread_only: bool = False,
    ) -> list[dict]:
        """List messages (newest first). Optionally restrict to a folder or unread only."""
        path = f"/me/mailFolders/{graph_id(folder_id, name='folder_id')}/messages" if folder_id else "/me/messages"
        params = {
            "$top": top,
            "$skip": skip,
            "$select": _LIST_SELECT,
            "$orderby": "receivedDateTime desc",
        }
        if unread_only:
            params["$filter"] = "isRead eq false"
        data = await client.get(path, params=params)
        return [shape_message(m) for m in data.get("value", [])]

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Search mail messages",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def search_messages(query: str, top: int = 25) -> list[dict]:
        """Full-text search across the mailbox (Graph $search)."""
        params = {"$search": f'"{query}"', "$top": top, "$select": _LIST_SELECT}
        data = await client.get("/me/messages", params=params)
        return [shape_message(m) for m in data.get("value", [])]

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get a mail message",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def get_message(message_id: str, include_body: bool = True) -> dict:
        """Get a single message; include_body returns the full body text/HTML."""
        sel = _LIST_SELECT + (",body,toRecipients,ccRecipients" if include_body else "")
        mid = graph_id(message_id, name="message_id")
        m = await client.get(f"/me/messages/{mid}", params={"$select": sel})
        out = shape_message(m)
        if include_body:
            out["body"] = (m.get("body") or {}).get("content")
            out["to"] = [r["emailAddress"]["address"] for r in m.get("toRecipients", [])]
        return out

    @mcp.tool(
        annotations=ToolAnnotations(
            title="List message attachments",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def list_attachments(message_id: str) -> list[dict]:
        """List a message's attachments (id, name, size, contentType)."""
        mid = graph_id(message_id, name="message_id")
        data = await client.get(
            f"/me/messages/{mid}/attachments",
            params={"$select": "id,name,size,contentType"},
        )
        return [
            {
                "id": a["id"],
                "name": a.get("name"),
                "size": a.get("size"),
                "content_type": a.get("contentType"),
            }
            for a in data.get("value", [])
        ]

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Download an attachment to a local file",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def download_attachment(
        message_id: str, attachment_id: str, save_path: str
    ) -> dict:
        """Download an attachment to a local file path."""
        mid = graph_id(message_id, name="message_id")
        aid = graph_id(attachment_id, name="attachment_id")
        target = resolve_write_file(settings, save_path)
        raw = await client.get(
            f"/me/messages/{mid}/attachments/{aid}/$value",
            raw=True,
        )
        ensure_size_within_limit(settings, raw, label="attachment")
        with open(target, "xb") as fh:
            fh.write(raw)
        return {"saved": str(target), "bytes": len(raw)}

    # ---------- send ----------
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Send an email",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def send_mail(
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        html: bool = False,
        save_to_sent: bool = True,
    ) -> dict:
        """Send an email. Set save_to_sent=False to skip saving to Sent Items."""
        payload = SendMailInput(
            to=to,
            subject=subject,
            body=body,
            cc=cc or [],
            bcc=bcc or [],
            html=html,
            save_to_sent=save_to_sent,
        ).to_graph()
        await client.post("/me/sendMail", json=payload)
        return {"sent": True, "to": to, "subject": subject}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Reply to a message",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def reply(
        message_id: str, comment: str, reply_all: bool = False
    ) -> dict:
        """Reply to a message (set reply_all to reply to everyone)."""
        action = "replyAll" if reply_all else "reply"
        mid = graph_id(message_id, name="message_id")
        await client.post(
            f"/me/messages/{mid}/{action}", json={"comment": comment}
        )
        return {"replied": message_id, "reply_all": reply_all}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Forward a message",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def forward(
        message_id: str, to: list[str], comment: str = ""
    ) -> dict:
        """Forward a message to recipients with an optional comment."""
        mid = graph_id(message_id, name="message_id")
        payload = {
            "comment": comment,
            "toRecipients": [{"emailAddress": {"address": a}} for a in to],
        }
        await client.post(f"/me/messages/{mid}/forward", json=payload)
        return {"forwarded": message_id, "to": to}

    # ---------- organize ----------
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Move a message to a folder",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=True,
        )
    )
    async def move_message(
        message_id: str, destination_folder_id: str
    ) -> dict:
        """Move a message to another folder."""
        mid = graph_id(message_id, name="message_id")
        m = await client.post(
            f"/me/messages/{mid}/move",
            json={"destinationId": destination_folder_id},
        )
        return {"id": (m or {}).get("id"), "moved_to": destination_folder_id}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Copy a message to a folder",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def copy_message(
        message_id: str, destination_folder_id: str
    ) -> dict:
        """Copy a message to another folder."""
        mid = graph_id(message_id, name="message_id")
        m = await client.post(
            f"/me/messages/{mid}/copy",
            json={"destinationId": destination_folder_id},
        )
        return {"id": (m or {}).get("id"), "copied_to": destination_folder_id}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Mark a message read or unread",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=True,
        )
    )
    async def mark_read(message_id: str, read: bool = True) -> dict:
        """Mark a message read (read=True) or unread (read=False)."""
        mid = graph_id(message_id, name="message_id")
        await client.patch(f"/me/messages/{mid}", json={"isRead": read})
        return {"id": message_id, "is_read": read}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Flag or unflag a message",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=True,
        )
    )
    async def flag_message(message_id: str, flagged: bool = True) -> dict:
        """Flag or unflag a message."""
        status = "flagged" if flagged else "notFlagged"
        mid = graph_id(message_id, name="message_id")
        await client.patch(
            f"/me/messages/{mid}", json={"flag": {"flagStatus": status}}
        )
        return {"id": message_id, "flagged": flagged}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete a message (moves to Deleted Items)",
            readOnlyHint=False,
            destructiveHint=True,
            openWorldHint=True,
        )
    )
    async def delete_message(message_id: str) -> dict:
        """Delete a message (moves it to Deleted Items; reversible)."""
        mid = graph_id(message_id, name="message_id")
        await client.delete(f"/me/messages/{mid}")
        return {"deleted": message_id, "permanent": False}

    if settings.allow_permanent_delete:
        # NOTE: POST /me/messages/{id}/permanentDelete IS available on Graph v1.0
        # (not beta-only). Confirmed at:
        # https://learn.microsoft.com/en-us/graph/api/message-permanentdelete?view=graph-rest-1.0
        @mcp.tool(
            annotations=ToolAnnotations(
                title="Permanently delete a message (irreversible)",
                readOnlyHint=False,
                destructiveHint=True,
                openWorldHint=True,
            )
        )
        async def permanent_delete(message_id: str) -> dict:
            """PERMANENTLY delete a message (irreversible). Enabled via env flag."""
            mid = graph_id(message_id, name="message_id")
            await client.post(f"/me/messages/{mid}/permanentDelete")
            return {"deleted": message_id, "permanent": True}
