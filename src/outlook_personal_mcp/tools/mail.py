from __future__ import annotations

from ..models import SendMailInput, shape_message

_LIST_SELECT = "id,subject,from,isRead,receivedDateTime,bodyPreview,hasAttachments"


def register(mcp, client, settings):
    # ---------- read ----------
    @mcp.tool()
    async def list_messages(
        folder_id: str | None = None,
        top: int = 25,
        skip: int = 0,
        unread_only: bool = False,
    ) -> list[dict]:
        """List messages (newest first). Optionally restrict to a folder or unread only."""
        path = f"/me/mailFolders/{folder_id}/messages" if folder_id else "/me/messages"
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

    @mcp.tool()
    async def search_messages(query: str, top: int = 25) -> list[dict]:
        """Full-text search across the mailbox (Graph $search)."""
        params = {"$search": f'"{query}"', "$top": top, "$select": _LIST_SELECT}
        data = await client.get("/me/messages", params=params)
        return [shape_message(m) for m in data.get("value", [])]

    @mcp.tool()
    async def get_message(message_id: str, include_body: bool = True) -> dict:
        """Get a single message; include_body returns the full body text/HTML."""
        sel = _LIST_SELECT + (",body,toRecipients,ccRecipients" if include_body else "")
        m = await client.get(f"/me/messages/{message_id}", params={"$select": sel})
        out = shape_message(m)
        if include_body:
            out["body"] = (m.get("body") or {}).get("content")
            out["to"] = [r["emailAddress"]["address"] for r in m.get("toRecipients", [])]
        return out

    @mcp.tool()
    async def list_attachments(message_id: str) -> list[dict]:
        """List a message's attachments (id, name, size, contentType)."""
        data = await client.get(
            f"/me/messages/{message_id}/attachments",
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

    @mcp.tool()
    async def download_attachment(
        message_id: str, attachment_id: str, save_path: str
    ) -> dict:
        """Download an attachment to a local file path."""
        raw = await client.get(
            f"/me/messages/{message_id}/attachments/{attachment_id}/$value",
            raw=True,
        )
        with open(save_path, "wb") as fh:
            fh.write(raw)
        return {"saved": save_path, "bytes": len(raw)}

    # ---------- send ----------
    @mcp.tool()
    async def send_mail(
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        html: bool = False,
    ) -> dict:
        """Send an email."""
        payload = SendMailInput(
            to=to,
            subject=subject,
            body=body,
            cc=cc or [],
            bcc=bcc or [],
            html=html,
        ).to_graph()
        await client.post("/me/sendMail", json=payload)
        return {"sent": True, "to": to, "subject": subject}

    @mcp.tool()
    async def reply(
        message_id: str, comment: str, reply_all: bool = False
    ) -> dict:
        """Reply to a message (set reply_all to reply to everyone)."""
        action = "replyAll" if reply_all else "reply"
        await client.post(
            f"/me/messages/{message_id}/{action}", json={"comment": comment}
        )
        return {"replied": message_id, "reply_all": reply_all}

    @mcp.tool()
    async def forward(
        message_id: str, to: list[str], comment: str = ""
    ) -> dict:
        """Forward a message to recipients with an optional comment."""
        payload = {
            "comment": comment,
            "toRecipients": [{"emailAddress": {"address": a}} for a in to],
        }
        await client.post(f"/me/messages/{message_id}/forward", json=payload)
        return {"forwarded": message_id, "to": to}

    # ---------- organize ----------
    @mcp.tool()
    async def move_message(
        message_id: str, destination_folder_id: str
    ) -> dict:
        """Move a message to another folder."""
        m = await client.post(
            f"/me/messages/{message_id}/move",
            json={"destinationId": destination_folder_id},
        )
        return {"id": (m or {}).get("id"), "moved_to": destination_folder_id}

    @mcp.tool()
    async def copy_message(
        message_id: str, destination_folder_id: str
    ) -> dict:
        """Copy a message to another folder."""
        m = await client.post(
            f"/me/messages/{message_id}/copy",
            json={"destinationId": destination_folder_id},
        )
        return {"id": (m or {}).get("id"), "copied_to": destination_folder_id}

    @mcp.tool()
    async def mark_read(message_id: str, read: bool = True) -> dict:
        """Mark a message read (read=True) or unread (read=False)."""
        await client.patch(f"/me/messages/{message_id}", json={"isRead": read})
        return {"id": message_id, "is_read": read}

    @mcp.tool()
    async def flag_message(message_id: str, flagged: bool = True) -> dict:
        """Flag or unflag a message."""
        status = "flagged" if flagged else "notFlagged"
        await client.patch(
            f"/me/messages/{message_id}", json={"flag": {"flagStatus": status}}
        )
        return {"id": message_id, "flagged": flagged}

    @mcp.tool()
    async def delete_message(message_id: str) -> dict:
        """Delete a message (moves it to Deleted Items; reversible)."""
        await client.delete(f"/me/messages/{message_id}")
        return {"deleted": message_id, "permanent": False}

    if settings.allow_permanent_delete:
        # NOTE: POST /me/messages/{id}/permanentDelete IS available on Graph v1.0
        # (not beta-only). Confirmed at:
        # https://learn.microsoft.com/en-us/graph/api/message-permanentdelete?view=graph-rest-1.0
        @mcp.tool()
        async def permanent_delete(message_id: str) -> dict:
            """PERMANENTLY delete a message (irreversible). Enabled via env flag."""
            await client.post(f"/me/messages/{message_id}/permanentDelete")
            return {"deleted": message_id, "permanent": True}
