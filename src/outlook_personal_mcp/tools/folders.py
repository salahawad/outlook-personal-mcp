from __future__ import annotations

from mcp.types import ToolAnnotations

from ..safety import graph_id


def register(mcp, client):
    @mcp.tool(
        annotations=ToolAnnotations(
            title="List mail folders",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def list_folders() -> list[dict]:
        """List mail folders with unread/total counts."""
        data = await client.get("/me/mailFolders", params={"$top": 100})
        return [
            {
                "id": f["id"],
                "name": f["displayName"],
                "unread": f.get("unreadItemCount"),
                "total": f.get("totalItemCount"),
            }
            for f in data.get("value", [])
        ]

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create a mail folder",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def create_folder(name: str, parent_folder_id: str | None = None) -> dict:
        """Create a mail folder, optionally nested under parent_folder_id."""
        path = (
            f"/me/mailFolders/{graph_id(parent_folder_id, name='parent_folder_id')}/childFolders"
            if parent_folder_id
            else "/me/mailFolders"
        )
        f = await client.post(path, json={"displayName": name})
        return {"id": f["id"], "name": f["displayName"]}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Rename a mail folder",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def rename_folder(folder_id: str, new_name: str) -> dict:
        """Rename a mail folder."""
        fid = graph_id(folder_id, name="folder_id")
        f = await client.patch(
            f"/me/mailFolders/{fid}", json={"displayName": new_name}
        )
        return {"id": f["id"], "name": f["displayName"]}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete a mail folder",
            readOnlyHint=False,
            destructiveHint=True,
            openWorldHint=True,
        )
    )
    async def delete_folder(folder_id: str) -> dict:
        """Delete a mail folder (moves it to Deleted Items)."""
        fid = graph_id(folder_id, name="folder_id")
        await client.delete(f"/me/mailFolders/{fid}")
        return {"deleted": folder_id}
