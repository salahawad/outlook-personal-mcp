from __future__ import annotations


def register(mcp, client):
    @mcp.tool()
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

    @mcp.tool()
    async def create_folder(name: str, parent_folder_id: str | None = None) -> dict:
        """Create a mail folder, optionally nested under parent_folder_id."""
        path = (
            f"/me/mailFolders/{parent_folder_id}/childFolders"
            if parent_folder_id
            else "/me/mailFolders"
        )
        f = await client.post(path, json={"displayName": name})
        return {"id": f["id"], "name": f["displayName"]}

    @mcp.tool()
    async def rename_folder(folder_id: str, new_name: str) -> dict:
        """Rename a mail folder."""
        f = await client.patch(
            f"/me/mailFolders/{folder_id}", json={"displayName": new_name}
        )
        return {"id": f["id"], "name": f["displayName"]}

    @mcp.tool()
    async def delete_folder(folder_id: str) -> dict:
        """Delete a mail folder (moves it to Deleted Items)."""
        await client.delete(f"/me/mailFolders/{folder_id}")
        return {"deleted": folder_id}
