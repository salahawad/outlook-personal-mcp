"""
Tests for tools/folders.py — list_folders, create_folder, rename_folder, delete_folder.

FastMCP call_tool return shape (verified against mcp>=1.2):
  - Tool returning dict   → list[TextContent] (len 1), item.text is JSON string
  - Tool returning list   → tuple(list[TextContent], {"result": [...]})
    result[1] is the structured dict {"result": <original return value>}
    result[0] is list[TextContent], one TextContent per list element

Simplest universal assertion: `assert "expected_string" in str(result)` works for both shapes,
because str() over the tuple/list includes the TextContent .text which is JSON-encoded data.
For structural assertions on dict tools: json.loads(result[0].text) gives the dict.
"""

import json

import httpx
import pytest
import respx
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.tools import folders


class StubTokens:
    def get_token(self):
        return "T"


@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    client = GraphClient(StubTokens(), backoff_base=0)
    folders.register(mcp, client)
    return mcp


@respx.mock
async def test_list_folders(ctx):
    respx.get("https://graph.microsoft.com/v1.0/me/mailFolders").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "f1",
                        "displayName": "Inbox",
                        "unreadItemCount": 3,
                        "totalItemCount": 10,
                    }
                ]
            },
        )
    )
    result = await ctx.call_tool("list_folders", {})
    # result is tuple(list[TextContent], {"result": [...]})
    assert "Inbox" in str(result)
    structured = result[1]["result"]
    assert structured[0] == {"id": "f1", "name": "Inbox", "unread": 3, "total": 10}


@respx.mock
async def test_list_folders_uses_top_100(ctx):
    """Verifies $top=100 is passed in the query string."""
    route = respx.get("https://graph.microsoft.com/v1.0/me/mailFolders").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    await ctx.call_tool("list_folders", {})
    assert route.called
    assert route.calls.last.request.url.params.get("$top") == "100"


@respx.mock
async def test_create_folder(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/mailFolders").mock(
        return_value=httpx.Response(201, json={"id": "f2", "displayName": "Projects"})
    )
    result = await ctx.call_tool("create_folder", {"name": "Projects"})
    assert route.called
    assert json.loads(route.calls.last.request.read())["displayName"] == "Projects"
    # dict return → list[TextContent]
    data = json.loads(result[0].text)
    assert data == {"id": "f2", "name": "Projects"}


@respx.mock
async def test_create_folder_nested(ctx):
    """create_folder with parent_folder_id posts to childFolders sub-path."""
    route = respx.post(
        "https://graph.microsoft.com/v1.0/me/mailFolders/parentX/childFolders"
    ).mock(
        return_value=httpx.Response(
            201, json={"id": "f3", "displayName": "Nested"}
        )
    )
    result = await ctx.call_tool(
        "create_folder", {"name": "Nested", "parent_folder_id": "parentX"}
    )
    assert route.called
    data = json.loads(result[0].text)
    assert data == {"id": "f3", "name": "Nested"}


@respx.mock
async def test_rename_folder(ctx):
    route = respx.patch(
        "https://graph.microsoft.com/v1.0/me/mailFolders/f1"
    ).mock(
        return_value=httpx.Response(
            200, json={"id": "f1", "displayName": "Archive"}
        )
    )
    result = await ctx.call_tool("rename_folder", {"folder_id": "f1", "new_name": "Archive"})
    assert route.called
    assert json.loads(route.calls.last.request.read())["displayName"] == "Archive"
    data = json.loads(result[0].text)
    assert data == {"id": "f1", "name": "Archive"}


@respx.mock
async def test_delete_folder(ctx):
    route = respx.delete(
        "https://graph.microsoft.com/v1.0/me/mailFolders/f9"
    ).mock(return_value=httpx.Response(204))
    result = await ctx.call_tool("delete_folder", {"folder_id": "f9"})
    assert route.called
    data = json.loads(result[0].text)
    assert data == {"deleted": "f9"}
