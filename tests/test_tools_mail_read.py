import httpx
import respx
import pytest
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.config import Settings
from outlook_personal_mcp.tools import mail

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx(tmp_path):
    mcp = FastMCP("test")
    settings = Settings(client_id="abc", file_root=str(tmp_path))
    mail.register(mcp, GraphClient(StubTokens(), backoff_base=0), settings)
    return mcp

@respx.mock
async def test_list_messages_shapes_output(ctx):
    respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
        return_value=httpx.Response(200, json={"value": [
            {"id": "m1", "subject": "Hi", "isRead": False,
             "from": {"emailAddress": {"address": "a@x.com"}},
             "receivedDateTime": "2026-06-27T10:00:00Z", "bodyPreview": "p",
             "hasAttachments": False}]}))
    res = await ctx.call_tool("list_messages", {"top": 5})
    assert "Hi" in str(res)

@respx.mock
async def test_search_messages_passes_search_param(ctx):
    route = respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
        return_value=httpx.Response(200, json={"value": []}))
    await ctx.call_tool("search_messages", {"query": "invoice"})
    assert route.calls.last.request.url.params["$search"] == '"invoice"'

@respx.mock
async def test_get_message_includes_body(ctx):
    respx.get("https://graph.microsoft.com/v1.0/me/messages/m1").mock(
        return_value=httpx.Response(200, json={"id": "m1", "subject": "S",
            "body": {"contentType": "text", "content": "HELLO BODY"},
            "toRecipients": [{"emailAddress": {"address": "a@x.com"}}]}))
    res = await ctx.call_tool("get_message", {"message_id": "m1"})
    assert "HELLO BODY" in str(res)

@respx.mock
async def test_get_message_encodes_id_path_segments(ctx):
    route = respx.get("https://graph.microsoft.com/v1.0/me/messages/m%2F1").mock(
        return_value=httpx.Response(200, json={"id": "m/1", "subject": "S"})
    )
    await ctx.call_tool("get_message", {"message_id": "m/1", "include_body": False})
    assert route.called

@respx.mock
async def test_download_attachment_writes_under_file_root(ctx, tmp_path):
    route = respx.get(
        "https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1/$value"
    ).mock(return_value=httpx.Response(200, content=b"hello"))
    result = await ctx.call_tool(
        "download_attachment",
        {"message_id": "m1", "attachment_id": "a1", "save_path": "nested/f.txt"},
    )
    saved = tmp_path / "nested" / "f.txt"
    assert route.called
    assert saved.read_bytes() == b"hello"
    assert str(saved) in result[0].text

async def test_download_attachment_rejects_path_outside_root(ctx, tmp_path):
    outside = tmp_path.parent / f"{tmp_path.name}-outside.txt"
    with pytest.raises(ToolError) as ei:
        await ctx.call_tool(
            "download_attachment",
            {"message_id": "m1", "attachment_id": "a1", "save_path": str(outside)},
        )
    assert "OUTLOOK_MCP_FILE_ROOT" in str(ei.value)

async def test_download_attachment_refuses_overwrite(ctx, tmp_path):
    existing = tmp_path / "exists.txt"
    existing.write_bytes(b"old")
    with pytest.raises(ToolError) as ei:
        await ctx.call_tool(
            "download_attachment",
            {"message_id": "m1", "attachment_id": "a1", "save_path": str(existing)},
        )
    assert "already exists" in str(ei.value)
    assert existing.read_bytes() == b"old"

@respx.mock
async def test_download_attachment_rejects_large_attachment(tmp_path):
    mcp = FastMCP("test")
    settings = Settings(client_id="abc", file_root=str(tmp_path), max_file_bytes=3)
    mail.register(mcp, GraphClient(StubTokens(), backoff_base=0), settings)
    respx.get(
        "https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1/$value"
    ).mock(return_value=httpx.Response(200, content=b"abcd"))
    with pytest.raises(ToolError) as ei:
        await mcp.call_tool(
            "download_attachment",
            {"message_id": "m1", "attachment_id": "a1", "save_path": "large.txt"},
        )
    assert "OUTLOOK_MCP_MAX_FILE_BYTES" in str(ei.value)
    assert not (tmp_path / "large.txt").exists()
