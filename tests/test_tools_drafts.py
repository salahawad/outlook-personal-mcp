import httpx
import json

import pytest
import respx
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError
from outlook_personal_mcp.config import Settings
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.tools import drafts

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx(tmp_path):
    mcp = FastMCP("test")
    settings = Settings(client_id="abc", file_root=str(tmp_path))
    drafts.register(mcp, GraphClient(StubTokens(), backoff_base=0), settings)
    return mcp

@respx.mock
async def test_create_draft(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages").mock(
        return_value=httpx.Response(201, json={"id": "d1"}))
    res = await ctx.call_tool("create_draft", {"to": ["a@x.com"], "subject": "S", "body": "B"})
    body = json.loads(route.calls.last.request.read())
    assert body["subject"] == "S"
    assert body["toRecipients"][0]["emailAddress"]["address"] == "a@x.com"
    assert "d1" in str(res)

@respx.mock
async def test_update_draft(ctx):
    route = respx.patch("https://graph.microsoft.com/v1.0/me/messages/d1").mock(
        return_value=httpx.Response(200, json={"id": "d1"}))
    await ctx.call_tool("update_draft", {"draft_id": "d1", "subject": "New"})
    assert json.loads(route.calls.last.request.read())["subject"] == "New"

@respx.mock
async def test_send_draft(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages/d1/send").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("send_draft", {"draft_id": "d1"})
    assert route.called

@respx.mock
async def test_add_attachment(ctx, tmp_path):
    p = tmp_path / "f.txt"
    p.write_bytes(b"hello")
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages/d1/attachments").mock(
        return_value=httpx.Response(201, json={"id": "att1", "name": "f.txt"}))
    await ctx.call_tool("add_attachment", {"draft_id": "d1", "file_path": str(p)})
    body = json.loads(route.calls.last.request.read())
    assert body["@odata.type"] == "#microsoft.graph.fileAttachment"
    assert body["name"] == "f.txt"
    assert body["contentBytes"] == "aGVsbG8="  # base64 of "hello"

async def test_add_attachment_rejects_file_outside_root(ctx, tmp_path):
    outside = tmp_path.parent / f"{tmp_path.name}-outside.txt"
    outside.write_bytes(b"secret")
    try:
        with pytest.raises(ToolError) as ei:
            await ctx.call_tool("add_attachment", {"draft_id": "d1", "file_path": str(outside)})
        assert "OUTLOOK_MCP_FILE_ROOT" in str(ei.value)
    finally:
        outside.unlink(missing_ok=True)

async def test_add_attachment_rejects_large_file(tmp_path):
    mcp = FastMCP("test")
    settings = Settings(client_id="abc", file_root=str(tmp_path), max_file_bytes=3)
    drafts.register(mcp, GraphClient(StubTokens(), backoff_base=0), settings)
    p = tmp_path / "large.txt"
    p.write_bytes(b"abcd")
    with pytest.raises(ToolError) as ei:
        await mcp.call_tool("add_attachment", {"draft_id": "d1", "file_path": str(p)})
    assert "OUTLOOK_MCP_MAX_FILE_BYTES" in str(ei.value)

async def test_add_attachment_rejects_symlink(ctx, tmp_path):
    target = tmp_path / "target.txt"
    target.write_bytes(b"hello")
    link = tmp_path / "link.txt"
    link.symlink_to(target)
    with pytest.raises(ToolError) as ei:
        await ctx.call_tool("add_attachment", {"draft_id": "d1", "file_path": str(link)})
    assert "symlinks" in str(ei.value)

@respx.mock
async def test_update_draft_omits_unset_fields(ctx):
    route = respx.patch("https://graph.microsoft.com/v1.0/me/messages/d1").mock(
        return_value=httpx.Response(200, json={"id": "d1"}))
    await ctx.call_tool("update_draft", {"draft_id": "d1", "subject": "New"})
    body = json.loads(route.calls.last.request.read())
    assert body["subject"] == "New"
    assert "body" not in body

@respx.mock
async def test_send_draft_returns_sent(ctx):
    respx.post("https://graph.microsoft.com/v1.0/me/messages/d1/send").mock(
        return_value=httpx.Response(202))
    result = await ctx.call_tool("send_draft", {"draft_id": "d1"})
    assert json.loads(result[0].text) == {"sent": "d1"}
