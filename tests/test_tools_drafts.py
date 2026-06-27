import httpx
import json

import pytest
import respx
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.tools import drafts

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    drafts.register(mcp, GraphClient(StubTokens(), backoff_base=0))
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
