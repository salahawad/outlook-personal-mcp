import httpx
import respx
import json
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.config import Settings
from outlook_personal_mcp.tools import mail

class StubTokens:
    def get_token(self): return "T"

def _ctx(allow_perm=False):
    mcp = FastMCP("test")
    mail.register(mcp, GraphClient(StubTokens(), backoff_base=0),
                  Settings(client_id="abc", allow_permanent_delete=allow_perm))
    return mcp

@respx.mock
async def test_move_message():
    ctx = _ctx()
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages/m1/move").mock(
        return_value=httpx.Response(201, json={"id": "m1b"}))
    await ctx.call_tool("move_message", {"message_id": "m1", "destination_folder_id": "f2"})
    assert json.loads(route.calls.last.request.read())["destinationId"] == "f2"

@respx.mock
async def test_mark_read():
    ctx = _ctx()
    route = respx.patch("https://graph.microsoft.com/v1.0/me/messages/m1").mock(
        return_value=httpx.Response(200, json={"id": "m1"}))
    await ctx.call_tool("mark_read", {"message_id": "m1", "read": True})
    assert json.loads(route.calls.last.request.read())["isRead"] is True

@respx.mock
async def test_delete_message():
    ctx = _ctx()
    route = respx.delete("https://graph.microsoft.com/v1.0/me/messages/m1").mock(
        return_value=httpx.Response(204))
    await ctx.call_tool("delete_message", {"message_id": "m1"})
    assert route.called

async def test_permanent_delete_absent_by_default():
    ctx = _ctx(allow_perm=False)
    names = {t.name for t in await ctx.list_tools()}
    assert "permanent_delete" not in names
    assert "delete_message" in names

async def test_permanent_delete_present_when_enabled():
    ctx = _ctx(allow_perm=True)
    names = {t.name for t in await ctx.list_tools()}
    assert "permanent_delete" in names
