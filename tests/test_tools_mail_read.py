import httpx
import respx
import pytest
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.config import Settings
from outlook_personal_mcp.tools import mail

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    mail.register(mcp, GraphClient(StubTokens(), backoff_base=0), Settings(client_id="abc"))
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
