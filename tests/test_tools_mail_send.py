import httpx
import respx
import pytest
import json
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
async def test_send_mail_builds_graph_payload(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/sendMail").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("send_mail", {"to": ["a@x.com"], "subject": "S", "body": "B"})
    body = json.loads(route.calls.last.request.read())
    assert body["message"]["toRecipients"][0]["emailAddress"]["address"] == "a@x.com"

@respx.mock
async def test_reply_all(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages/m1/replyAll").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("reply", {"message_id": "m1", "comment": "ok", "reply_all": True})
    assert route.called

@respx.mock
async def test_forward_passes_recipients(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages/m1/forward").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("forward", {"message_id": "m1", "to": ["b@x.com"], "comment": "fyi"})
    body = json.loads(route.calls.last.request.read())
    assert body["toRecipients"][0]["emailAddress"]["address"] == "b@x.com"
    assert body["comment"] == "fyi"
