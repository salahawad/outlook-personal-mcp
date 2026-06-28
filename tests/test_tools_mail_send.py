import httpx
import respx
import pytest
import json
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.config import Settings
from outlook_personal_mcp.tools import mail

BASE = "https://graph.microsoft.com/v1.0"

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    mail.register(mcp, GraphClient(StubTokens(), backoff_base=0), Settings(client_id="abc"))
    return mcp


def _draft_routes(action, draft_id="DRAFT1", body=None):
    """Mock the createReply/createReplyAll/createForward -> patch -> send sequence."""
    if body is None:
        body = {"contentType": "HTML", "content": "<html><body><hr>quoted original</body></html>"}
    create = respx.post(f"{BASE}/me/messages/m1/{action}").mock(
        return_value=httpx.Response(201, json={"id": draft_id, "body": body}))
    patch = respx.patch(f"{BASE}/me/messages/{draft_id}").mock(
        return_value=httpx.Response(200, json={"id": draft_id}))
    send = respx.post(f"{BASE}/me/messages/{draft_id}/send").mock(
        return_value=httpx.Response(202))
    return create, patch, send


@respx.mock
async def test_send_mail_builds_graph_payload(ctx):
    route = respx.post(f"{BASE}/me/sendMail").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("send_mail", {"to": ["a@x.com"], "subject": "S", "body": "B"})
    body = json.loads(route.calls.last.request.read())
    assert body["message"]["toRecipients"][0]["emailAddress"]["address"] == "a@x.com"


@respx.mock
async def test_send_mail_save_to_sent_false(ctx):
    route = respx.post(f"{BASE}/me/sendMail").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("send_mail", {"to": ["a@x.com"], "subject": "S", "body": "B",
                                      "save_to_sent": False})
    assert json.loads(route.calls.last.request.read())["saveToSentItems"] is False


# ---------- reply: createReply -> PATCH HTML body -> send ----------

@respx.mock
async def test_reply_preserves_line_breaks(ctx):
    _, patch, send = _draft_routes("createReply")
    await ctx.call_tool("reply", {"message_id": "m1", "comment": "Bonjour,\n\nLigne deux."})
    content = json.loads(patch.calls.last.request.read())["body"]["content"]
    assert "Bonjour,<br><br>Ligne deux." in content   # newlines became <br>
    assert "quoted original" in content                # the quoted original is kept
    assert send.called


@respx.mock
async def test_reply_escapes_html_in_comment(ctx):
    _, patch, _ = _draft_routes("createReply")
    await ctx.call_tool("reply", {"message_id": "m1", "comment": "a < b & c"})
    content = json.loads(patch.calls.last.request.read())["body"]["content"]
    assert "a &lt; b &amp; c" in content


@respx.mock
async def test_reply_all_uses_create_reply_all(ctx):
    create, _, send = _draft_routes("createReplyAll")
    await ctx.call_tool("reply", {"message_id": "m1", "comment": "ok", "reply_all": True})
    assert create.called and send.called


@respx.mock
async def test_reply_plain_text_draft_keeps_newlines(ctx):
    _, patch, _ = _draft_routes(
        "createReply", body={"contentType": "Text", "content": "quoted"})
    await ctx.call_tool("reply", {"message_id": "m1", "comment": "line1\nline2"})
    body = json.loads(patch.calls.last.request.read())["body"]
    assert body["contentType"] == "Text"
    assert "line1\nline2" in body["content"]
    assert "<br>" not in body["content"]


# ---------- forward: createForward -> PATCH recipients+body -> send ----------

@respx.mock
async def test_forward_sets_recipients_and_formats_comment(ctx):
    _, patch, send = _draft_routes("createForward")
    await ctx.call_tool("forward", {"message_id": "m1", "to": ["b@x.com"],
                                    "comment": "see\nbelow"})
    patched = json.loads(patch.calls.last.request.read())
    assert patched["toRecipients"][0]["emailAddress"]["address"] == "b@x.com"
    assert "see<br>below" in patched["body"]["content"]
    assert send.called
