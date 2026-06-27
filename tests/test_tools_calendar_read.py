import httpx
import respx
import pytest
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.tools import calendar

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    calendar.register(mcp, GraphClient(StubTokens(), backoff_base=0))
    return mcp

@respx.mock
async def test_list_events_window_uses_calendarview(ctx):
    route = respx.get("https://graph.microsoft.com/v1.0/me/calendarView").mock(
        return_value=httpx.Response(200, json={"value": [
            {"id": "e1", "subject": "Sync",
             "start": {"dateTime": "2026-06-27T09:00:00", "timeZone": "UTC"},
             "end": {"dateTime": "2026-06-27T09:30:00", "timeZone": "UTC"}}]}))
    res = await ctx.call_tool("list_events",
                              {"start": "2026-06-27T00:00:00Z", "end": "2026-06-28T00:00:00Z"})
    assert route.calls.last.request.url.params["startDateTime"] == "2026-06-27T00:00:00Z"
    assert "Sync" in str(res)

@respx.mock
async def test_list_events_no_window_uses_events(ctx):
    route = respx.get("https://graph.microsoft.com/v1.0/me/events").mock(
        return_value=httpx.Response(200, json={"value": []}))
    await ctx.call_tool("list_events", {})
    assert route.called

@respx.mock
async def test_get_event(ctx):
    respx.get("https://graph.microsoft.com/v1.0/me/events/e1").mock(
        return_value=httpx.Response(200, json={"id": "e1", "subject": "1:1",
            "start": {"dateTime": "x", "timeZone": "UTC"},
            "end": {"dateTime": "y", "timeZone": "UTC"}}))
    res = await ctx.call_tool("get_event", {"event_id": "e1"})
    assert "1:1" in str(res)
