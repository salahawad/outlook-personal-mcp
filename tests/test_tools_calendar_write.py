import httpx
import respx
import pytest
import json
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
async def test_create_event_builds_payload(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/events").mock(
        return_value=httpx.Response(201, json={"id": "e9", "subject": "Demo",
            "start": {"dateTime": "2026-07-01T15:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-07-01T15:30:00", "timeZone": "UTC"}}))
    await ctx.call_tool("create_event", {"subject": "Demo",
        "start": "2026-07-01T15:00:00", "end": "2026-07-01T15:30:00",
        "attendees": ["a@x.com"], "is_online_meeting": True})
    body = json.loads(route.calls.last.request.read())
    assert body["subject"] == "Demo"
    assert body["isOnlineMeeting"] is True
    assert body["attendees"][0]["emailAddress"]["address"] == "a@x.com"

@respx.mock
async def test_respond_event_accept(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/events/e1/accept").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("respond_event", {"event_id": "e1", "response": "accept", "comment": "yes"})
    body = json.loads(route.calls.last.request.read())
    assert body["sendResponse"] is True and body["comment"] == "yes"

@respx.mock
async def test_respond_event_tentative_maps_action(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/events/e1/tentativelyAccept").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("respond_event", {"event_id": "e1", "response": "tentative"})
    assert route.called

@respx.mock
async def test_find_availability(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/calendar/getSchedule").mock(
        return_value=httpx.Response(200, json={"value": [{"scheduleId": "a@x.com",
            "availabilityView": "000"}]}))
    res = await ctx.call_tool("find_availability", {"emails": ["a@x.com"],
        "start": "2026-07-01T09:00:00", "end": "2026-07-01T17:00:00"})
    assert json.loads(route.calls.last.request.read())["schedules"] == ["a@x.com"]
    assert "a@x.com" in str(res)

@respx.mock
async def test_update_event_partial_patch(ctx):
    route = respx.patch("https://graph.microsoft.com/v1.0/me/events/e1").mock(
        return_value=httpx.Response(200, json={"id": "e1", "subject": "X",
            "start": {"dateTime": "x", "timeZone": "UTC"},
            "end": {"dateTime": "y", "timeZone": "UTC"}}))
    await ctx.call_tool("update_event", {"event_id": "e1", "start": "2026-07-02T10:00:00",
                                         "time_zone": "UTC"})
    body = json.loads(route.calls.last.request.read())
    assert body["start"] == {"dateTime": "2026-07-02T10:00:00", "timeZone": "UTC"}
    assert "subject" not in body
    assert "end" not in body

async def test_respond_event_rejects_bad_response(ctx):
    from mcp.server.fastmcp.exceptions import ToolError
    with pytest.raises(ToolError) as ei:
        await ctx.call_tool("respond_event", {"event_id": "e1", "response": "maybe"})
    assert "accept" in str(ei.value)
    assert "decline" in str(ei.value)
    assert "tentative" in str(ei.value)
