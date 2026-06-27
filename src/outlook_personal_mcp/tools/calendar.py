from __future__ import annotations

from mcp.types import ToolAnnotations

from ..models import EventInput, shape_event

_SELECT = "id,subject,start,end,location,isAllDay,organizer"
_RESPONSE_ACTIONS = {"accept": "accept", "decline": "decline", "tentative": "tentativelyAccept"}

def register(mcp, client):
    @mcp.tool(
        annotations=ToolAnnotations(
            title="List calendars",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def list_calendars() -> list[dict]:
        """List the user's calendars."""
        data = await client.get("/me/calendars", params={"$select": "id,name,isDefaultCalendar"})
        return [{"id": c["id"], "name": c.get("name"),
                 "is_default": c.get("isDefaultCalendar")} for c in data.get("value", [])]

    @mcp.tool(
        annotations=ToolAnnotations(
            title="List calendar events",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def list_events(start: str | None = None, end: str | None = None,
                          top: int = 25) -> list[dict]:
        """List events. If start and end (ISO 8601) are given, returns that window via calendarView."""
        if start and end:
            data = await client.get("/me/calendarView",
                                    params={"startDateTime": start, "endDateTime": end,
                                            "$select": _SELECT, "$top": top,
                                            "$orderby": "start/dateTime"})
        else:
            data = await client.get("/me/events",
                                    params={"$select": _SELECT, "$top": top,
                                            "$orderby": "start/dateTime"})
        return [shape_event(e) for e in data.get("value", [])]

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Search calendar events",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def search_events(query: str, top: int = 25) -> list[dict]:
        """Search events by free text."""
        data = await client.get("/me/events",
                                params={"$search": f'"{query}"', "$select": _SELECT, "$top": top})
        return [shape_event(e) for e in data.get("value", [])]

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get a calendar event",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def get_event(event_id: str) -> dict:
        """Get one event including body and attendees."""
        e = await client.get(f"/me/events/{event_id}",
                             params={"$select": _SELECT + ",body,attendees,onlineMeeting"})
        out = shape_event(e)
        out["body"] = (e.get("body") or {}).get("content")
        out["attendees"] = [a["emailAddress"]["address"] for a in e.get("attendees", [])]
        join = (e.get("onlineMeeting") or {}).get("joinUrl")
        if join:
            out["join_url"] = join
        return out

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create a calendar event",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def create_event(subject: str, start: str, end: str, time_zone: str = "UTC",
                           attendees: list[str] | None = None, location: str | None = None,
                           body: str | None = None, is_online_meeting: bool = False) -> dict:
        """Create a calendar event. Times are ISO 8601 in the given time_zone."""
        payload = EventInput(subject=subject, start=start, end=end, time_zone=time_zone,
                             attendees=attendees or [], location=location, body=body,
                             is_online_meeting=is_online_meeting).to_graph()
        e = await client.post("/me/events", json=payload)
        return shape_event(e)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Update a calendar event",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=True,
        )
    )
    async def update_event(event_id: str, subject: str | None = None,
                           start: str | None = None, end: str | None = None,
                           time_zone: str = "UTC", location: str | None = None) -> dict:
        """Update fields on an existing event (only provided fields change)."""
        patch: dict = {}
        if subject is not None:
            patch["subject"] = subject
        if start is not None:
            patch["start"] = {"dateTime": start, "timeZone": time_zone}
        if end is not None:
            patch["end"] = {"dateTime": end, "timeZone": time_zone}
        if location is not None:
            patch["location"] = {"displayName": location}
        e = await client.patch(f"/me/events/{event_id}", json=patch)
        return shape_event(e)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete a calendar event",
            readOnlyHint=False,
            destructiveHint=True,
            openWorldHint=True,
        )
    )
    async def delete_event(event_id: str) -> dict:
        """Delete/cancel a calendar event."""
        await client.delete(f"/me/events/{event_id}")
        return {"deleted": event_id}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Respond to a meeting invite",
            readOnlyHint=False,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def respond_event(event_id: str, response: str, comment: str = "",
                            send_response: bool = True) -> dict:
        """Respond to a meeting invite. response is one of: accept, decline, tentative."""
        action = _RESPONSE_ACTIONS.get(response)
        if not action:
            raise ValueError("response must be accept, decline, or tentative")
        await client.post(f"/me/events/{event_id}/{action}",
                          json={"comment": comment, "sendResponse": send_response})
        return {"event_id": event_id, "response": response}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Find free/busy availability",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def find_availability(emails: list[str], start: str, end: str,
                                time_zone: str = "UTC", interval_minutes: int = 30) -> list[dict]:
        """Free/busy across the given people for a time window (Graph getSchedule)."""
        payload = {"schedules": emails,
                   "startTime": {"dateTime": start, "timeZone": time_zone},
                   "endTime": {"dateTime": end, "timeZone": time_zone},
                   "availabilityViewInterval": interval_minutes}
        data = await client.post("/me/calendar/getSchedule", json=payload)
        return [{"email": s.get("scheduleId"), "availability_view": s.get("availabilityView")}
                for s in data.get("value", [])]
