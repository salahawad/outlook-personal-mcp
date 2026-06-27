from __future__ import annotations

from typing import Optional

from pydantic import BaseModel, Field

Recipient = str  # plain email address


def _recipients(addrs: list[str]) -> list[dict]:
    return [{"emailAddress": {"address": a}} for a in addrs]


class SendMailInput(BaseModel):
    to: list[str]
    subject: str
    body: str
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    html: bool = False
    save_to_sent: bool = True

    def to_graph(self) -> dict:
        msg = {
            "subject": self.subject,
            "body": {"contentType": "HTML" if self.html else "Text", "content": self.body},
            "toRecipients": _recipients(self.to),
        }
        if self.cc:
            msg["ccRecipients"] = _recipients(self.cc)
        if self.bcc:
            msg["bccRecipients"] = _recipients(self.bcc)
        return {"message": msg, "saveToSentItems": self.save_to_sent}


class EventInput(BaseModel):
    subject: str
    start: str  # ISO 8601, e.g. 2026-06-27T09:00:00
    end: str
    time_zone: str = "UTC"
    attendees: list[str] = Field(default_factory=list)
    location: Optional[str] = None
    body: Optional[str] = None
    is_online_meeting: bool = False

    def to_graph(self) -> dict:
        ev: dict = {
            "subject": self.subject,
            "start": {"dateTime": self.start, "timeZone": self.time_zone},
            "end": {"dateTime": self.end, "timeZone": self.time_zone},
        }
        if self.attendees:
            ev["attendees"] = [
                {"emailAddress": {"address": a}, "type": "required"} for a in self.attendees
            ]
        if self.location:
            ev["location"] = {"displayName": self.location}
        if self.body:
            ev["body"] = {"contentType": "HTML", "content": self.body}
        if self.is_online_meeting:
            ev["isOnlineMeeting"] = True
            ev["onlineMeetingProvider"] = "teamsForBusiness"
        return ev


def shape_message(m: dict) -> dict:
    return {
        "id": m.get("id"),
        "subject": m.get("subject"),
        "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
        "is_read": m.get("isRead"),
        "received": m.get("receivedDateTime"),
        "preview": m.get("bodyPreview"),
        "has_attachments": m.get("hasAttachments"),
    }


def shape_event(e: dict) -> dict:
    def _dt(d: dict | None) -> str:
        d = d or {}
        return f"{d.get('dateTime')} {d.get('timeZone')}".strip()

    return {
        "id": e.get("id"),
        "subject": e.get("subject"),
        "start": _dt(e.get("start")),
        "end": _dt(e.get("end")),
        "location": (e.get("location") or {}).get("displayName"),
        "is_all_day": e.get("isAllDay"),
        "organizer": (e.get("organizer") or {}).get("emailAddress", {}).get("address"),
    }
