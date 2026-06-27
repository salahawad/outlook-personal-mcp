from outlook_personal_mcp.models import (Recipient, SendMailInput, shape_message, shape_event)


def test_recipient_importable():
    assert Recipient is str


def test_send_mail_input_normalizes_recipients():
    m = SendMailInput(to=["a@x.com"], subject="hi", body="hello")
    payload = m.to_graph()
    assert payload["message"]["toRecipients"] == [{"emailAddress": {"address": "a@x.com"}}]
    assert payload["message"]["body"]["contentType"] == "Text"
    assert payload["message"]["subject"] == "hi"


def test_shape_message_trims_fields():
    raw = {"id": "1", "subject": "S", "isRead": False,
           "from": {"emailAddress": {"address": "a@x.com", "name": "A"}},
           "receivedDateTime": "2026-06-27T10:00:00Z", "bodyPreview": "p",
           "hasAttachments": True, "extra": "drop me"}
    s = shape_message(raw)
    assert s == {"id": "1", "subject": "S", "from": "a@x.com", "is_read": False,
                 "received": "2026-06-27T10:00:00Z", "preview": "p", "has_attachments": True}


def test_shape_event_trims_fields():
    raw = {"id": "e1", "subject": "Sync", "start": {"dateTime": "2026-06-27T09:00:00", "timeZone": "UTC"},
           "end": {"dateTime": "2026-06-27T09:30:00", "timeZone": "UTC"},
           "location": {"displayName": "Room"}, "isAllDay": False, "junk": 1}
    s = shape_event(raw)
    assert s["id"] == "e1" and s["subject"] == "Sync"
    assert s["start"] == "2026-06-27T09:00:00 UTC" and s["location"] == "Room"
