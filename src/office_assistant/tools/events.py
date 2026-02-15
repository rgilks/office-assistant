"""MCP tools for calendar event CRUD operations."""

from __future__ import annotations

from typing import Any
from urllib.parse import quote

from mcp.server.fastmcp import Context

from office_assistant.app import mcp
from office_assistant.tools._helpers import get_graph, validate_emails

# Fields to request from the Graph API when listing events.
_EVENT_FIELDS = (
    "id,subject,start,end,location,isOnlineMeeting,onlineMeetingUrl,"
    "organizer,attendees,bodyPreview,isCancelled,showAs,isAllDay"
)


def _format_event(event: dict[str, Any]) -> dict[str, Any]:
    """Extract the useful fields from a Graph API event object."""
    attendees = [
        {
            "name": att.get("emailAddress", {}).get("name"),
            "email": att.get("emailAddress", {}).get("address"),
            "status": att.get("status", {}).get("response"),
        }
        for att in event.get("attendees", [])
    ]

    return {
        "id": event.get("id"),
        "subject": event.get("subject"),
        "start": event.get("start", {}).get("dateTime"),
        "startTimezone": event.get("start", {}).get("timeZone"),
        "end": event.get("end", {}).get("dateTime"),
        "endTimezone": event.get("end", {}).get("timeZone"),
        "location": event.get("location", {}).get("displayName"),
        "isOnlineMeeting": event.get("isOnlineMeeting", False),
        "onlineMeetingUrl": event.get("onlineMeetingUrl"),
        "organizer": event.get("organizer", {}).get("emailAddress", {}).get("name"),
        "organizerEmail": (event.get("organizer", {}).get("emailAddress", {}).get("address")),
        "attendees": attendees,
        "bodyPreview": event.get("bodyPreview"),
        "isCancelled": event.get("isCancelled", False),
        "showAs": event.get("showAs"),
        "isAllDay": event.get("isAllDay", False),
    }


@mcp.tool()
async def list_events(
    start_datetime: str,
    end_datetime: str,
    ctx: Context,
    user_email: str | None = None,
) -> dict[str, Any]:
    """List calendar events for a date range.

    Args:
        start_datetime: Start of the range in ISO 8601 format
            (e.g. "2026-02-16T00:00:00").
        end_datetime: End of the range in ISO 8601 format
            (e.g. "2026-02-16T23:59:59").
        user_email: Email address of another user whose calendar to view.
            Omit to view the authenticated user's calendar.
            Requires that the user has shared their calendar with you.
    """
    if user_email and (err := validate_emails([user_email])):
        return {"error": err}

    graph = get_graph(ctx)
    base = f"/users/{quote(user_email, safe='')}" if user_email else "/me"
    params = {
        "startDateTime": start_datetime,
        "endDateTime": end_datetime,
        "$orderby": "start/dateTime",
        "$top": "50",
        "$select": _EVENT_FIELDS,
    }

    try:
        data = await graph.get(f"{base}/calendarview", params=params)
    except Exception as exc:
        error_text = str(exc)
        if "ErrorAccessDenied" in error_text or "403" in error_text:
            return {
                "error": (
                    f"You don't have permission to view {user_email}'s calendar. "
                    "Ask them to share their calendar with you in Outlook."
                )
            }
        raise

    events = [_format_event(ev) for ev in data.get("value", [])]
    return {"events": events, "count": len(events)}


@mcp.tool()
async def create_event(
    subject: str,
    start_datetime: str,
    start_timezone: str,
    end_datetime: str,
    end_timezone: str,
    ctx: Context,
    attendees: list[str] | None = None,
    body: str | None = None,
    location: str | None = None,
    is_online_meeting: bool = True,
) -> dict[str, Any]:
    """Create a new calendar event.

    Args:
        subject: The meeting title / subject.
        start_datetime: Start time in ISO 8601 format
            (e.g. "2026-02-17T09:00:00").
        start_timezone: IANA timezone for the start time
            (e.g. "Europe/London").
        end_datetime: End time in ISO 8601 format
            (e.g. "2026-02-17T10:00:00").
        end_timezone: IANA timezone for the end time
            (e.g. "Europe/London").
        attendees: List of email addresses to invite.
        body: HTML or plain-text body for the meeting invitation.
        location: Physical location name.
        is_online_meeting: Whether to create a Teams meeting link
            (default: True).
    """
    if attendees and (err := validate_emails(attendees)):
        return {"error": err}

    graph = get_graph(ctx)

    event_body: dict[str, Any] = {
        "subject": subject,
        "start": {"dateTime": start_datetime, "timeZone": start_timezone},
        "end": {"dateTime": end_datetime, "timeZone": end_timezone},
        "isOnlineMeeting": is_online_meeting,
    }

    if is_online_meeting:
        event_body["onlineMeetingProvider"] = "teamsForBusiness"

    if attendees:
        event_body["attendees"] = [
            {"emailAddress": {"address": email}, "type": "required"} for email in attendees
        ]

    if body:
        event_body["body"] = {"contentType": "text", "content": body}

    if location:
        event_body["location"] = {"displayName": location}

    data = await graph.post("/me/calendar/events", json=event_body)
    return _format_event(data)


@mcp.tool()
async def update_event(
    event_id: str,
    ctx: Context,
    subject: str | None = None,
    start_datetime: str | None = None,
    start_timezone: str | None = None,
    end_datetime: str | None = None,
    end_timezone: str | None = None,
    attendees: list[str] | None = None,
    body: str | None = None,
    location: str | None = None,
    is_online_meeting: bool | None = None,
) -> dict[str, Any]:
    """Update an existing calendar event.

    Only provide the fields you want to change; others remain unchanged.

    Args:
        event_id: The ID of the event to update (from list_events).
        subject: New meeting title.
        start_datetime: New start time in ISO 8601 format.
        start_timezone: IANA timezone for the new start time.
        end_datetime: New end time in ISO 8601 format.
        end_timezone: IANA timezone for the new end time.
        attendees: New complete list of attendee email addresses
            (replaces existing).
        body: New meeting body text.
        location: New location name.
        is_online_meeting: Whether it should be a Teams meeting.
    """
    if attendees is not None and (err := validate_emails(attendees)):
        return {"error": err}

    graph = get_graph(ctx)

    # If the caller provides a datetime without a timezone (or vice versa),
    # fetch the existing event so we can fill in the missing piece.
    need_existing = (start_datetime is not None) != (start_timezone is not None) or (
        end_datetime is not None
    ) != (end_timezone is not None)
    existing: dict[str, Any] = {}
    if need_existing:
        existing = await graph.get(f"/me/events/{event_id}", params={"$select": "start,end"})

    updates: dict[str, Any] = {}
    if subject is not None:
        updates["subject"] = subject
    if start_datetime is not None or start_timezone is not None:
        dt = start_datetime or existing.get("start", {}).get("dateTime")
        tz = start_timezone or existing.get("start", {}).get("timeZone")
        if dt and tz:
            updates["start"] = {"dateTime": dt, "timeZone": tz}
    if end_datetime is not None or end_timezone is not None:
        dt = end_datetime or existing.get("end", {}).get("dateTime")
        tz = end_timezone or existing.get("end", {}).get("timeZone")
        if dt and tz:
            updates["end"] = {"dateTime": dt, "timeZone": tz}
    if attendees is not None:
        updates["attendees"] = [
            {"emailAddress": {"address": email}, "type": "required"} for email in attendees
        ]
    if body is not None:
        updates["body"] = {"contentType": "text", "content": body}
    if location is not None:
        updates["location"] = {"displayName": location}
    if is_online_meeting is not None:
        updates["isOnlineMeeting"] = is_online_meeting
        if is_online_meeting:
            updates["onlineMeetingProvider"] = "teamsForBusiness"

    if not updates:
        return {"error": "No fields to update. Provide at least one field to change."}

    data = await graph.patch(f"/me/events/{event_id}", json=updates)
    return _format_event(data)


@mcp.tool()
async def cancel_event(
    event_id: str,
    ctx: Context,
    comment: str | None = None,
) -> dict[str, Any]:
    """Cancel (delete) a calendar event.

    If you are the organiser, this sends a cancellation notice to all
    attendees.  If you are an attendee, this declines and removes the
    event from your calendar.

    Args:
        event_id: The ID of the event to cancel (from list_events).
        comment: Optional message to include in the cancellation notice.
    """
    graph = get_graph(ctx)

    if comment:
        # The /cancel endpoint sends a cancellation message to attendees.
        await graph.post(f"/me/events/{event_id}/cancel", json={"comment": comment})
    else:
        await graph.delete(f"/me/events/{event_id}")

    return {"status": "cancelled", "event_id": event_id}
