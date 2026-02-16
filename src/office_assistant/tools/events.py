"""MCP tools for calendar event CRUD operations."""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import quote

from mcp.server.fastmcp import Context

from office_assistant.app import mcp
from office_assistant.auth import AuthenticationRequired
from office_assistant.graph_client import GraphApiError
from office_assistant.tools._helpers import (
    auth_required_response,
    get_graph,
    graph_error_response,
    validate_datetime_order,
    validate_emails,
    validate_timezone,
)

# Fields to request from the Graph API when listing events.
_EVENT_FIELDS = (
    "id,subject,start,end,location,isOnlineMeeting,onlineMeetingUrl,"
    "organizer,attendees,bodyPreview,isCancelled,showAs,isAllDay,recurrence"
)

_VALID_DAYS = {
    "sunday",
    "monday",
    "tuesday",
    "wednesday",
    "thursday",
    "friday",
    "saturday",
}

_VALID_PATTERN_TYPES = {
    "daily",
    "weekly",
    "absoluteMonthly",
    "relativeMonthly",
    "absoluteYearly",
    "relativeYearly",
}

_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

_RESPONSE_ENDPOINTS = {
    "accept": "accept",
    "decline": "decline",
    "tentatively_accept": "tentativelyAccept",
}


def _format_event(event: dict[str, Any]) -> dict[str, Any]:
    """Extract the useful fields from a Graph API event object."""
    attendees = [
        {
            "name": att.get("emailAddress", {}).get("name"),
            "email": att.get("emailAddress", {}).get("address"),
            "type": att.get("type"),
            "status": att.get("status", {}).get("response"),
        }
        for att in event.get("attendees", [])
    ]

    result: dict[str, Any] = {
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

    recurrence = event.get("recurrence")
    if recurrence:
        result["recurrence"] = recurrence

    return result


def _is_access_denied(exc: GraphApiError) -> bool:
    return exc.status_code == 403 or (exc.code or "").lower() == "erroraccessdenied"


def _user_base(user_email: str | None) -> str:
    """Build the Graph API path prefix for the target user."""
    if user_email:
        return f"/users/{quote(user_email, safe='')}"
    return "/me"


def _delegate_error(exc: GraphApiError, user_email: str | None) -> dict[str, Any]:
    """Return a clear error when delegate access is denied."""
    if user_email and _is_access_denied(exc):
        return graph_error_response(
            exc,
            fallback_message=(
                f"You don't have permission to modify {user_email}'s calendar. "
                "They need to grant you delegate access in Outlook."
            ),
        )
    return graph_error_response(exc)


def _build_recurrence(
    pattern: str,
    interval: int | None,
    days_of_week: list[str] | None,
    end_date: str | None,
    count: int | None,
    start_datetime: str,
) -> dict[str, Any] | str:
    """Build and validate a Graph API recurrence object.

    Returns the recurrence dict on success, or an error string on failure.
    """
    if pattern not in _VALID_PATTERN_TYPES:
        return f"recurrence_pattern must be one of: {', '.join(sorted(_VALID_PATTERN_TYPES))}."

    effective_interval = interval if interval is not None else 1
    if effective_interval < 1:
        return "recurrence_interval must be at least 1."

    if days_of_week:
        lowered = [d.lower() for d in days_of_week]
        invalid = [d for d in lowered if d not in _VALID_DAYS]
        if invalid:
            return f"Invalid days of week: {', '.join(invalid)}."
    else:
        lowered = None

    if pattern == "weekly" and not lowered:
        return "recurrence_days_of_week is required for weekly patterns."

    pattern_obj: dict[str, Any] = {
        "type": pattern,
        "interval": effective_interval,
    }
    if lowered:
        pattern_obj["daysOfWeek"] = lowered

    # Determine range type from provided end conditions.
    start_date = start_datetime.split("T")[0]
    if end_date:
        if not _DATE_RE.match(end_date):
            return "recurrence_end_date must be in YYYY-MM-DD format."
        range_obj: dict[str, Any] = {
            "type": "endDate",
            "startDate": start_date,
            "endDate": end_date,
        }
    elif count is not None:
        if count < 1:
            return "recurrence_count must be at least 1."
        range_obj = {
            "type": "numbered",
            "startDate": start_date,
            "numberOfOccurrences": count,
        }
    else:
        range_obj = {"type": "noEnd", "startDate": start_date}

    return {"pattern": pattern_obj, "range": range_obj}


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
    if err := validate_datetime_order(start_datetime, end_datetime):
        return {"error": err}
    if user_email and (err := validate_emails([user_email])):
        return {"error": err}

    graph = get_graph(ctx)
    base = _user_base(user_email)
    params = {
        "startDateTime": start_datetime,
        "endDateTime": end_datetime,
        "$orderby": "start/dateTime",
        "$top": "50",
        "$select": _EVENT_FIELDS,
    }

    try:
        data = await graph.get_all(f"{base}/calendarview", params=params)
    except AuthenticationRequired as exc:
        return auth_required_response(exc)
    except GraphApiError as exc:
        if user_email and _is_access_denied(exc):
            return graph_error_response(
                exc,
                fallback_message=(
                    f"You don't have permission to view {user_email}'s calendar. "
                    "Ask them to share their calendar with you in Outlook."
                ),
            )
        return graph_error_response(exc)

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
    room_emails: list[str] | None = None,
    body: str | None = None,
    location: str | None = None,
    is_online_meeting: bool = True,
    user_email: str | None = None,
    recurrence_pattern: str | None = None,
    recurrence_interval: int | None = None,
    recurrence_days_of_week: list[str] | None = None,
    recurrence_end_date: str | None = None,
    recurrence_count: int | None = None,
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
        room_emails: List of room/resource email addresses to book.
            Use list_rooms to find available rooms (work/school only).
        body: HTML or plain-text body for the meeting invitation.
        location: Physical location name.
        is_online_meeting: Whether to create a Teams meeting link
            (default: True).
        user_email: Create the event on another user's calendar
            (requires delegate access, work/school accounts only).
        recurrence_pattern: Recurrence type: "daily", "weekly",
            "absoluteMonthly", "relativeMonthly", "absoluteYearly",
            or "relativeYearly". Omit for a one-off event.
        recurrence_interval: How often the pattern repeats (e.g. 1 = every
            week, 2 = every other week). Defaults to 1.
        recurrence_days_of_week: Days for weekly patterns, e.g.
            ["monday", "wednesday"]. Required for weekly recurrence.
        recurrence_end_date: Stop recurring after this date (YYYY-MM-DD).
            Provide this OR recurrence_count, or omit both for no end.
        recurrence_count: Stop after this many occurrences.
            Provide this OR recurrence_end_date, or omit both for no end.
    """
    all_emails = list(attendees or [])
    if room_emails and (err := validate_emails(room_emails)):
        return {"error": err}
    if all_emails and (err := validate_emails(all_emails)):
        return {"error": err}
    if user_email and (err := validate_emails([user_email])):
        return {"error": err}
    if err := validate_timezone(start_timezone, "start_timezone"):
        return {"error": err}
    if err := validate_timezone(end_timezone, "end_timezone"):
        return {"error": err}
    if err := validate_datetime_order(
        start_datetime,
        end_datetime,
        start_timezone=start_timezone,
        end_timezone=end_timezone,
    ):
        return {"error": err}

    graph = get_graph(ctx)
    base = _user_base(user_email)

    event_body: dict[str, Any] = {
        "subject": subject,
        "start": {"dateTime": start_datetime, "timeZone": start_timezone},
        "end": {"dateTime": end_datetime, "timeZone": end_timezone},
    }

    # Online meetings (Teams) are only supported for work/school accounts.
    # Personal accounts silently ignore isOnlineMeeting.
    event_body["isOnlineMeeting"] = is_online_meeting
    if is_online_meeting:
        event_body["onlineMeetingProvider"] = "teamsForBusiness"

    # Build attendees list: people as "required", rooms as "resource".
    attendee_list: list[dict[str, Any]] = []
    if attendees:
        attendee_list.extend(
            {"emailAddress": {"address": email}, "type": "required"} for email in attendees
        )
    if room_emails:
        attendee_list.extend(
            {"emailAddress": {"address": email}, "type": "resource"} for email in room_emails
        )
    if attendee_list:
        event_body["attendees"] = attendee_list

    if body:
        event_body["body"] = {"contentType": "text", "content": body}

    if location:
        event_body["location"] = {"displayName": location}

    # Recurrence
    if recurrence_pattern:
        recurrence = _build_recurrence(
            recurrence_pattern,
            recurrence_interval,
            recurrence_days_of_week,
            recurrence_end_date,
            recurrence_count,
            start_datetime,
        )
        if isinstance(recurrence, str):
            return {"error": recurrence}
        event_body["recurrence"] = recurrence
    elif any(
        v is not None
        for v in [
            recurrence_interval,
            recurrence_days_of_week,
            recurrence_end_date,
            recurrence_count,
        ]
    ):
        return {
            "error": "recurrence_pattern is required when any recurrence parameter is provided."
        }

    try:
        data = await graph.post(f"{base}/calendar/events", json=event_body)
    except AuthenticationRequired as exc:
        return auth_required_response(exc)
    except GraphApiError as exc:
        return _delegate_error(exc, user_email)
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
    user_email: str | None = None,
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
        user_email: Update an event on another user's calendar
            (requires delegate access, work/school accounts only).
    """
    if attendees is not None and (err := validate_emails(attendees)):
        return {"error": err}
    if user_email and (err := validate_emails([user_email])):
        return {"error": err}
    if start_timezone is not None and (err := validate_timezone(start_timezone, "start_timezone")):
        return {"error": err}
    if end_timezone is not None and (err := validate_timezone(end_timezone, "end_timezone")):
        return {"error": err}

    graph = get_graph(ctx)
    base = _user_base(user_email)

    # If the caller provides a datetime without a timezone (or vice versa),
    # or updates only one side of the time window, fetch the existing event
    # so we can fill in the missing pieces.
    start_touched = start_datetime is not None or start_timezone is not None
    end_touched = end_datetime is not None or end_timezone is not None
    need_existing = (
        (start_datetime is not None) != (start_timezone is not None)
        or (end_datetime is not None) != (end_timezone is not None)
        or (start_touched != end_touched)
    )
    existing: dict[str, Any] = {}
    if need_existing:
        try:
            existing = await graph.get(
                f"{base}/events/{event_id}", params={"$select": "start,end"}
            )
        except AuthenticationRequired as exc:
            return auth_required_response(exc)
        except GraphApiError as exc:
            return _delegate_error(exc, user_email)

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

    candidate_start = updates.get("start", existing.get("start"))
    candidate_end = updates.get("end", existing.get("end"))
    if (
        candidate_start
        and candidate_end
        and (
            err := validate_datetime_order(
                candidate_start.get("dateTime", ""),
                candidate_end.get("dateTime", ""),
            )
        )
    ):
        return {"error": err}

    try:
        data = await graph.patch(f"{base}/events/{event_id}", json=updates)
    except AuthenticationRequired as exc:
        return auth_required_response(exc)
    except GraphApiError as exc:
        return _delegate_error(exc, user_email)
    return _format_event(data)


@mcp.tool()
async def cancel_event(
    event_id: str,
    ctx: Context,
    comment: str | None = None,
    user_email: str | None = None,
) -> dict[str, Any]:
    """Cancel (delete) a calendar event.

    If you are the organiser, this sends a cancellation notice to all
    attendees.  If you are an attendee, this declines and removes the
    event from your calendar.

    Args:
        event_id: The ID of the event to cancel (from list_events).
        comment: Optional message to include in the cancellation notice.
        user_email: Cancel an event on another user's calendar
            (requires delegate access, work/school accounts only).
    """
    if user_email and (err := validate_emails([user_email])):
        return {"error": err}

    graph = get_graph(ctx)
    base = _user_base(user_email)

    if comment:
        # The /cancel endpoint sends a cancellation message to attendees.
        try:
            await graph.post(f"{base}/events/{event_id}/cancel", json={"comment": comment})
        except AuthenticationRequired as exc:
            return auth_required_response(exc)
        except GraphApiError as exc:
            return _delegate_error(exc, user_email)
    else:
        try:
            await graph.delete(f"{base}/events/{event_id}")
        except AuthenticationRequired as exc:
            return auth_required_response(exc)
        except GraphApiError as exc:
            return _delegate_error(exc, user_email)

    return {"status": "cancelled", "event_id": event_id}


@mcp.tool()
async def respond_to_event(
    event_id: str,
    response: str,
    ctx: Context,
    comment: str | None = None,
    send_response: bool = True,
    user_email: str | None = None,
) -> dict[str, Any]:
    """Respond to a meeting invitation (accept, decline, or tentatively accept).

    Args:
        event_id: The ID of the event to respond to (from list_events).
        response: Your response: "accept", "decline", or
            "tentatively_accept".
        comment: Optional message to include with your response.
        send_response: Whether to notify the organiser (default: True).
        user_email: Respond on behalf of another user whose calendar
            you have delegate access to (work/school accounts only).
    """
    if user_email and (err := validate_emails([user_email])):
        return {"error": err}

    endpoint = _RESPONSE_ENDPOINTS.get(response)
    if not endpoint:
        valid = ", ".join(sorted(_RESPONSE_ENDPOINTS))
        return {"error": f"response must be one of: {valid}."}

    graph = get_graph(ctx)
    base = _user_base(user_email)

    body: dict[str, Any] = {
        "comment": comment or "",
        "sendResponse": send_response,
    }

    try:
        await graph.post(f"{base}/events/{event_id}/{endpoint}", json=body)
    except AuthenticationRequired as exc:
        return auth_required_response(exc)
    except GraphApiError as exc:
        if user_email and _is_access_denied(exc):
            return graph_error_response(
                exc,
                fallback_message=(
                    f"You don't have permission to respond to events on {user_email}'s calendar."
                ),
            )
        return graph_error_response(exc)

    return {"status": "responded", "event_id": event_id, "response": response}
