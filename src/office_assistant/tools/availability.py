"""MCP tools for checking availability and finding meeting times."""

from __future__ import annotations

from typing import Any

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

# Microsoft Graph encodes each availability slot as a single character.
_AVAILABILITY_CODES: dict[str, str] = {
    "0": "free",
    "1": "tentative",
    "2": "busy",
    "3": "out_of_office",
    "4": "working_elsewhere",
}


@mcp.tool()
async def get_free_busy(
    emails: list[str],
    start_datetime: str,
    end_datetime: str,
    start_timezone: str,
    ctx: Context,
    availability_view_interval: int = 30,
) -> dict[str, Any]:
    """Get free/busy schedule for one or more people.

    Returns each person's availability status across the time range,
    broken into slots of the specified interval.

    Args:
        emails: List of email addresses to check availability for.
        start_datetime: Start of the range in ISO 8601 format
            (e.g. "2026-02-17T09:00:00").
        end_datetime: End of the range in ISO 8601 format
            (e.g. "2026-02-17T17:00:00").
        start_timezone: IANA timezone (e.g. "Europe/London").
        availability_view_interval: Size of each time slot in minutes
            (default: 30).
    """
    if not emails:
        return {"error": "At least one email address is required."}
    if err := validate_emails(emails):
        return {"error": err}
    if not 5 <= availability_view_interval <= 1440 or availability_view_interval % 5 != 0:
        return {
            "error": (
                "availability_view_interval must be between 5 and 1440 minutes "
                "in 5-minute increments."
            )
        }
    if err := validate_timezone(start_timezone, "start_timezone"):
        return {"error": err}
    if err := validate_datetime_order(
        start_datetime,
        end_datetime,
        start_timezone=start_timezone,
        end_timezone=start_timezone,
    ):
        return {"error": err}

    graph = get_graph(ctx)

    body = {
        "schedules": emails,
        "startTime": {"dateTime": start_datetime, "timeZone": start_timezone},
        "endTime": {"dateTime": end_datetime, "timeZone": start_timezone},
        "availabilityViewInterval": availability_view_interval,
    }

    try:
        data = await graph.post("/me/calendar/getSchedule", json=body)
    except AuthenticationRequired as exc:
        return auth_required_response(exc)
    except GraphApiError as exc:
        return graph_error_response(exc)

    results = []
    for schedule in data.get("value", []):
        items = [
            {
                "status": item.get("status"),
                "subject": item.get("subject"),
                "start": item.get("start", {}).get("dateTime"),
                "end": item.get("end", {}).get("dateTime"),
            }
            for item in schedule.get("scheduleItems", [])
        ]

        # availabilityView is a string like "0022200000" where each char
        # maps to a slot status via _AVAILABILITY_CODES.
        availability_view = schedule.get("availabilityView", "")
        slot_summary = [_AVAILABILITY_CODES.get(ch, "unknown") for ch in availability_view]

        results.append(
            {
                "email": schedule.get("scheduleId"),
                "scheduleItems": items,
                "availabilitySlots": slot_summary,
                "workingHours": schedule.get("workingHours"),
            }
        )

    return {"schedules": results, "count": len(results)}


@mcp.tool()
async def find_meeting_times(
    attendees: list[str],
    duration_minutes: int,
    ctx: Context,
    start_datetime: str | None = None,
    end_datetime: str | None = None,
    start_timezone: str | None = None,
    max_candidates: int = 5,
) -> dict[str, Any]:
    """Find available meeting times that work for a group of people.

    Microsoft Graph checks each attendee's calendar and suggests times
    where everyone is available.  Returns suggestions sorted by
    confidence.

    Args:
        attendees: List of email addresses of required attendees.
        duration_minutes: How long the meeting should be in minutes
            (e.g. 30, 60).
        start_datetime: Start of the search window in ISO 8601
            (default: now).
        end_datetime: End of the search window in ISO 8601
            (default: 5 business days from now).
        start_timezone: IANA timezone (e.g. "Europe/London").
            Defaults to UTC.
        max_candidates: Maximum number of suggestions to return
            (default: 5).
    """
    if not attendees:
        return {"error": "At least one attendee is required."}
    if err := validate_emails(attendees):
        return {"error": err}
    if duration_minutes <= 0:
        return {"error": "duration_minutes must be greater than 0."}
    if not 1 <= max_candidates <= 20:
        return {"error": "max_candidates must be between 1 and 20."}

    graph = get_graph(ctx)

    body: dict[str, Any] = {
        "attendees": [
            {"emailAddress": {"address": email}, "type": "required"} for email in attendees
        ],
        "meetingDuration": f"PT{duration_minutes}M",
        "maxCandidates": max_candidates,
        "returnSuggestionReasons": True,
    }

    if bool(start_datetime) != bool(end_datetime):
        return {"error": "Provide both start_datetime and end_datetime, or omit both."}
    if start_timezone and not start_datetime:
        return {
            "error": (
                "start_timezone can only be provided when start_datetime and end_datetime "
                "are also provided."
            )
        }

    if start_datetime and end_datetime:
        tz = start_timezone or "UTC"
        if err := validate_timezone(tz, "start_timezone"):
            return {"error": err}
        if err := validate_datetime_order(
            start_datetime,
            end_datetime,
            start_timezone=tz,
            end_timezone=tz,
        ):
            return {"error": err}
        body["timeConstraint"] = {
            "timeslots": [
                {
                    "start": {"dateTime": start_datetime, "timeZone": tz},
                    "end": {"dateTime": end_datetime, "timeZone": tz},
                }
            ]
        }

    try:
        data = await graph.post("/me/findMeetingTimes", json=body)
    except AuthenticationRequired as exc:
        return auth_required_response(exc)
    except GraphApiError as exc:
        return graph_error_response(exc)

    suggestions = []
    for suggestion in data.get("meetingTimeSuggestions", []):
        slot = suggestion.get("meetingTimeSlot", {})
        attendee_availability = [
            {
                "email": (att.get("attendee", {}).get("emailAddress", {}).get("address")),
                "availability": att.get("availability"),
            }
            for att in suggestion.get("attendeeAvailability", [])
        ]

        suggestions.append(
            {
                "start": slot.get("start", {}).get("dateTime"),
                "end": slot.get("end", {}).get("dateTime"),
                "timezone": slot.get("start", {}).get("timeZone"),
                "confidence": suggestion.get("confidence"),
                "attendeeAvailability": attendee_availability,
                "suggestionReason": suggestion.get("suggestionReason"),
            }
        )

    result: dict[str, Any] = {"suggestions": suggestions, "count": len(suggestions)}

    empty_reason = data.get("emptySuggestionsReason")
    if empty_reason:
        result["emptySuggestionsReason"] = empty_reason

    return result
