"""Integration-style tests for core tool workflows.

These tests use the real GraphClient with mocked HTTP responses so each
tool exercises end-to-end request/response behavior.
"""

from __future__ import annotations

import json
from unittest.mock import MagicMock

import pytest
import respx
from httpx import Response

from office_assistant.graph_client import GraphClient
from office_assistant.tools.availability import find_meeting_times, get_free_busy
from office_assistant.tools.calendars import get_my_profile, list_calendars
from office_assistant.tools.events import cancel_event, create_event, list_events, update_event


@pytest.fixture
async def integration_ctx(monkeypatch):
    monkeypatch.setattr("office_assistant.graph_client.get_token", lambda: "fake-token")
    client = GraphClient()
    ctx = MagicMock()
    ctx.request_context.lifespan_context.graph = client
    try:
        yield ctx
    finally:
        await client.close()


@pytest.mark.asyncio
async def test_profile_and_calendar_list_flow(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        router.get("/v1.0/me").mock(
            return_value=Response(
                200,
                json={
                    "displayName": "Alice Smith",
                    "mail": "alice@company.com",
                    "userPrincipalName": "alice@company.com",
                    "mailboxSettings": {"timeZone": "Europe/London"},
                },
            )
        )
        router.get("/v1.0/me/calendars").mock(
            return_value=Response(
                200,
                json={
                    "value": [
                        {
                            "id": "cal-1",
                            "name": "Calendar",
                            "owner": {"name": "Alice", "address": "alice@company.com"},
                            "canEdit": True,
                            "isDefaultCalendar": True,
                        }
                    ]
                },
            )
        )

        profile = await get_my_profile(integration_ctx)
        calendars = await list_calendars(integration_ctx)

    assert profile["email"] == "alice@company.com"
    assert profile["timezone"] == "Europe/London"
    assert calendars["count"] == 1
    assert calendars["calendars"][0]["id"] == "cal-1"


@pytest.mark.asyncio
async def test_event_create_update_cancel_flow(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        create_route = router.post("/v1.0/me/calendar/events").mock(
            return_value=Response(
                201,
                json={
                    "id": "event-1",
                    "subject": "Planning",
                    "start": {"dateTime": "2026-03-02T09:00:00", "timeZone": "Europe/London"},
                    "end": {"dateTime": "2026-03-02T09:30:00", "timeZone": "Europe/London"},
                    "attendees": [],
                },
            )
        )
        update_route = router.patch("/v1.0/me/events/event-1").mock(
            return_value=Response(
                200,
                json={
                    "id": "event-1",
                    "subject": "Planning - Updated",
                    "start": {"dateTime": "2026-03-02T10:00:00", "timeZone": "Europe/London"},
                    "end": {"dateTime": "2026-03-02T10:30:00", "timeZone": "Europe/London"},
                    "attendees": [],
                },
            )
        )
        cancel_route = router.delete("/v1.0/me/events/event-1").mock(return_value=Response(204))

        created = await create_event(
            subject="Planning",
            start_datetime="2026-03-02T09:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-03-02T09:30:00",
            end_timezone="Europe/London",
            ctx=integration_ctx,
        )
        updated = await update_event(
            event_id="event-1",
            ctx=integration_ctx,
            subject="Planning - Updated",
            start_datetime="2026-03-02T10:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-03-02T10:30:00",
            end_timezone="Europe/London",
        )
        cancelled = await cancel_event(event_id="event-1", ctx=integration_ctx)

    create_payload = json.loads(create_route.calls[0].request.content.decode())
    update_payload = json.loads(update_route.calls[0].request.content.decode())
    assert create_payload["subject"] == "Planning"
    assert update_payload["subject"] == "Planning - Updated"
    assert cancel_route.called
    assert created["id"] == "event-1"
    assert updated["subject"] == "Planning - Updated"
    assert cancelled["status"] == "cancelled"


@pytest.mark.asyncio
async def test_availability_and_meeting_suggestions_flow(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        router.post("/v1.0/me/calendar/getSchedule").mock(
            return_value=Response(
                200,
                json={
                    "value": [
                        {
                            "scheduleId": "alice@company.com",
                            "availabilityView": "0022",
                            "scheduleItems": [
                                {
                                    "status": "busy",
                                    "subject": "Standup",
                                    "start": {"dateTime": "2026-03-02T09:00:00"},
                                    "end": {"dateTime": "2026-03-02T09:30:00"},
                                }
                            ],
                        }
                    ]
                },
            )
        )
        router.post("/v1.0/me/findMeetingTimes").mock(
            return_value=Response(
                200,
                json={
                    "meetingTimeSuggestions": [
                        {
                            "meetingTimeSlot": {
                                "start": {
                                    "dateTime": "2026-03-03T11:00:00",
                                    "timeZone": "Europe/London",
                                },
                                "end": {
                                    "dateTime": "2026-03-03T11:30:00",
                                    "timeZone": "Europe/London",
                                },
                            },
                            "confidence": 95.0,
                            "attendeeAvailability": [
                                {
                                    "attendee": {"emailAddress": {"address": "alice@company.com"}},
                                    "availability": "free",
                                }
                            ],
                            "suggestionReason": "All attendees are available.",
                        }
                    ]
                },
            )
        )

        free_busy = await get_free_busy(
            emails=["alice@company.com"],
            start_datetime="2026-03-02T09:00:00",
            end_datetime="2026-03-02T12:00:00",
            start_timezone="Europe/London",
            ctx=integration_ctx,
        )
        suggestions = await find_meeting_times(
            attendees=["alice@company.com"],
            duration_minutes=30,
            start_datetime="2026-03-03T09:00:00",
            end_datetime="2026-03-04T17:00:00",
            start_timezone="Europe/London",
            ctx=integration_ctx,
        )

    assert free_busy["count"] == 1
    assert free_busy["schedules"][0]["availabilitySlots"][0] == "free"
    assert suggestions["count"] == 1
    assert suggestions["suggestions"][0]["confidence"] == 95.0


@pytest.mark.asyncio
async def test_permission_denied_is_mapped_in_list_events(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        router.get("/v1.0/users/secret%40company.com/calendarview").mock(
            return_value=Response(
                403,
                json={
                    "error": {
                        "code": "ErrorAccessDenied",
                        "message": "Access is denied.",
                    }
                },
                headers={"request-id": "req-42"},
            )
        )

        result = await list_events(
            start_datetime="2026-03-02T00:00:00",
            end_datetime="2026-03-02T23:59:59",
            user_email="secret@company.com",
            ctx=integration_ctx,
        )

    assert result["errorType"] == "permission_denied"
    assert result["statusCode"] == 403
    assert result["requestId"] == "req-42"
    assert "share their calendar" in result["error"]
