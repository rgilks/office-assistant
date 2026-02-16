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
from office_assistant.tools.events import (
    cancel_event,
    create_event,
    list_events,
    respond_to_event,
    update_event,
)
from office_assistant.tools.rooms import list_rooms


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
                },
            )
        )
        router.get("/v1.0/me/mailboxSettings").mock(
            return_value=Response(
                200,
                json={"timeZone": "Europe/London"},
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


@pytest.mark.asyncio
async def test_create_recurring_event_flow(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        create_route = router.post("/v1.0/me/calendar/events").mock(
            return_value=Response(
                201,
                json={
                    "id": "recurring-1",
                    "subject": "Team Sync",
                    "start": {"dateTime": "2026-03-03T10:00:00", "timeZone": "UTC"},
                    "end": {"dateTime": "2026-03-03T11:00:00", "timeZone": "UTC"},
                    "attendees": [],
                    "recurrence": {
                        "pattern": {
                            "type": "weekly",
                            "interval": 1,
                            "daysOfWeek": ["tuesday"],
                        },
                        "range": {
                            "type": "numbered",
                            "startDate": "2026-03-03",
                            "numberOfOccurrences": 10,
                        },
                    },
                },
            )
        )

        result = await create_event(
            subject="Team Sync",
            start_datetime="2026-03-03T10:00:00",
            start_timezone="UTC",
            end_datetime="2026-03-03T11:00:00",
            end_timezone="UTC",
            ctx=integration_ctx,
            recurrence_pattern="weekly",
            recurrence_days_of_week=["tuesday"],
            recurrence_count=10,
        )

    payload = json.loads(create_route.calls[0].request.content.decode())
    assert payload["recurrence"]["pattern"]["type"] == "weekly"
    assert payload["recurrence"]["range"]["numberOfOccurrences"] == 10
    assert result["recurrence"]["pattern"]["daysOfWeek"] == ["tuesday"]


@pytest.mark.asyncio
async def test_delegate_create_event_flow(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        create_route = router.post("/v1.0/users/boss%40company.com/calendar/events").mock(
            return_value=Response(
                201,
                json={
                    "id": "delegate-1",
                    "subject": "Board Meeting",
                    "start": {"dateTime": "2026-03-03T14:00:00", "timeZone": "UTC"},
                    "end": {"dateTime": "2026-03-03T15:00:00", "timeZone": "UTC"},
                    "attendees": [],
                },
            )
        )

        result = await create_event(
            subject="Board Meeting",
            start_datetime="2026-03-03T14:00:00",
            start_timezone="UTC",
            end_datetime="2026-03-03T15:00:00",
            end_timezone="UTC",
            ctx=integration_ctx,
            user_email="boss@company.com",
        )

    assert create_route.called
    assert result["id"] == "delegate-1"


@pytest.mark.asyncio
async def test_respond_to_invitation_flow(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        accept_route = router.post("/v1.0/me/events/event-1/accept").mock(
            return_value=Response(202)
        )

        result = await respond_to_event(
            event_id="event-1",
            response="accept",
            ctx=integration_ctx,
            comment="Looking forward to it!",
        )

    assert accept_route.called
    assert result["status"] == "responded"
    payload = json.loads(accept_route.calls[0].request.content.decode())
    assert payload["comment"] == "Looking forward to it!"


@pytest.mark.asyncio
async def test_list_rooms_flow(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        router.get("/v1.0/places/microsoft.graph.room").mock(
            return_value=Response(
                200,
                json={
                    "value": [
                        {
                            "displayName": "Boardroom",
                            "emailAddress": "boardroom@company.com",
                            "capacity": 12,
                            "building": "HQ",
                            "floorLabel": "Floor 5",
                        }
                    ]
                },
            )
        )

        result = await list_rooms(ctx=integration_ctx)

    assert result["count"] == 1
    assert result["rooms"][0]["email"] == "boardroom@company.com"


@pytest.mark.asyncio
async def test_delegate_update_event_flow(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        update_route = router.patch("/v1.0/users/boss%40company.com/events/event-1").mock(
            return_value=Response(
                200,
                json={
                    "id": "event-1",
                    "subject": "Updated by EA",
                    "start": {"dateTime": "2026-03-03T14:00:00", "timeZone": "UTC"},
                    "end": {"dateTime": "2026-03-03T15:00:00", "timeZone": "UTC"},
                    "attendees": [],
                },
            )
        )

        result = await update_event(
            event_id="event-1",
            ctx=integration_ctx,
            subject="Updated by EA",
            user_email="boss@company.com",
        )

    assert update_route.called
    assert result["subject"] == "Updated by EA"


@pytest.mark.asyncio
async def test_delegate_cancel_event_flow(integration_ctx):
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        cancel_route = router.delete("/v1.0/users/boss%40company.com/events/event-1").mock(
            return_value=Response(204)
        )

        result = await cancel_event(
            event_id="event-1",
            ctx=integration_ctx,
            user_email="boss@company.com",
        )

    assert cancel_route.called
    assert result["status"] == "cancelled"


@pytest.mark.asyncio
async def test_delegate_access_denied_flow(integration_ctx):
    """Integration test: 403 on delegate calendar returns a clear message."""
    with respx.mock(base_url="https://graph.microsoft.com") as router:
        router.post("/v1.0/users/boss%40company.com/calendar/events").mock(
            return_value=Response(
                403,
                json={
                    "error": {
                        "code": "ErrorAccessDenied",
                        "message": "Access is denied.",
                    }
                },
                headers={"request-id": "req-99"},
            )
        )

        result = await create_event(
            subject="Test",
            start_datetime="2026-03-03T14:00:00",
            start_timezone="UTC",
            end_datetime="2026-03-03T15:00:00",
            end_timezone="UTC",
            ctx=integration_ctx,
            user_email="boss@company.com",
        )

    assert result["errorType"] == "permission_denied"
    assert "delegate access" in result["error"]
