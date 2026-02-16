"""Tests for event CRUD tools."""

from __future__ import annotations

import pytest

from office_assistant.graph_client import GraphApiError
from office_assistant.tools.events import (
    cancel_event,
    create_event,
    list_events,
    respond_to_event,
    update_event,
)

SAMPLE_EVENT = {
    "id": "event-1",
    "subject": "Team Standup",
    "start": {"dateTime": "2026-02-16T09:00:00", "timeZone": "Europe/London"},
    "end": {"dateTime": "2026-02-16T09:30:00", "timeZone": "Europe/London"},
    "location": {"displayName": "Room 42"},
    "isOnlineMeeting": True,
    "onlineMeetingUrl": "https://teams.microsoft.com/meet/123",
    "organizer": {"emailAddress": {"name": "Alice", "address": "alice@company.com"}},
    "attendees": [
        {
            "emailAddress": {"name": "Bob", "address": "bob@company.com"},
            "type": "required",
            "status": {"response": "accepted"},
        }
    ],
    "bodyPreview": "Daily standup",
    "isCancelled": False,
    "showAs": "busy",
    "isAllDay": False,
}


class TestListEvents:
    @pytest.mark.asyncio
    async def test_list_own_events(self, mock_ctx, mock_graph):
        mock_graph.get_all.return_value = {"value": [SAMPLE_EVENT]}

        result = await list_events(
            start_datetime="2026-02-16T00:00:00",
            end_datetime="2026-02-16T23:59:59",
            ctx=mock_ctx,
        )

        assert result["count"] == 1
        assert result["events"][0]["subject"] == "Team Standup"
        assert result["events"][0]["location"] == "Room 42"
        # Should use /me/calendarview
        call_args = mock_graph.get_all.call_args
        assert "/me/calendarview" in call_args[0][0]

    @pytest.mark.asyncio
    async def test_list_other_user_events(self, mock_ctx, mock_graph):
        mock_graph.get_all.return_value = {"value": []}

        await list_events(
            start_datetime="2026-02-16T00:00:00",
            end_datetime="2026-02-16T23:59:59",
            ctx=mock_ctx,
            user_email="bob@company.com",
        )

        call_args = mock_graph.get_all.call_args
        assert "/users/bob%40company.com/calendarview" in call_args[0][0]

    @pytest.mark.asyncio
    async def test_invalid_user_email_rejected(self, mock_ctx, mock_graph):
        result = await list_events(
            start_datetime="2026-02-16T00:00:00",
            end_datetime="2026-02-16T23:59:59",
            ctx=mock_ctx,
            user_email="not-an-email",
        )

        assert "error" in result
        assert "Invalid email" in result["error"]
        mock_graph.get_all.assert_not_called()

    @pytest.mark.asyncio
    async def test_access_denied(self, mock_ctx, mock_graph):
        mock_graph.get_all.side_effect = GraphApiError(
            status_code=403,
            code="ErrorAccessDenied",
            message="Forbidden",
        )

        result = await list_events(
            start_datetime="2026-02-16T00:00:00",
            end_datetime="2026-02-16T23:59:59",
            ctx=mock_ctx,
            user_email="secret@company.com",
        )

        assert "error" in result
        assert "permission" in result["error"].lower()
        assert result["errorType"] == "permission_denied"

    @pytest.mark.asyncio
    async def test_invalid_datetime_range_rejected(self, mock_ctx, mock_graph):
        result = await list_events(
            start_datetime="2026-02-17T00:00:00",
            end_datetime="2026-02-16T23:59:59",
            ctx=mock_ctx,
        )

        assert "error" in result
        assert "must be before" in result["error"]
        mock_graph.get_all.assert_not_called()


class TestCreateEvent:
    @pytest.mark.asyncio
    async def test_create_basic_event(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = SAMPLE_EVENT

        result = await create_event(
            subject="Team Standup",
            start_datetime="2026-02-16T09:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-16T09:30:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
        )

        assert result["subject"] == "Team Standup"
        call_args = mock_graph.post.call_args
        body = call_args[1]["json"]
        # Default is False (safe for personal accounts); work accounts
        # should explicitly pass is_online_meeting=True.
        assert body["isOnlineMeeting"] is False
        assert "onlineMeetingProvider" not in body

    @pytest.mark.asyncio
    async def test_create_with_attendees(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = SAMPLE_EVENT

        await create_event(
            subject="Review",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            attendees=["bob@company.com", "carol@company.com"],
        )

        body = mock_graph.post.call_args[1]["json"]
        assert len(body["attendees"]) == 2
        assert body["attendees"][0]["emailAddress"]["address"] == "bob@company.com"

    @pytest.mark.asyncio
    async def test_create_no_teams(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = SAMPLE_EVENT

        await create_event(
            subject="In-person meeting",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            is_online_meeting=False,
            location="Board Room",
        )

        body = mock_graph.post.call_args[1]["json"]
        assert body["isOnlineMeeting"] is False
        assert "onlineMeetingProvider" not in body
        assert body["location"]["displayName"] == "Board Room"

    @pytest.mark.asyncio
    async def test_create_invalid_timezone_rejected(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Not/AZone",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
        )

        assert "error" in result
        assert "IANA timezone" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_create_with_room_emails(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = SAMPLE_EVENT

        await create_event(
            subject="Board Meeting",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            attendees=["alice@company.com"],
            room_emails=["room-a@company.com"],
        )

        body = mock_graph.post.call_args[1]["json"]
        assert len(body["attendees"]) == 2
        assert body["attendees"][0]["type"] == "required"
        assert body["attendees"][0]["emailAddress"]["address"] == "alice@company.com"
        assert body["attendees"][1]["type"] == "resource"
        assert body["attendees"][1]["emailAddress"]["address"] == "room-a@company.com"


class TestRecurrence:
    @pytest.mark.asyncio
    async def test_create_weekly_recurring(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {
            **SAMPLE_EVENT,
            "recurrence": {
                "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
                "range": {
                    "type": "endDate",
                    "startDate": "2026-02-16",
                    "endDate": "2026-06-16",
                },
            },
        }

        result = await create_event(
            subject="Weekly Standup",
            start_datetime="2026-02-16T09:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-16T09:30:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            recurrence_pattern="weekly",
            recurrence_days_of_week=["monday"],
            recurrence_end_date="2026-06-16",
        )

        assert "recurrence" in result
        body = mock_graph.post.call_args[1]["json"]
        assert body["recurrence"]["pattern"]["type"] == "weekly"
        assert body["recurrence"]["pattern"]["daysOfWeek"] == ["monday"]
        assert body["recurrence"]["range"]["type"] == "endDate"
        assert body["recurrence"]["range"]["endDate"] == "2026-06-16"

    @pytest.mark.asyncio
    async def test_create_daily_no_end(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = SAMPLE_EVENT

        await create_event(
            subject="Daily Reminder",
            start_datetime="2026-02-16T08:00:00",
            start_timezone="UTC",
            end_datetime="2026-02-16T08:15:00",
            end_timezone="UTC",
            ctx=mock_ctx,
            recurrence_pattern="daily",
        )

        body = mock_graph.post.call_args[1]["json"]
        assert body["recurrence"]["pattern"]["type"] == "daily"
        assert body["recurrence"]["range"]["type"] == "noEnd"

    @pytest.mark.asyncio
    async def test_create_numbered_recurrence(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = SAMPLE_EVENT

        await create_event(
            subject="Sprint Planning",
            start_datetime="2026-02-16T10:00:00",
            start_timezone="UTC",
            end_datetime="2026-02-16T11:00:00",
            end_timezone="UTC",
            ctx=mock_ctx,
            recurrence_pattern="weekly",
            recurrence_interval=2,
            recurrence_days_of_week=["tuesday"],
            recurrence_count=10,
        )

        body = mock_graph.post.call_args[1]["json"]
        assert body["recurrence"]["pattern"]["interval"] == 2
        assert body["recurrence"]["range"]["type"] == "numbered"
        assert body["recurrence"]["range"]["numberOfOccurrences"] == 10

    @pytest.mark.asyncio
    async def test_recurrence_requires_pattern(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-16T09:00:00",
            start_timezone="UTC",
            end_datetime="2026-02-16T09:30:00",
            end_timezone="UTC",
            ctx=mock_ctx,
            recurrence_interval=1,
        )

        assert "error" in result
        assert "recurrence_pattern is required" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_weekly_requires_days(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-16T09:00:00",
            start_timezone="UTC",
            end_datetime="2026-02-16T09:30:00",
            end_timezone="UTC",
            ctx=mock_ctx,
            recurrence_pattern="weekly",
        )

        assert "error" in result
        assert "days_of_week" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_invalid_pattern_type(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-16T09:00:00",
            start_timezone="UTC",
            end_datetime="2026-02-16T09:30:00",
            end_timezone="UTC",
            ctx=mock_ctx,
            recurrence_pattern="biweekly",
        )

        assert "error" in result
        assert "recurrence_pattern must be one of" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_recurrence_shown_in_list_events(self, mock_ctx, mock_graph):
        recurring_event = {
            **SAMPLE_EVENT,
            "recurrence": {
                "pattern": {"type": "daily", "interval": 1},
                "range": {"type": "noEnd", "startDate": "2026-02-16"},
            },
        }
        mock_graph.get_all.return_value = {"value": [recurring_event]}

        result = await list_events(
            start_datetime="2026-02-16T00:00:00",
            end_datetime="2026-02-16T23:59:59",
            ctx=mock_ctx,
        )

        assert result["events"][0]["recurrence"]["pattern"]["type"] == "daily"


class TestDelegateAccess:
    @pytest.mark.asyncio
    async def test_create_on_delegate_calendar(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = SAMPLE_EVENT

        await create_event(
            subject="Meeting for Boss",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            user_email="boss@company.com",
        )

        call_args = mock_graph.post.call_args
        assert "/users/boss%40company.com/calendar/events" in call_args[0][0]

    @pytest.mark.asyncio
    async def test_update_delegate_event(self, mock_ctx, mock_graph):
        mock_graph.patch.return_value = SAMPLE_EVENT

        await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            subject="Updated by EA",
            user_email="boss@company.com",
        )

        call_args = mock_graph.patch.call_args
        assert "/users/boss%40company.com/events/event-1" in call_args[0][0]

    @pytest.mark.asyncio
    async def test_cancel_delegate_event(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {}

        await cancel_event(
            event_id="event-1",
            ctx=mock_ctx,
            comment="Cancelled by EA",
            user_email="boss@company.com",
        )

        call_args = mock_graph.post.call_args
        assert "/users/boss%40company.com/events/event-1/cancel" in call_args[0][0]

    @pytest.mark.asyncio
    async def test_delegate_access_denied(self, mock_ctx, mock_graph):
        mock_graph.post.side_effect = GraphApiError(
            status_code=403,
            code="ErrorAccessDenied",
            message="Forbidden",
        )

        result = await create_event(
            subject="Test",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            user_email="boss@company.com",
        )

        assert "error" in result
        assert "delegate access" in result["error"]

    @pytest.mark.asyncio
    async def test_update_delegate_fetches_existing(self, mock_ctx, mock_graph):
        """Partial time updates fetch from the delegate's calendar."""
        mock_graph.get.return_value = {
            "start": {"dateTime": "2026-02-16T09:00:00", "timeZone": "Europe/London"},
            "end": {"dateTime": "2026-02-16T09:30:00", "timeZone": "Europe/London"},
        }
        mock_graph.patch.return_value = SAMPLE_EVENT

        await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            start_datetime="2026-02-16T09:15:00",
            user_email="boss@company.com",
        )

        get_args = mock_graph.get.call_args
        assert "/users/boss%40company.com/events/event-1" in get_args[0][0]

    @pytest.mark.asyncio
    async def test_update_delegate_access_denied(self, mock_ctx, mock_graph):
        mock_graph.patch.side_effect = GraphApiError(
            status_code=403,
            code="ErrorAccessDenied",
            message="Forbidden",
        )

        result = await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            subject="New title",
            user_email="boss@company.com",
        )

        assert "error" in result
        assert "delegate access" in result["error"]

    @pytest.mark.asyncio
    async def test_cancel_delegate_without_comment(self, mock_ctx, mock_graph):
        """Cancel via DELETE on a delegate calendar."""
        result = await cancel_event(
            event_id="event-1",
            ctx=mock_ctx,
            user_email="boss@company.com",
        )

        assert result["status"] == "cancelled"
        mock_graph.delete.assert_called_once()
        call_args = mock_graph.delete.call_args
        assert "/users/boss%40company.com/events/event-1" in call_args[0][0]

    @pytest.mark.asyncio
    async def test_cancel_delegate_access_denied(self, mock_ctx, mock_graph):
        mock_graph.delete.side_effect = GraphApiError(
            status_code=403,
            code="ErrorAccessDenied",
            message="Forbidden",
        )

        result = await cancel_event(
            event_id="event-1",
            ctx=mock_ctx,
            user_email="boss@company.com",
        )

        assert "error" in result
        assert "delegate access" in result["error"]


class TestRespondToEvent:
    @pytest.mark.asyncio
    async def test_accept_meeting(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {}

        result = await respond_to_event(
            event_id="event-1",
            response="accept",
            ctx=mock_ctx,
        )

        assert result["status"] == "responded"
        assert result["response"] == "accept"
        call_args = mock_graph.post.call_args
        assert "/me/events/event-1/accept" in call_args[0][0]
        assert call_args[1]["json"]["sendResponse"] is True

    @pytest.mark.asyncio
    async def test_decline_with_comment(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {}

        result = await respond_to_event(
            event_id="event-1",
            response="decline",
            ctx=mock_ctx,
            comment="Conflict with another meeting",
        )

        assert result["response"] == "decline"
        body = mock_graph.post.call_args[1]["json"]
        assert body["comment"] == "Conflict with another meeting"

    @pytest.mark.asyncio
    async def test_tentatively_accept(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {}

        await respond_to_event(
            event_id="event-1",
            response="tentatively_accept",
            ctx=mock_ctx,
        )

        call_args = mock_graph.post.call_args
        assert "/me/events/event-1/tentativelyAccept" in call_args[0][0]

    @pytest.mark.asyncio
    async def test_respond_on_delegate_calendar(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {}

        await respond_to_event(
            event_id="event-1",
            response="accept",
            ctx=mock_ctx,
            user_email="boss@company.com",
        )

        call_args = mock_graph.post.call_args
        assert "/users/boss%40company.com/events/event-1/accept" in call_args[0][0]

    @pytest.mark.asyncio
    async def test_invalid_response_rejected(self, mock_ctx, mock_graph):
        result = await respond_to_event(
            event_id="event-1",
            response="maybe",
            ctx=mock_ctx,
        )

        assert "error" in result
        assert "must be one of" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_silent_response(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {}

        await respond_to_event(
            event_id="event-1",
            response="accept",
            ctx=mock_ctx,
            send_response=False,
        )

        body = mock_graph.post.call_args[1]["json"]
        assert body["sendResponse"] is False

    @pytest.mark.asyncio
    async def test_respond_delegate_access_denied(self, mock_ctx, mock_graph):
        mock_graph.post.side_effect = GraphApiError(
            status_code=403,
            code="ErrorAccessDenied",
            message="Forbidden",
        )

        result = await respond_to_event(
            event_id="event-1",
            response="accept",
            ctx=mock_ctx,
            user_email="boss@company.com",
        )

        assert "error" in result
        assert "permission" in result["error"].lower()
        assert result["errorType"] == "permission_denied"


class TestUpdateEvent:
    @pytest.mark.asyncio
    async def test_update_subject(self, mock_ctx, mock_graph):
        mock_graph.patch.return_value = {**SAMPLE_EVENT, "subject": "Updated Standup"}

        result = await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            subject="Updated Standup",
        )

        assert result["subject"] == "Updated Standup"
        body = mock_graph.patch.call_args[1]["json"]
        assert body == {"subject": "Updated Standup"}

    @pytest.mark.asyncio
    async def test_update_no_fields(self, mock_ctx, mock_graph):
        result = await update_event(event_id="event-1", ctx=mock_ctx)
        assert "error" in result
        mock_graph.patch.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_time(self, mock_ctx, mock_graph):
        mock_graph.patch.return_value = SAMPLE_EVENT

        await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            start_datetime="2026-02-16T10:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-16T10:30:00",
            end_timezone="Europe/London",
        )

        body = mock_graph.patch.call_args[1]["json"]
        assert body["start"]["dateTime"] == "2026-02-16T10:00:00"
        assert body["end"]["dateTime"] == "2026-02-16T10:30:00"

    @pytest.mark.asyncio
    async def test_create_rejects_invalid_email(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            attendees=["not-an-email"],
        )

        assert "error" in result
        assert "not-an-email" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_partial_start_time(self, mock_ctx, mock_graph):
        """Providing start_datetime without timezone fetches the existing timezone."""
        mock_graph.get.return_value = {
            "start": {"dateTime": "2026-02-16T09:00:00", "timeZone": "Europe/London"},
            "end": {"dateTime": "2026-02-16T09:30:00", "timeZone": "Europe/London"},
        }
        mock_graph.patch.return_value = SAMPLE_EVENT

        await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            start_datetime="2026-02-16T09:15:00",
        )

        body = mock_graph.patch.call_args[1]["json"]
        assert body["start"]["dateTime"] == "2026-02-16T09:15:00"
        assert body["start"]["timeZone"] == "Europe/London"

    @pytest.mark.asyncio
    async def test_update_start_only_fetches_existing_end_for_validation(
        self, mock_ctx, mock_graph
    ):
        mock_graph.get.return_value = {
            "start": {"dateTime": "2026-02-16T09:00:00", "timeZone": "Europe/London"},
            "end": {"dateTime": "2026-02-16T09:30:00", "timeZone": "Europe/London"},
        }
        mock_graph.patch.return_value = SAMPLE_EVENT

        await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            start_datetime="2026-02-16T09:15:00",
            start_timezone="Europe/London",
        )

        mock_graph.get.assert_called_once()
        body = mock_graph.patch.call_args[1]["json"]
        assert "start" in body
        assert "end" not in body

    @pytest.mark.asyncio
    async def test_update_enables_online_meeting_provider(self, mock_ctx, mock_graph):
        mock_graph.patch.return_value = SAMPLE_EVENT

        await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            is_online_meeting=True,
        )

        body = mock_graph.patch.call_args[1]["json"]
        assert body["isOnlineMeeting"] is True
        assert body["onlineMeetingProvider"] == "teamsForBusiness"

    @pytest.mark.asyncio
    async def test_update_invalid_datetime_order_rejected(self, mock_ctx, mock_graph):
        result = await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            start_datetime="2026-02-17T16:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
        )

        assert "error" in result
        assert "must be before" in result["error"]
        mock_graph.patch.assert_not_called()


class TestCancelEvent:
    @pytest.mark.asyncio
    async def test_cancel_without_comment(self, mock_ctx, mock_graph):
        result = await cancel_event(event_id="event-1", ctx=mock_ctx)

        assert result["status"] == "cancelled"
        mock_graph.delete.assert_called_once_with("/me/events/event-1")

    @pytest.mark.asyncio
    async def test_cancel_with_comment(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {}

        result = await cancel_event(
            event_id="event-1",
            ctx=mock_ctx,
            comment="Meeting postponed",
        )

        assert result["status"] == "cancelled"
        mock_graph.post.assert_called_once()
        call_args = mock_graph.post.call_args
        assert "/cancel" in call_args[0][0]
        assert call_args[1]["json"]["comment"] == "Meeting postponed"

    @pytest.mark.asyncio
    async def test_cancel_graph_error_normalized(self, mock_ctx, mock_graph):
        mock_graph.delete.side_effect = GraphApiError(
            status_code=404,
            code="ErrorItemNotFound",
            message="Item not found",
        )

        result = await cancel_event(event_id="event-1", ctx=mock_ctx)

        assert result["errorType"] == "not_found"
        assert result["statusCode"] == 404


class TestCreateEventValidation:
    """Additional validation edge cases for create_event."""

    @pytest.mark.asyncio
    async def test_invalid_room_email_rejected(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            room_emails=["not-valid"],
        )

        assert "error" in result
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_invalid_delegate_email_rejected(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            user_email="bad",
        )

        assert "error" in result
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_invalid_end_timezone_rejected(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Not/AZone",
            ctx=mock_ctx,
        )

        assert "error" in result
        assert "IANA timezone" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_end_before_start_rejected(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-17T16:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
        )

        assert "error" in result
        assert "must be before" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_create_with_body(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = SAMPLE_EVENT

        await create_event(
            subject="Test",
            start_datetime="2026-02-17T14:00:00",
            start_timezone="Europe/London",
            end_datetime="2026-02-17T15:00:00",
            end_timezone="Europe/London",
            ctx=mock_ctx,
            body="Meeting agenda here",
        )

        body = mock_graph.post.call_args[1]["json"]
        assert body["body"]["content"] == "Meeting agenda here"

    @pytest.mark.asyncio
    async def test_list_events_non_delegate_graph_error(self, mock_ctx, mock_graph):
        """Non-delegate list_events error (no user_email) uses generic error."""
        mock_graph.get_all.side_effect = GraphApiError(
            status_code=500,
            code="InternalServerError",
            message="Something broke",
        )

        result = await list_events(
            start_datetime="2026-02-16T00:00:00",
            end_datetime="2026-02-16T23:59:59",
            ctx=mock_ctx,
        )

        assert "error" in result
        assert result["statusCode"] == 500


class TestRecurrenceValidation:
    """Additional recurrence validation edge cases."""

    @pytest.mark.asyncio
    async def test_recurrence_interval_below_one(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-16T09:00:00",
            start_timezone="UTC",
            end_datetime="2026-02-16T09:30:00",
            end_timezone="UTC",
            ctx=mock_ctx,
            recurrence_pattern="daily",
            recurrence_interval=0,
        )

        assert "error" in result
        assert "at least 1" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_invalid_days_of_week(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-16T09:00:00",
            start_timezone="UTC",
            end_datetime="2026-02-16T09:30:00",
            end_timezone="UTC",
            ctx=mock_ctx,
            recurrence_pattern="weekly",
            recurrence_days_of_week=["funday"],
        )

        assert "error" in result
        assert "Invalid days" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_bad_end_date_format(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-16T09:00:00",
            start_timezone="UTC",
            end_datetime="2026-02-16T09:30:00",
            end_timezone="UTC",
            ctx=mock_ctx,
            recurrence_pattern="daily",
            recurrence_end_date="Feb 16 2026",
        )

        assert "error" in result
        assert "YYYY-MM-DD" in result["error"]
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_recurrence_count_below_one(self, mock_ctx, mock_graph):
        result = await create_event(
            subject="Test",
            start_datetime="2026-02-16T09:00:00",
            start_timezone="UTC",
            end_datetime="2026-02-16T09:30:00",
            end_timezone="UTC",
            ctx=mock_ctx,
            recurrence_pattern="daily",
            recurrence_count=0,
        )

        assert "error" in result
        assert "at least 1" in result["error"]
        mock_graph.post.assert_not_called()


class TestUpdateEventValidation:
    """Additional update_event validation edge cases."""

    @pytest.mark.asyncio
    async def test_update_invalid_attendees(self, mock_ctx, mock_graph):
        result = await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            attendees=["bad-email"],
        )

        assert "error" in result
        mock_graph.patch.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_invalid_user_email(self, mock_ctx, mock_graph):
        result = await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            subject="New title",
            user_email="bad",
        )

        assert "error" in result
        mock_graph.patch.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_invalid_start_timezone(self, mock_ctx, mock_graph):
        result = await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            start_timezone="Not/AZone",
        )

        assert "error" in result
        assert "IANA timezone" in result["error"]

    @pytest.mark.asyncio
    async def test_update_invalid_end_timezone(self, mock_ctx, mock_graph):
        result = await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            end_timezone="Not/AZone",
        )

        assert "error" in result
        assert "IANA timezone" in result["error"]

    @pytest.mark.asyncio
    async def test_update_fetch_existing_fails(self, mock_ctx, mock_graph):
        """When fetching existing event for partial update fails."""
        mock_graph.get.side_effect = GraphApiError(
            status_code=404,
            code="ErrorItemNotFound",
            message="Not found",
        )

        result = await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            start_datetime="2026-02-16T09:15:00",
        )

        assert "error" in result
        assert result["statusCode"] == 404

    @pytest.mark.asyncio
    async def test_update_attendees_body_location(self, mock_ctx, mock_graph):
        """Update attendees, body, and location together."""
        mock_graph.patch.return_value = SAMPLE_EVENT

        await update_event(
            event_id="event-1",
            ctx=mock_ctx,
            attendees=["alice@company.com"],
            body="New agenda",
            location="Board Room",
        )

        body = mock_graph.patch.call_args[1]["json"]
        assert body["attendees"][0]["emailAddress"]["address"] == "alice@company.com"
        assert body["body"]["content"] == "New agenda"
        assert body["location"]["displayName"] == "Board Room"


class TestCancelEventValidation:
    """Additional cancel_event edge cases."""

    @pytest.mark.asyncio
    async def test_cancel_invalid_user_email(self, mock_ctx, mock_graph):
        result = await cancel_event(
            event_id="event-1",
            ctx=mock_ctx,
            user_email="bad",
        )

        assert "error" in result

    @pytest.mark.asyncio
    async def test_cancel_with_comment_delegate_error(self, mock_ctx, mock_graph):
        """Cancel via /cancel endpoint fails with delegate 403."""
        mock_graph.post.side_effect = GraphApiError(
            status_code=403,
            code="ErrorAccessDenied",
            message="Forbidden",
        )

        result = await cancel_event(
            event_id="event-1",
            ctx=mock_ctx,
            comment="Cancelled",
            user_email="boss@company.com",
        )

        assert "error" in result
        assert "delegate access" in result["error"]


class TestRespondValidation:
    """Additional respond_to_event edge cases."""

    @pytest.mark.asyncio
    async def test_respond_invalid_user_email(self, mock_ctx, mock_graph):
        result = await respond_to_event(
            event_id="event-1",
            response="accept",
            ctx=mock_ctx,
            user_email="bad",
        )

        assert "error" in result

    @pytest.mark.asyncio
    async def test_respond_non_delegate_error(self, mock_ctx, mock_graph):
        """Non-delegate respond error (no user_email) uses generic error."""
        mock_graph.post.side_effect = GraphApiError(
            status_code=500,
            code="InternalServerError",
            message="Something broke",
        )

        result = await respond_to_event(
            event_id="event-1",
            response="accept",
            ctx=mock_ctx,
        )

        assert "error" in result
        assert result["statusCode"] == 500

    @pytest.mark.asyncio
    async def test_respond_organiser_gets_clear_error(self, mock_ctx, mock_graph):
        """Responding to own event returns a clear organiser message."""
        mock_graph.post.side_effect = GraphApiError(
            status_code=400,
            code="ErrorInvalidRequest",
            message="You can't respond to this meeting because you're the meeting organizer.",
        )

        result = await respond_to_event(
            event_id="event-1",
            response="tentatively_accept",
            ctx=mock_ctx,
        )

        assert "error" in result
        assert "organiser" in result["error"].lower()


class TestIsAccessDenied:
    """Tests for the narrowed _is_access_denied helper."""

    @pytest.mark.asyncio
    async def test_403_access_denied_on_delegate(self, mock_ctx, mock_graph):
        """403 + ErrorAccessDenied with user_email is a delegate permission error."""
        mock_graph.post.side_effect = GraphApiError(
            status_code=403,
            code="ErrorAccessDenied",
            message="Access is denied.",
        )

        result = await respond_to_event(
            event_id="event-1",
            response="accept",
            ctx=mock_ctx,
            user_email="boss@company.com",
        )

        assert "error" in result
        assert "permission" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_403_non_access_denied_code_not_treated_as_delegate_error(
        self, mock_ctx, mock_graph
    ):
        """403 with a non-access-denied code should not match _is_access_denied."""
        mock_graph.post.side_effect = GraphApiError(
            status_code=403,
            code="InvalidAuthenticationToken",
            message="Token expired",
        )

        result = await respond_to_event(
            event_id="event-1",
            response="accept",
            ctx=mock_ctx,
            user_email="boss@company.com",
        )

        assert "error" in result
        # Should be a generic error, not the delegate-specific message
        assert "permission to respond" not in result.get("error", "")
