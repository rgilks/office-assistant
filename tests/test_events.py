"""Tests for event CRUD tools."""

from __future__ import annotations

import pytest

from office_assistant.graph_client import GraphApiError
from office_assistant.tools.events import (
    cancel_event,
    create_event,
    list_events,
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
        mock_graph.get.return_value = {"value": [SAMPLE_EVENT]}

        result = await list_events(
            start_datetime="2026-02-16T00:00:00",
            end_datetime="2026-02-16T23:59:59",
            ctx=mock_ctx,
        )

        assert result["count"] == 1
        assert result["events"][0]["subject"] == "Team Standup"
        assert result["events"][0]["location"] == "Room 42"
        # Should use /me/calendarview
        call_args = mock_graph.get.call_args
        assert "/me/calendarview" in call_args[0][0]

    @pytest.mark.asyncio
    async def test_list_other_user_events(self, mock_ctx, mock_graph):
        mock_graph.get.return_value = {"value": []}

        await list_events(
            start_datetime="2026-02-16T00:00:00",
            end_datetime="2026-02-16T23:59:59",
            ctx=mock_ctx,
            user_email="bob@company.com",
        )

        call_args = mock_graph.get.call_args
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
        mock_graph.get.assert_not_called()

    @pytest.mark.asyncio
    async def test_access_denied(self, mock_ctx, mock_graph):
        mock_graph.get.side_effect = GraphApiError(
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
        mock_graph.get.assert_not_called()


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
        assert body["isOnlineMeeting"] is True
        assert body["onlineMeetingProvider"] == "teamsForBusiness"

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
