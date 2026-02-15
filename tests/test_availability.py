"""Tests for availability tools."""

from __future__ import annotations

import pytest

from office_assistant.tools.availability import find_meeting_times, get_free_busy


class TestGetFreeBusy:
    @pytest.mark.asyncio
    async def test_empty_emails_rejected(self, mock_ctx, mock_graph):
        result = await get_free_busy(
            emails=[],
            start_datetime="2026-02-16T09:00:00",
            end_datetime="2026-02-16T17:00:00",
            start_timezone="Europe/London",
            ctx=mock_ctx,
        )

        assert "error" in result
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_single_person(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {
            "value": [
                {
                    "scheduleId": "alice@company.com",
                    "availabilityView": "0022200000",
                    "scheduleItems": [
                        {
                            "status": "busy",
                            "subject": "Standup",
                            "start": {"dateTime": "2026-02-16T09:00:00"},
                            "end": {"dateTime": "2026-02-16T09:30:00"},
                        }
                    ],
                    "workingHours": {
                        "startTime": "09:00:00",
                        "endTime": "17:00:00",
                    },
                }
            ]
        }

        result = await get_free_busy(
            emails=["alice@company.com"],
            start_datetime="2026-02-16T09:00:00",
            end_datetime="2026-02-16T17:00:00",
            start_timezone="Europe/London",
            ctx=mock_ctx,
        )

        assert result["count"] == 1
        schedule = result["schedules"][0]
        assert schedule["email"] == "alice@company.com"
        assert len(schedule["scheduleItems"]) == 1
        assert schedule["scheduleItems"][0]["status"] == "busy"
        # Check availability slots are decoded
        assert schedule["availabilitySlots"][0] == "free"
        assert schedule["availabilitySlots"][2] == "busy"

    @pytest.mark.asyncio
    async def test_multiple_people(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {
            "value": [
                {
                    "scheduleId": "alice@company.com",
                    "availabilityView": "00",
                    "scheduleItems": [],
                },
                {
                    "scheduleId": "bob@company.com",
                    "availabilityView": "22",
                    "scheduleItems": [
                        {
                            "status": "busy",
                            "subject": "Meeting",
                            "start": {"dateTime": "2026-02-16T09:00:00"},
                            "end": {"dateTime": "2026-02-16T10:00:00"},
                        }
                    ],
                },
            ]
        }

        result = await get_free_busy(
            emails=["alice@company.com", "bob@company.com"],
            start_datetime="2026-02-16T09:00:00",
            end_datetime="2026-02-16T10:00:00",
            start_timezone="Europe/London",
            ctx=mock_ctx,
        )

        assert result["count"] == 2


class TestFindMeetingTimes:
    @pytest.mark.asyncio
    async def test_empty_attendees_rejected(self, mock_ctx, mock_graph):
        result = await find_meeting_times(
            attendees=[],
            duration_minutes=30,
            ctx=mock_ctx,
        )

        assert "error" in result
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_suggestions_found(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {
            "meetingTimeSuggestions": [
                {
                    "meetingTimeSlot": {
                        "start": {"dateTime": "2026-02-17T10:00:00", "timeZone": "Europe/London"},
                        "end": {"dateTime": "2026-02-17T10:30:00", "timeZone": "Europe/London"},
                    },
                    "confidence": 100.0,
                    "attendeeAvailability": [
                        {
                            "attendee": {"emailAddress": {"address": "alice@company.com"}},
                            "availability": "free",
                        }
                    ],
                    "suggestionReason": "Suggested because all attendees are free.",
                }
            ]
        }

        result = await find_meeting_times(
            attendees=["alice@company.com"],
            duration_minutes=30,
            ctx=mock_ctx,
        )

        assert result["count"] == 1
        suggestion = result["suggestions"][0]
        assert suggestion["confidence"] == 100.0
        assert suggestion["start"] == "2026-02-17T10:00:00"
        assert suggestion["attendeeAvailability"][0]["availability"] == "free"

    @pytest.mark.asyncio
    async def test_no_suggestions(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {
            "meetingTimeSuggestions": [],
            "emptySuggestionsReason": "AttendeesUnavailable",
        }

        result = await find_meeting_times(
            attendees=["alice@company.com", "bob@company.com"],
            duration_minutes=60,
            ctx=mock_ctx,
        )

        assert result["count"] == 0
        assert result["emptySuggestionsReason"] == "AttendeesUnavailable"

    @pytest.mark.asyncio
    async def test_partial_time_constraint_rejected(self, mock_ctx, mock_graph):
        result = await find_meeting_times(
            attendees=["alice@company.com"],
            duration_minutes=30,
            ctx=mock_ctx,
            start_datetime="2026-02-17T09:00:00",
            # end_datetime omitted
        )

        assert "error" in result
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_invalid_email_rejected(self, mock_ctx, mock_graph):
        result = await get_free_busy(
            emails=["not-valid"],
            start_datetime="2026-02-16T09:00:00",
            end_datetime="2026-02-16T17:00:00",
            start_timezone="Europe/London",
            ctx=mock_ctx,
        )

        assert "error" in result
        mock_graph.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_with_time_constraint(self, mock_ctx, mock_graph):
        mock_graph.post.return_value = {"meetingTimeSuggestions": []}

        await find_meeting_times(
            attendees=["alice@company.com"],
            duration_minutes=30,
            ctx=mock_ctx,
            start_datetime="2026-02-17T09:00:00",
            end_datetime="2026-02-21T17:00:00",
            start_timezone="Europe/London",
        )

        body = mock_graph.post.call_args[1]["json"]
        assert "timeConstraint" in body
        slot = body["timeConstraint"]["timeslots"][0]
        assert slot["start"]["dateTime"] == "2026-02-17T09:00:00"
        assert slot["start"]["timeZone"] == "Europe/London"
