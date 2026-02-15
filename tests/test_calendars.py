"""Tests for calendar and profile tools."""

from __future__ import annotations

import pytest

from office_assistant.graph_client import GraphApiError
from office_assistant.tools.calendars import get_my_profile, list_calendars


class TestGetMyProfile:
    @pytest.mark.asyncio
    async def test_returns_profile(self, mock_ctx, mock_graph):
        mock_graph.get.return_value = {
            "displayName": "Alice Smith",
            "mail": "alice@company.com",
            "userPrincipalName": "alice@company.com",
            "mailboxSettings": {"timeZone": "Europe/London"},
        }

        result = await get_my_profile(mock_ctx)

        assert result["displayName"] == "Alice Smith"
        assert result["email"] == "alice@company.com"
        assert result["timezone"] == "Europe/London"
        mock_graph.get.assert_called_once()

    @pytest.mark.asyncio
    async def test_falls_back_to_upn(self, mock_ctx, mock_graph):
        """When mail is null, use userPrincipalName."""
        mock_graph.get.return_value = {
            "displayName": "Bob",
            "mail": None,
            "userPrincipalName": "bob@company.com",
            "mailboxSettings": {},
        }

        result = await get_my_profile(mock_ctx)
        assert result["email"] == "bob@company.com"

    @pytest.mark.asyncio
    async def test_graph_error_is_normalized(self, mock_ctx, mock_graph):
        mock_graph.get.side_effect = GraphApiError(
            status_code=401,
            code="InvalidAuthenticationToken",
            message="Access token expired",
        )

        result = await get_my_profile(mock_ctx)

        assert result["errorType"] == "auth_error"
        assert result["statusCode"] == 401


class TestListCalendars:
    @pytest.mark.asyncio
    async def test_returns_calendars(self, mock_ctx, mock_graph):
        mock_graph.get.return_value = {
            "value": [
                {
                    "id": "cal-1",
                    "name": "Calendar",
                    "owner": {"name": "Alice", "address": "alice@company.com"},
                    "canEdit": True,
                    "isDefaultCalendar": True,
                },
                {
                    "id": "cal-2",
                    "name": "Team Calendar",
                    "owner": {"name": "Team", "address": "team@company.com"},
                    "canEdit": False,
                    "isDefaultCalendar": False,
                },
            ]
        }

        result = await list_calendars(mock_ctx)

        assert result["count"] == 2
        assert result["calendars"][0]["name"] == "Calendar"
        assert result["calendars"][0]["canEdit"] is True
        assert result["calendars"][1]["name"] == "Team Calendar"

    @pytest.mark.asyncio
    async def test_empty_calendars(self, mock_ctx, mock_graph):
        mock_graph.get.return_value = {"value": []}

        result = await list_calendars(mock_ctx)
        assert result["count"] == 0
        assert result["calendars"] == []

    @pytest.mark.asyncio
    async def test_list_calendars_graph_error_is_normalized(self, mock_ctx, mock_graph):
        mock_graph.get.side_effect = GraphApiError(
            status_code=403,
            code="ErrorAccessDenied",
            message="Forbidden",
        )

        result = await list_calendars(mock_ctx)

        assert result["errorType"] == "permission_denied"
        assert result["statusCode"] == 403
