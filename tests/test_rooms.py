"""Tests for room discovery tools."""

from __future__ import annotations

import pytest

from office_assistant.graph_client import GraphApiError
from office_assistant.tools.rooms import list_rooms


class TestListRooms:
    @pytest.mark.asyncio
    async def test_list_all_rooms(self, mock_ctx, mock_graph):
        mock_graph.get_all.return_value = {
            "value": [
                {
                    "displayName": "Conference Room A",
                    "emailAddress": "room-a@company.com",
                    "capacity": 10,
                    "building": "Building 1",
                    "floorLabel": "Floor 3",
                },
                {
                    "displayName": "Conference Room B",
                    "emailAddress": "room-b@company.com",
                    "capacity": 20,
                    "building": "Building 2",
                    "floorLabel": "Floor 1",
                },
            ]
        }

        result = await list_rooms(ctx=mock_ctx)

        assert result["count"] == 2
        assert result["rooms"][0]["displayName"] == "Conference Room A"
        assert result["rooms"][0]["capacity"] == 10
        assert result["rooms"][0]["email"] == "room-a@company.com"

    @pytest.mark.asyncio
    async def test_filter_by_building(self, mock_ctx, mock_graph):
        mock_graph.get_all.return_value = {
            "value": [
                {
                    "displayName": "Room A",
                    "emailAddress": "a@c.com",
                    "building": "Building 1",
                },
                {
                    "displayName": "Room B",
                    "emailAddress": "b@c.com",
                    "building": "Building 2",
                },
            ]
        }

        result = await list_rooms(ctx=mock_ctx, building="Building 1")

        assert result["count"] == 1
        assert result["rooms"][0]["displayName"] == "Room A"

    @pytest.mark.asyncio
    async def test_building_filter_case_insensitive(self, mock_ctx, mock_graph):
        mock_graph.get_all.return_value = {
            "value": [
                {
                    "displayName": "Room A",
                    "emailAddress": "a@c.com",
                    "building": "HQ Tower",
                },
            ]
        }

        result = await list_rooms(ctx=mock_ctx, building="hq tower")

        assert result["count"] == 1

    @pytest.mark.asyncio
    async def test_personal_account_error(self, mock_ctx, mock_graph):
        mock_graph.get_all.side_effect = GraphApiError(
            status_code=403,
            code="AccessDenied",
            message="Forbidden",
        )

        result = await list_rooms(ctx=mock_ctx)

        assert "error" in result
        assert "work/school" in result["error"]

    @pytest.mark.asyncio
    async def test_no_rooms_returns_empty(self, mock_ctx, mock_graph):
        mock_graph.get_all.return_value = {"value": []}

        result = await list_rooms(ctx=mock_ctx)

        assert result["count"] == 0
        assert result["rooms"] == []

    @pytest.mark.asyncio
    async def test_building_filter_skips_rooms_with_no_building(self, mock_ctx, mock_graph):
        """Rooms with a null/missing building field should not crash the filter."""
        mock_graph.get_all.return_value = {
            "value": [
                {
                    "displayName": "Room A",
                    "emailAddress": "a@c.com",
                    "building": None,
                },
                {
                    "displayName": "Room B",
                    "emailAddress": "b@c.com",
                    "building": "HQ",
                },
            ]
        }

        result = await list_rooms(ctx=mock_ctx, building="HQ")

        assert result["count"] == 1
        assert result["rooms"][0]["displayName"] == "Room B"

    @pytest.mark.asyncio
    async def test_non_auth_error_falls_through(self, mock_ctx, mock_graph):
        """Non-403/400/404 errors use the generic error response."""
        mock_graph.get_all.side_effect = GraphApiError(
            status_code=500,
            code="InternalServerError",
            message="Something broke",
        )

        result = await list_rooms(ctx=mock_ctx)

        assert "error" in result
        assert result["statusCode"] == 500
