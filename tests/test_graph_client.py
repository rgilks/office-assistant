"""Tests for GraphClient HTTP behavior."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from office_assistant.graph_client import GraphClient


@pytest.mark.asyncio
async def test_patch_returns_json_body():
    client = GraphClient()
    try:
        response = MagicMock()
        response.content = b'{"ok": true}'
        response.json.return_value = {"ok": True}
        client._auth_headers = AsyncMock(return_value={"Authorization": "Bearer token"})
        client._http.patch = AsyncMock(return_value=response)

        result = await client.patch("/me/events/event-1", json={"subject": "Updated"})

        assert result == {"ok": True}
        response.raise_for_status.assert_called_once()
    finally:
        await client.close()


@pytest.mark.asyncio
async def test_patch_handles_empty_response_body():
    client = GraphClient()
    try:
        response = MagicMock()
        response.content = b""
        client._auth_headers = AsyncMock(return_value={"Authorization": "Bearer token"})
        client._http.patch = AsyncMock(return_value=response)

        result = await client.patch("/me/events/event-1", json={"subject": "Updated"})

        assert result == {}
        response.raise_for_status.assert_called_once()
        response.json.assert_not_called()
    finally:
        await client.close()
