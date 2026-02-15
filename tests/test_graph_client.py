"""Tests for GraphClient HTTP behavior."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from office_assistant.graph_client import GraphApiError, GraphClient


@pytest.mark.asyncio
async def test_patch_returns_json_body():
    client = GraphClient()
    try:
        response = MagicMock()
        response.is_error = False
        response.content = b'{"ok": true}'
        response.json.return_value = {"ok": True}
        client._auth_headers = AsyncMock(return_value={"Authorization": "Bearer token"})
        client._http.patch = AsyncMock(return_value=response)

        result = await client.patch("/me/events/event-1", json={"subject": "Updated"})

        assert result == {"ok": True}
    finally:
        await client.close()


@pytest.mark.asyncio
async def test_patch_handles_empty_response_body():
    client = GraphClient()
    try:
        response = MagicMock()
        response.is_error = False
        response.content = b""
        client._auth_headers = AsyncMock(return_value={"Authorization": "Bearer token"})
        client._http.patch = AsyncMock(return_value=response)

        result = await client.patch("/me/events/event-1", json={"subject": "Updated"})

        assert result == {}
        response.json.assert_not_called()
    finally:
        await client.close()


@pytest.mark.asyncio
async def test_get_raises_normalized_graph_error():
    client = GraphClient()
    try:
        response = MagicMock()
        response.is_error = True
        response.status_code = 403
        response.headers = {"request-id": "req-123"}
        response.json.return_value = {
            "error": {
                "code": "ErrorAccessDenied",
                "message": "Access is denied. Check credentials and try again.",
            }
        }
        client._auth_headers = AsyncMock(return_value={"Authorization": "Bearer token"})
        client._http.get = AsyncMock(return_value=response)

        with pytest.raises(GraphApiError) as exc_info:
            await client.get("/me")

        exc = exc_info.value
        assert exc.status_code == 403
        assert exc.code == "ErrorAccessDenied"
        assert exc.request_id == "req-123"
        assert "Access is denied" in exc.message
    finally:
        await client.close()


@pytest.mark.asyncio
async def test_delete_maps_retry_after_for_throttling():
    client = GraphClient()
    try:
        response = MagicMock()
        response.is_error = True
        response.status_code = 429
        response.headers = {"Retry-After": "30"}
        response.json.return_value = {
            "error": {
                "code": "TooManyRequests",
                "message": "Please retry again later.",
            }
        }
        client._auth_headers = AsyncMock(return_value={"Authorization": "Bearer token"})
        client._http.delete = AsyncMock(return_value=response)

        with pytest.raises(GraphApiError) as exc_info:
            await client.delete("/me/events/event-1")

        exc = exc_info.value
        assert exc.status_code == 429
        assert exc.code == "TooManyRequests"
        assert exc.retry_after_seconds == 30
    finally:
        await client.close()
