"""Tests for GraphClient HTTP behavior."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from office_assistant.graph_client import GraphApiError, GraphClient


def _mock_response(
    status_code: int = 200,
    json_data: dict | None = None,
    content: bytes = b'{"ok": true}',
    headers: dict | None = None,
) -> MagicMock:
    resp = MagicMock()
    resp.status_code = status_code
    resp.is_error = status_code >= 400
    resp.content = content
    resp.headers = headers or {}
    resp.json.return_value = json_data if json_data is not None else {"ok": True}
    return resp


@pytest.fixture
async def client():
    c = GraphClient()
    c._auth_headers = AsyncMock(return_value={"Authorization": "Bearer token"})
    try:
        yield c
    finally:
        await c.close()


@pytest.mark.asyncio
async def test_patch_returns_json_body(client):
    resp = _mock_response(json_data={"ok": True})
    client._http.request = AsyncMock(return_value=resp)

    result = await client.patch("/me/events/event-1", json={"subject": "Updated"})

    assert result == {"ok": True}


@pytest.mark.asyncio
async def test_patch_handles_empty_response_body(client):
    resp = _mock_response(content=b"")
    client._http.request = AsyncMock(return_value=resp)

    result = await client.patch("/me/events/event-1", json={"subject": "Updated"})

    assert result == {}
    resp.json.assert_not_called()


@pytest.mark.asyncio
async def test_get_raises_normalized_graph_error(client):
    resp = _mock_response(
        status_code=403,
        json_data={
            "error": {
                "code": "ErrorAccessDenied",
                "message": "Access is denied. Check credentials and try again.",
            }
        },
        headers={"request-id": "req-123"},
    )
    client._http.request = AsyncMock(return_value=resp)

    with pytest.raises(GraphApiError) as exc_info:
        await client.get("/me")

    exc = exc_info.value
    assert exc.status_code == 403
    assert exc.code == "ErrorAccessDenied"
    assert exc.request_id == "req-123"
    assert "Access is denied" in exc.message


@pytest.mark.asyncio
async def test_delete_maps_retry_after_for_throttling(client):
    # First 3 calls return 429 (all retry attempts), then raise
    resp = _mock_response(
        status_code=429,
        json_data={"error": {"code": "TooManyRequests", "message": "Please retry again later."}},
        headers={"Retry-After": "0"},
    )
    client._http.request = AsyncMock(return_value=resp)

    with pytest.raises(GraphApiError) as exc_info:
        await client.delete("/me/events/event-1")

    exc = exc_info.value
    assert exc.status_code == 429
    assert exc.code == "TooManyRequests"
    assert exc.retry_after_seconds == 0


@pytest.mark.asyncio
async def test_retry_on_429_then_succeed(client):
    throttled = _mock_response(
        status_code=429,
        json_data={"error": {"code": "TooManyRequests", "message": "Slow down"}},
        headers={"Retry-After": "0"},
    )
    success = _mock_response(json_data={"id": "event-1"})
    client._http.request = AsyncMock(side_effect=[throttled, success])

    result = await client.get("/me/events/event-1")

    assert result == {"id": "event-1"}
    assert client._http.request.call_count == 2


@pytest.mark.asyncio
async def test_retry_on_503_then_succeed(client):
    unavailable = _mock_response(status_code=503, content=b"")
    unavailable.is_error = True
    success = _mock_response(json_data={"ok": True})
    client._http.request = AsyncMock(side_effect=[unavailable, success])

    result = await client.post("/me/calendar/events", json={"subject": "Test"})

    assert result == {"ok": True}
    assert client._http.request.call_count == 2


@pytest.mark.asyncio
async def test_no_retry_on_non_transient_error(client):
    not_found = _mock_response(
        status_code=404,
        json_data={"error": {"code": "ErrorItemNotFound", "message": "Not found"}},
    )
    client._http.request = AsyncMock(return_value=not_found)

    with pytest.raises(GraphApiError) as exc_info:
        await client.get("/me/events/missing")

    assert exc_info.value.status_code == 404
    assert client._http.request.call_count == 1


@pytest.mark.asyncio
async def test_get_all_follows_next_link(client):
    page1 = _mock_response(
        json_data={
            "value": [{"id": "1"}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/events?$skip=1",
        }
    )
    page2 = _mock_response(json_data={"value": [{"id": "2"}]})
    client._http.request = AsyncMock(side_effect=[page1, page2])

    result = await client.get_all("/me/events")

    assert len(result["value"]) == 2
    assert result["value"][0]["id"] == "1"
    assert result["value"][1]["id"] == "2"
    assert "@odata.nextLink" not in result


@pytest.mark.asyncio
async def test_get_all_respects_max_pages(client):
    def make_page(n: int, has_next: bool = True) -> MagicMock:
        data: dict = {"value": [{"id": str(n)}]}
        if has_next:
            data["@odata.nextLink"] = f"https://graph.microsoft.com/v1.0/me/events?$skip={n}"
        return _mock_response(json_data=data)

    # Return pages indefinitely — but max_pages=2 should stop us
    client._http.request = AsyncMock(side_effect=[make_page(1), make_page(2)])

    result = await client.get_all("/me/events", max_pages=2)

    assert len(result["value"]) == 2
    assert client._http.request.call_count == 2


@pytest.mark.asyncio
async def test_get_all_single_page(client):
    page = _mock_response(json_data={"value": [{"id": "1"}]})
    client._http.request = AsyncMock(return_value=page)

    result = await client.get_all("/me/events")

    assert len(result["value"]) == 1
    assert client._http.request.call_count == 1
