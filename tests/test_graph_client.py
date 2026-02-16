"""Tests for GraphClient HTTP behavior."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from office_assistant.auth import AuthenticationRequired
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
    # ErrorAccessDenied is a genuine permission error, NOT re-auth — only 1 request
    assert client._http.request.call_count == 1


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


def test_graph_api_error_str_with_code():
    exc = GraphApiError(status_code=403, message="Forbidden", code="ErrorAccessDenied")
    assert str(exc) == "Graph API error 403 [ErrorAccessDenied]: Forbidden"


def test_graph_api_error_str_without_code():
    exc = GraphApiError(status_code=500, message="Internal error")
    assert str(exc) == "Graph API error 500: Internal error"


def test_parse_retry_after_non_numeric():
    """Non-integer Retry-After header returns None."""
    result = GraphClient._parse_retry_after("Thu, 01 Jan 2026 00:00:00 GMT")
    assert result is None


@pytest.mark.asyncio
async def test_raise_graph_error_invalid_json(client):
    """When the response body is not valid JSON, we still get a structured error."""
    resp = _mock_response(status_code=502)
    resp.json.side_effect = ValueError("No JSON")
    client._http.request = AsyncMock(return_value=resp)

    with pytest.raises(GraphApiError) as exc_info:
        await client.get("/me")

    exc = exc_info.value
    assert exc.status_code == 502
    assert "HTTP 502" in exc.message


@pytest.mark.asyncio
async def test_raise_graph_error_top_level_message(client):
    """When the error payload has a top-level 'message' but no 'error' dict."""
    resp = _mock_response(
        status_code=400,
        json_data={"message": "Bad request parameter"},
    )
    client._http.request = AsyncMock(return_value=resp)

    with pytest.raises(GraphApiError) as exc_info:
        await client.get("/me")

    assert exc_info.value.message == "Bad request parameter"
    assert exc_info.value.code is None


class TestAuthRetry:
    """Tests for automatic re-authentication on 401/403."""

    @pytest.mark.asyncio
    @patch("office_assistant.graph_client.clear_cache")
    async def test_401_triggers_reauth_and_retries(self, mock_clear, client):
        """401 should clear cache, get new token, and retry the request."""
        unauthorized = _mock_response(
            status_code=401,
            json_data={
                "error": {"code": "InvalidAuthenticationToken", "message": "Token expired"},
            },
        )
        success = _mock_response(json_data={"displayName": "Robert"})
        client._http.request = AsyncMock(side_effect=[unauthorized, success])

        result = await client.get("/me")

        assert result == {"displayName": "Robert"}
        mock_clear.assert_called_once()
        assert client._auth_headers.call_count == 2  # initial + retry

    @pytest.mark.asyncio
    @patch("office_assistant.graph_client.clear_cache")
    async def test_403_with_invalid_token_triggers_reauth(self, mock_clear, client):
        """403 + InvalidAuthenticationToken (personal expired token) triggers re-auth."""
        forbidden = _mock_response(
            status_code=403,
            json_data={
                "error": {
                    "code": "InvalidAuthenticationToken",
                    "message": "Access token has expired or is not yet valid.",
                }
            },
        )
        success = _mock_response(json_data={"displayName": "Robert"})
        client._http.request = AsyncMock(side_effect=[forbidden, success])

        result = await client.get("/me")

        assert result == {"displayName": "Robert"}
        mock_clear.assert_called_once()

    @pytest.mark.asyncio
    @patch("office_assistant.graph_client.clear_cache")
    async def test_403_with_access_denied_does_not_trigger_reauth(self, mock_clear, client):
        """403 + ErrorAccessDenied (genuine permission error) should NOT trigger re-auth."""
        forbidden = _mock_response(
            status_code=403,
            json_data={
                "error": {"code": "ErrorAccessDenied", "message": "Genuine permission error"},
            },
        )
        client._http.request = AsyncMock(return_value=forbidden)

        with pytest.raises(GraphApiError) as exc_info:
            await client.get("/me")

        assert exc_info.value.status_code == 403
        mock_clear.assert_not_called()  # no re-auth attempted
        assert client._http.request.call_count == 1  # no retry

    @pytest.mark.asyncio
    @patch("office_assistant.graph_client.clear_cache")
    async def test_persistent_401_raises_after_one_retry(self, mock_clear, client):
        """If re-auth doesn't fix the 401, it should raise normally."""
        unauthorized = _mock_response(
            status_code=401,
            json_data={
                "error": {"code": "InvalidAuthenticationToken", "message": "Token expired"},
            },
        )
        client._http.request = AsyncMock(return_value=unauthorized)

        with pytest.raises(GraphApiError) as exc_info:
            await client.get("/me")

        assert exc_info.value.status_code == 401
        mock_clear.assert_called_once()  # only tried once

    @pytest.mark.asyncio
    @patch("office_assistant.graph_client.clear_cache")
    async def test_auth_retry_catches_authentication_required(self, mock_clear, client):
        """If re-auth raises AuthenticationRequired, return the original error."""
        unauthorized = _mock_response(
            status_code=401,
            json_data={
                "error": {"code": "InvalidAuthenticationToken", "message": "Token expired"},
            },
        )
        client._http.request = AsyncMock(return_value=unauthorized)
        # First call succeeds (initial auth), second raises (re-auth after cache clear)
        client._auth_headers = AsyncMock(
            side_effect=[
                {"Authorization": "Bearer token"},
                AuthenticationRequired(
                    url="https://microsoft.com/devicelogin",
                    user_code="ABC123",
                    message="Sign in",
                    flow={},
                ),
            ]
        )

        with pytest.raises(GraphApiError) as exc_info:
            await client.get("/me")

        assert exc_info.value.status_code == 401
        mock_clear.assert_called_once()

    @pytest.mark.asyncio
    @patch("office_assistant.graph_client.clear_cache")
    async def test_auth_retry_only_on_first_attempt(self, mock_clear, client):
        """Auth retry should only happen on the first attempt, not after transient retries."""
        throttled = _mock_response(
            status_code=429,
            json_data={"error": {"code": "TooManyRequests", "message": "Slow down"}},
            headers={"Retry-After": "0"},
        )
        unauthorized = _mock_response(
            status_code=401,
            json_data={"error": {"code": "InvalidAuthenticationToken", "message": "Expired"}},
        )
        # 429 on attempt 0 (retry), then 401 on attempt 1 (should NOT trigger re-auth)
        client._http.request = AsyncMock(side_effect=[throttled, unauthorized])

        with pytest.raises(GraphApiError) as exc_info:
            await client.get("/me")

        assert exc_info.value.status_code == 401
        mock_clear.assert_not_called()
