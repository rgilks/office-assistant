"""Thin async HTTP client for Microsoft Graph API."""

from __future__ import annotations

import asyncio
from dataclasses import dataclass
from typing import Any

import httpx

from office_assistant.auth import get_token

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


@dataclass(slots=True)
class GraphApiError(Exception):
    """Normalized Graph API failure with structured metadata."""

    status_code: int
    message: str
    code: str | None = None
    request_id: str | None = None
    retry_after_seconds: int | None = None

    def __str__(self) -> str:
        code = f" [{self.code}]" if self.code else ""
        return f"Graph API error {self.status_code}{code}: {self.message}"


class GraphClient:
    """Async wrapper around the Microsoft Graph REST API.

    Each method acquires a fresh token (via the MSAL cache) so that
    expired tokens are automatically refreshed.
    """

    def __init__(self) -> None:
        self._http = httpx.AsyncClient(base_url=GRAPH_BASE_URL, timeout=30.0)

    async def close(self) -> None:
        """Close the underlying HTTP connection pool."""
        await self._http.aclose()

    async def _auth_headers(self) -> dict[str, str]:
        # get_token() is synchronous (MSAL) and may block during device-code
        # flow, so run it in a thread to avoid blocking the event loop.
        token = await asyncio.to_thread(get_token)
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    @staticmethod
    def _parse_retry_after(value: str | None) -> int | None:
        if value is None:
            return None
        try:
            return int(value)
        except ValueError:
            return None

    def _raise_graph_error(self, resp: httpx.Response) -> None:
        code: str | None = None
        message = f"HTTP {resp.status_code}"
        request_id = resp.headers.get("request-id") or resp.headers.get("x-ms-request-id")
        retry_after_seconds = self._parse_retry_after(resp.headers.get("Retry-After"))

        try:
            payload = resp.json()
        except ValueError:
            payload = None

        if isinstance(payload, dict):
            err = payload.get("error")
            if isinstance(err, dict):
                code = err.get("code")
                message = err.get("message") or message
            elif isinstance(payload.get("message"), str):
                message = payload["message"]

        raise GraphApiError(
            status_code=resp.status_code,
            code=code,
            message=message,
            request_id=request_id,
            retry_after_seconds=retry_after_seconds,
        )

    def _ensure_success(self, resp: httpx.Response) -> None:
        if resp.is_error:
            self._raise_graph_error(resp)

    async def get(self, path: str, params: dict[str, str] | None = None) -> dict[str, Any]:
        resp = await self._http.get(path, headers=await self._auth_headers(), params=params)
        self._ensure_success(resp)
        return resp.json()

    async def post(self, path: str, json: dict[str, Any]) -> dict[str, Any]:
        resp = await self._http.post(path, headers=await self._auth_headers(), json=json)
        self._ensure_success(resp)
        # Some endpoints (e.g. /events/{id}/cancel) return 202 with no body.
        if resp.status_code == 202 or not resp.content:
            return {}
        return resp.json()

    async def patch(self, path: str, json: dict[str, Any]) -> dict[str, Any]:
        resp = await self._http.patch(path, headers=await self._auth_headers(), json=json)
        self._ensure_success(resp)
        if not resp.content:
            return {}
        return resp.json()

    async def delete(self, path: str) -> None:
        resp = await self._http.delete(path, headers=await self._auth_headers())
        self._ensure_success(resp)
