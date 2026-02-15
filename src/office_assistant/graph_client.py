"""Thin async HTTP client for Microsoft Graph API."""

from __future__ import annotations

import asyncio
from typing import Any

import httpx

from office_assistant.auth import get_token

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


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

    async def get(self, path: str, params: dict[str, str] | None = None) -> dict[str, Any]:
        resp = await self._http.get(path, headers=await self._auth_headers(), params=params)
        resp.raise_for_status()
        return resp.json()

    async def post(self, path: str, json: dict[str, Any]) -> dict[str, Any]:
        resp = await self._http.post(path, headers=await self._auth_headers(), json=json)
        resp.raise_for_status()
        # Some endpoints (e.g. /events/{id}/cancel) return 202 with no body.
        if resp.status_code == 202 or not resp.content:
            return {}
        return resp.json()

    async def patch(self, path: str, json: dict[str, Any]) -> dict[str, Any]:
        resp = await self._http.patch(path, headers=await self._auth_headers(), json=json)
        resp.raise_for_status()
        if not resp.content:
            return {}
        return resp.json()

    async def delete(self, path: str) -> None:
        resp = await self._http.delete(path, headers=await self._auth_headers())
        resp.raise_for_status()
