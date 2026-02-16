"""Thin async HTTP client for Microsoft Graph API."""

from __future__ import annotations

import asyncio
import logging
from dataclasses import dataclass
from typing import Any

import httpx

from office_assistant.auth import get_token

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

logger = logging.getLogger(__name__)

# Retry configuration for transient errors (429, 503, 504).
_MAX_RETRIES = 3
_RETRY_STATUS_CODES = {429, 503, 504}
_BASE_BACKOFF_SECONDS = 1.0


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
    expired tokens are automatically refreshed.  Transient errors
    (429, 503, 504) are retried with exponential backoff.
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

    async def _request_with_retry(
        self,
        method: str,
        path: str,
        **kwargs: Any,
    ) -> httpx.Response:
        """Execute an HTTP request with retry on transient failures."""
        headers = await self._auth_headers()
        resp: httpx.Response | None = None

        for attempt in range(_MAX_RETRIES):
            resp = await self._http.request(method, path, headers=headers, **kwargs)

            if resp.status_code not in _RETRY_STATUS_CODES:
                return resp

            # Parse retry delay from server or use exponential backoff.
            retry_after = self._parse_retry_after(resp.headers.get("Retry-After"))
            if retry_after is not None:
                delay = retry_after
            else:
                delay = _BASE_BACKOFF_SECONDS * (2**attempt)

            logger.warning(
                "Graph API %s %s returned %d (attempt %d/%d), retrying in %.1fs",
                method,
                path,
                resp.status_code,
                attempt + 1,
                _MAX_RETRIES,
                delay,
            )
            await asyncio.sleep(delay)

        # All retries exhausted — return the last response so the caller
        # gets a proper GraphApiError via _ensure_success.
        assert resp is not None
        logger.error(
            "Graph API %s %s failed after %d retries with status %d",
            method,
            path,
            _MAX_RETRIES,
            resp.status_code,
        )
        return resp

    async def get(self, path: str, params: dict[str, str] | None = None) -> dict[str, Any]:
        logger.debug("GET %s", path)
        resp = await self._request_with_retry("GET", path, params=params)
        self._ensure_success(resp)
        return resp.json()

    async def get_all(
        self,
        path: str,
        params: dict[str, str] | None = None,
        *,
        max_pages: int = 10,
    ) -> dict[str, Any]:
        """GET with automatic pagination via ``@odata.nextLink``.

        Collects all ``value`` items across pages up to *max_pages*.
        Returns a dict with the merged ``value`` list plus any other
        top-level keys from the first response.
        """
        logger.debug("GET (paginated) %s", path)
        data = await self.get(path, params=params)
        all_items = list(data.get("value", []))

        pages = 1
        next_link = data.get("@odata.nextLink")
        while next_link and pages < max_pages:
            # nextLink is an absolute URL; strip the base so _request_with_retry works.
            relative = next_link.replace(GRAPH_BASE_URL, "")
            logger.debug("Following nextLink (page %d): %s", pages + 1, relative)
            page_data = await self.get(relative)
            all_items.extend(page_data.get("value", []))
            next_link = page_data.get("@odata.nextLink")
            pages += 1

        if next_link:
            logger.warning("Pagination capped at %d pages, results may be incomplete", max_pages)

        data["value"] = all_items
        data.pop("@odata.nextLink", None)
        return data

    async def post(self, path: str, json: dict[str, Any]) -> dict[str, Any]:
        logger.debug("POST %s", path)
        resp = await self._request_with_retry("POST", path, json=json)
        self._ensure_success(resp)
        # Some endpoints (e.g. /events/{id}/cancel) return 202 with no body.
        if resp.status_code == 202 or not resp.content:
            return {}
        return resp.json()

    async def patch(self, path: str, json: dict[str, Any]) -> dict[str, Any]:
        logger.debug("PATCH %s", path)
        resp = await self._request_with_retry("PATCH", path, json=json)
        self._ensure_success(resp)
        if not resp.content:
            return {}
        return resp.json()

    async def delete(self, path: str) -> None:
        logger.debug("DELETE %s", path)
        resp = await self._request_with_retry("DELETE", path)
        self._ensure_success(resp)
