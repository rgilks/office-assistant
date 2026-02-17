"""Thin async HTTP client for Microsoft Graph API."""

from __future__ import annotations

import asyncio
import logging
from dataclasses import dataclass
from datetime import UTC, datetime
from email.utils import parsedate_to_datetime
from typing import Any

import httpx

from office_assistant.auth import AuthenticationRequired, clear_cache, get_token

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

logger = logging.getLogger(__name__)

# Retry configuration for transient errors (429, 503, 504).
_MAX_RETRIES = 3
_RETRY_STATUS_CODES = {429, 503, 504}
_BASE_BACKOFF_SECONDS = 1.0

# Error codes that indicate an expired/invalid token (as opposed to a genuine
# permission error).  A 401 is always an auth failure.  A 403 with one of the
# codes below is how personal Microsoft accounts signal an expired token —
# other 403 codes (e.g. from org-only endpoints) are real permission errors
# and should NOT trigger a cache clear + re-auth.
_AUTH_FAILURE_CODES = {"invalidauthenticationtoken", "unauthorized"}


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
            # RFC 9110: delay-seconds is a non-negative decimal integer.
            return max(int(value), 0)
        except ValueError:
            pass
        # RFC 9110 also allows an HTTP-date Retry-After value.
        try:
            retry_at = parsedate_to_datetime(value)
        except (TypeError, ValueError):
            return None
        if retry_at.tzinfo is None:
            retry_at = retry_at.replace(tzinfo=UTC)
        now = datetime.now(UTC)
        return max(int((retry_at - now).total_seconds()), 0)

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

    @staticmethod
    def _is_auth_failure(resp: httpx.Response) -> bool:
        """Return True if the response indicates an expired/invalid token.

        A 401 is always an auth failure.  A 403 is only an auth failure
        when the error code is one that Microsoft uses for invalid tokens
        (e.g. ``InvalidAuthenticationToken``).  Other 403s — such as
        ``ErrorAccessDenied`` from org-only endpoints on personal accounts
        — are genuine permission errors and should not trigger re-auth.
        """
        if resp.status_code == 401:
            return True
        if resp.status_code != 403:
            return False
        # Parse the error code from the response body.
        try:
            payload = resp.json()
        except ValueError:
            return False
        error = payload.get("error", {}) if isinstance(payload, dict) else {}
        code = (error.get("code", "") if isinstance(error, dict) else "").lower()
        return code in _AUTH_FAILURE_CODES

    def _ensure_success(self, resp: httpx.Response) -> None:
        if resp.is_error:
            self._raise_graph_error(resp)

    async def _request_with_retry(
        self,
        method: str,
        path: str,
        **kwargs: Any,
    ) -> httpx.Response:
        """Execute an HTTP request with retry on transient failures.

        Also handles auth failures (401/403) by clearing the token cache
        and re-authenticating once before giving up.
        """
        headers = await self._auth_headers()
        resp: httpx.Response | None = None

        for attempt in range(_MAX_RETRIES):
            resp = await self._http.request(method, path, headers=headers, **kwargs)

            # Token expired/revoked — clear cache and get a fresh token once.
            # 401 is always an auth failure.  403 is only an auth failure if
            # the error code indicates an invalid token (personal Microsoft
            # accounts return 403 + InvalidAuthenticationToken for expired
            # tokens).  Other 403s are genuine permission errors.
            if attempt == 0 and self._is_auth_failure(resp):
                logger.warning(
                    "Got %d, clearing token cache and re-authenticating",
                    resp.status_code,
                )
                clear_cache()
                try:
                    headers = await self._auth_headers()
                except AuthenticationRequired:
                    # Token cache was cleared but silent refresh failed —
                    # the user needs to sign in again.  Return the original
                    # response so the caller gets a proper GraphApiError.
                    logger.warning(
                        "Re-authentication requires user sign-in, returning original error",
                    )
                    return resp
                continue

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
        if resp is None:  # pragma: no cover — unreachable when _MAX_RETRIES > 0
            raise RuntimeError("No response received after retries")
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
        result: dict[str, Any] = resp.json()
        return result

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
            logger.debug("Following nextLink (page %d): %s", pages + 1, next_link)
            # Graph returns an opaque URL. Request it as-is per Graph docs.
            page_data = await self.get(next_link)
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
        result: dict[str, Any] = resp.json()
        return result

    async def patch(self, path: str, json: dict[str, Any]) -> dict[str, Any]:
        logger.debug("PATCH %s", path)
        resp = await self._request_with_retry("PATCH", path, json=json)
        self._ensure_success(resp)
        if not resp.content:
            return {}
        result: dict[str, Any] = resp.json()
        return result

    async def delete(self, path: str) -> None:
        logger.debug("DELETE %s", path)
        resp = await self._request_with_retry("DELETE", path)
        self._ensure_success(resp)
