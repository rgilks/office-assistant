"""Shared helpers for MCP tool modules."""

from __future__ import annotations

import logging
import re
from datetime import datetime
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from mcp.server.fastmcp import Context

from office_assistant.graph_client import GraphApiError, GraphClient

logger = logging.getLogger(__name__)

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def get_graph(ctx: Context) -> GraphClient:
    """Extract the ``GraphClient`` from the MCP lifespan context."""
    client: GraphClient = ctx.request_context.lifespan_context.graph
    return client


def validate_emails(emails: list[str]) -> str | None:
    """Return an error message if any email is clearly invalid, else None."""
    bad = [e for e in emails if not _EMAIL_RE.match(e)]
    if bad:
        return f"Invalid email address(es): {', '.join(bad)}"
    return None


def validate_timezone(timezone_name: str, field_name: str) -> str | None:
    """Validate a timezone name as IANA (e.g. Europe/London)."""
    try:
        ZoneInfo(timezone_name)
    except ZoneInfoNotFoundError:
        return f"{field_name} must be a valid IANA timezone (for example, Europe/London)."
    return None


def _parse_iso_datetime(value: str) -> datetime:
    normalized = value.strip()
    if normalized.endswith("Z"):
        normalized = f"{normalized[:-1]}+00:00"
    return datetime.fromisoformat(normalized)


def validate_datetime(value: str, field_name: str) -> str | None:
    """Validate an ISO 8601 datetime string."""
    try:
        _parse_iso_datetime(value)
    except ValueError:
        return f"{field_name} must be a valid ISO 8601 datetime."
    return None


def _coerce_datetime(dt: datetime, tz_name: str | None) -> datetime:
    if not tz_name:
        return dt
    tz = ZoneInfo(tz_name)
    if dt.tzinfo is None:
        return dt.replace(tzinfo=tz)
    return dt.astimezone(tz)


def validate_datetime_order(
    start_datetime: str,
    end_datetime: str,
    *,
    start_field: str = "start_datetime",
    end_field: str = "end_datetime",
    start_timezone: str | None = None,
    end_timezone: str | None = None,
) -> str | None:
    """Validate start/end datetime format and ensure start is before end."""
    if err := validate_datetime(start_datetime, start_field):
        return err
    if err := validate_datetime(end_datetime, end_field):
        return err

    if start_timezone and (err := validate_timezone(start_timezone, "start_timezone")):
        return err
    if end_timezone and (err := validate_timezone(end_timezone, "end_timezone")):
        return err

    start = _coerce_datetime(_parse_iso_datetime(start_datetime), start_timezone)
    end = _coerce_datetime(_parse_iso_datetime(end_datetime), end_timezone)

    if (start.tzinfo is None) != (end.tzinfo is None):
        return (
            f"{start_field} and {end_field} must both include timezone offsets or both omit them."
        )

    if start >= end:
        return f"{start_field} must be before {end_field}."
    return None


def graph_error_response(
    exc: GraphApiError,
    *,
    fallback_message: str | None = None,
) -> dict[str, Any]:
    """Map a Graph API error to a consistent tool response shape."""
    logger.warning(
        "Graph API error %d [%s]: %s (request_id=%s)",
        exc.status_code,
        exc.code,
        exc.message,
        exc.request_id,
    )
    code = (exc.code or "").lower()
    if exc.status_code == 401 or code in {"invalidauthenticationtoken", "unauthorized"}:
        error_type = "auth_error"
        message = (
            fallback_message or "Authentication failed. Run /calendar-setup to sign in again."
        )
    elif exc.status_code == 403 or code in {"erroraccessdenied", "accessdenied"}:
        error_type = "permission_denied"
        message = fallback_message or "You don't have permission to perform this action."
    elif exc.status_code == 404 or code in {"erroritemnotfound", "resourcenotfound"}:
        error_type = "not_found"
        message = fallback_message or "The requested resource was not found."
    elif exc.status_code == 429:
        error_type = "throttled"
        message = fallback_message or "Microsoft Graph is rate-limiting requests. Please retry."
    elif exc.status_code == 400 or code.startswith("errorinvalid"):
        error_type = "validation_error"
        message = fallback_message or exc.message
    else:
        error_type = "graph_error"
        message = fallback_message or exc.message

    payload: dict[str, Any] = {
        "error": message,
        "errorType": error_type,
        "statusCode": exc.status_code,
    }
    if exc.code:
        payload["errorCode"] = exc.code
    if exc.request_id:
        payload["requestId"] = exc.request_id
    if exc.retry_after_seconds is not None:
        payload["retryAfterSeconds"] = exc.retry_after_seconds
    return payload
