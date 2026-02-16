"""MCP tools for calendar listing and user profile."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import Context

from office_assistant.app import mcp
from office_assistant.graph_client import GraphApiError
from office_assistant.tools._helpers import get_graph, graph_error_response


@mcp.tool()
async def get_my_profile(ctx: Context) -> dict[str, Any]:
    """Get the authenticated user's display name, email address, and timezone.

    Use this to understand who is currently logged in and what timezone
    they work in.  Call this before other tools if you need the user's
    timezone for date calculations.
    """
    graph = get_graph(ctx)
    try:
        data = await graph.get(
            "/me",
            params={"$select": "displayName,mail,userPrincipalName"},
        )
    except GraphApiError as exc:
        return graph_error_response(exc)

    # mailboxSettings (which contains timezone) is only available for
    # work/school accounts.  Personal Microsoft accounts return 403.
    timezone = None
    try:
        settings = await graph.get("/me/mailboxSettings", params={"$select": "timeZone"})
        timezone = settings.get("timeZone")
    except GraphApiError:
        pass

    return {
        "displayName": data.get("displayName"),
        "email": data.get("mail") or data.get("userPrincipalName"),
        "timezone": timezone,
    }


@mcp.tool()
async def list_calendars(ctx: Context) -> dict[str, Any]:
    """List all calendars the authenticated user has access to.

    Returns each calendar's name, ID, owner, and whether it can be edited.
    Use this when the user asks "which calendars do I have?" or wants to
    know about shared calendars.
    """
    graph = get_graph(ctx)
    try:
        data = await graph.get("/me/calendars", params={"$top": "50"})
    except GraphApiError as exc:
        return graph_error_response(exc)
    calendars = [
        {
            "id": cal.get("id"),
            "name": cal.get("name"),
            "owner": cal.get("owner", {}).get("name"),
            "ownerEmail": cal.get("owner", {}).get("address"),
            "canEdit": cal.get("canEdit", False),
            "isDefaultCalendar": cal.get("isDefaultCalendar", False),
        }
        for cal in data.get("value", [])
    ]
    return {"calendars": calendars, "count": len(calendars)}
