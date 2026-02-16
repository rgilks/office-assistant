"""MCP tools for room and resource discovery."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import Context

from office_assistant.app import mcp
from office_assistant.auth import AuthenticationRequired
from office_assistant.graph_client import GraphApiError
from office_assistant.tools._helpers import (
    auth_required_response,
    get_graph,
    graph_error_response,
)


@mcp.tool()
async def list_rooms(
    ctx: Context,
    building: str | None = None,
) -> dict[str, Any]:
    """List available meeting rooms in the organisation.

    Returns room names, email addresses, and capacity. Use the email
    address with ``room_emails`` in ``create_event`` to book a room.

    Only works with work/school Microsoft 365 accounts.

    Args:
        building: Optional building name to filter by (case-insensitive).
    """
    graph = get_graph(ctx)

    try:
        data = await graph.get_all("/places/microsoft.graph.room", params={"$top": "100"})
    except AuthenticationRequired as exc:
        return auth_required_response(exc)
    except GraphApiError as exc:
        if exc.status_code in {400, 403, 404}:
            return graph_error_response(
                exc,
                fallback_message=(
                    "Room discovery requires a work/school Microsoft 365 account "
                    "with organisational room resources configured."
                ),
            )
        return graph_error_response(exc)

    rooms = []
    for room in data.get("value", []):
        if building and building.lower() not in str(room.get("building", "")).lower():
            continue
        rooms.append(
            {
                "displayName": room.get("displayName"),
                "email": room.get("emailAddress"),
                "capacity": room.get("capacity"),
                "building": room.get("building"),
                "floorLabel": room.get("floorLabel"),
            }
        )

    return {"rooms": rooms, "count": len(rooms)}
