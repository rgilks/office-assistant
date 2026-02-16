"""MCP server entry point for Office 365 calendar operations."""

from __future__ import annotations

# Import tool modules so their @mcp.tool() decorators register with the server.
import office_assistant.tools.availability
import office_assistant.tools.calendars
import office_assistant.tools.events
import office_assistant.tools.rooms  # noqa: F401
from office_assistant.app import mcp


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
