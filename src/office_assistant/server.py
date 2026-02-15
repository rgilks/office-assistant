"""MCP server entry point for Office 365 calendar operations."""

from __future__ import annotations

import office_assistant.tools.availability

# Import tool modules to register them with the mcp instance.
import office_assistant.tools.calendars
import office_assistant.tools.events  # noqa: F401
from office_assistant.app import mcp


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
