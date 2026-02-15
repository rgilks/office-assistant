"""Shared FastMCP application instance.

Separated from ``server.py`` to avoid circular imports between the
server entry-point and the tool modules that register themselves on
the ``mcp`` instance.
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass

from mcp.server.fastmcp import FastMCP

from office_assistant.graph_client import GraphClient


@dataclass
class AppContext:
    """Lifespan state shared across all MCP tool invocations."""

    graph: GraphClient


@asynccontextmanager
async def app_lifespan(server: FastMCP) -> AsyncIterator[AppContext]:
    """Create and tear down the Graph client for the MCP session."""
    client = GraphClient()
    try:
        yield AppContext(graph=client)
    finally:
        await client.close()


mcp = FastMCP("office-assistant", lifespan=app_lifespan)
