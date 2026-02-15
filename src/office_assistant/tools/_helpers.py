"""Shared helpers for MCP tool modules."""

from __future__ import annotations

from mcp.server.fastmcp import Context

from office_assistant.graph_client import GraphClient


def get_graph(ctx: Context) -> GraphClient:
    """Extract the ``GraphClient`` from the MCP lifespan context."""
    return ctx.request_context.lifespan_context.graph
