"""Shared helpers for MCP tool modules."""

from __future__ import annotations

import re

from mcp.server.fastmcp import Context

from office_assistant.graph_client import GraphClient

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def get_graph(ctx: Context) -> GraphClient:
    """Extract the ``GraphClient`` from the MCP lifespan context."""
    return ctx.request_context.lifespan_context.graph


def validate_emails(emails: list[str]) -> str | None:
    """Return an error message if any email is clearly invalid, else None."""
    bad = [e for e in emails if not _EMAIL_RE.match(e)]
    if bad:
        return f"Invalid email address(es): {', '.join(bad)}"
    return None
