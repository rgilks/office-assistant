"""Shared test fixtures."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from office_assistant.graph_client import GraphClient


@pytest.fixture
def mock_graph():
    """A GraphClient with all HTTP methods mocked."""
    client = GraphClient()
    client.get = AsyncMock()
    client.post = AsyncMock()
    client.patch = AsyncMock()
    client.delete = AsyncMock()
    return client


@pytest.fixture
def mock_ctx(mock_graph):
    """A mock MCP Context that provides the graph client via lifespan context."""
    ctx = MagicMock()
    ctx.request_context.lifespan_context.graph = mock_graph
    return ctx


@pytest.fixture(autouse=True)
def patch_auth():
    """Prevent real auth calls during tests.

    Yields the patcher so tests that need the real get_token can
    stop/start it (see test_auth.py::TestGetToken).
    """
    patcher = patch("office_assistant.auth.get_token", return_value="fake-token")
    patcher.start()
    yield patcher
    patcher.stop()
