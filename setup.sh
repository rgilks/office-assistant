#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "=== Office Assistant Setup ==="
echo ""

# 1. Check uv is installed
if ! command -v uv &>/dev/null; then
    echo "Installing uv..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
fi

echo "Using $(uv --version)"

# 2. Create venv and install dependencies
echo "Installing dependencies..."
uv sync

# 3. Register MCP server with Claude Code
echo "Registering MCP server with Claude Code..."
claude mcp add \
    --transport stdio \
    --scope project \
    -e DOTENV_PATH="$SCRIPT_DIR/.env" \
    office-assistant -- \
    uv run --directory "$SCRIPT_DIR" python -m office_assistant

# 4. Interactive credential setup and authentication
echo ""
uv run python -m office_assistant.setup

echo ""
echo "=== Setup Complete ==="
echo ""
echo "Type /calendar in Claude Code to start managing your calendar."
