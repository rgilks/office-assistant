#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo ""
echo "=== Office Assistant Setup ==="
echo ""
echo "This will install the calendar assistant and connect it to Claude Code."
echo ""

# Step 1: Check uv is installed and install dependencies
echo "[Step 1/3] Installing dependencies..."
if ! command -v uv &>/dev/null; then
    echo "  Installing uv package manager..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
fi
uv sync --quiet 2>/dev/null || uv sync
echo "  Done"
echo ""

# Step 2: Register MCP server with Claude Code
echo "[Step 2/3] Registering with Claude Code..."
claude mcp add \
    --transport stdio \
    --scope project \
    -e DOTENV_PATH="$SCRIPT_DIR/.env" \
    office-assistant -- \
    uv run --directory "$SCRIPT_DIR" python -m office_assistant
echo "  Done"
echo ""

# Step 3: Interactive credential setup and authentication
echo "[Step 3/3] Connecting your Microsoft account..."
echo ""
uv run python -m office_assistant.setup

echo ""
echo "=== All done! ==="
echo ""
echo "Start a new Claude Code conversation and type /calendar to get started."
echo ""
