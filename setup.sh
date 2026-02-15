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

# 3. Create .env from example if needed
if [ ! -f .env ]; then
    cp .env.example .env
    echo ""
    echo ">>> Created .env file from template."
    echo ">>> You need to add your Azure credentials before authenticating."
    echo ">>> Run /calendar-setup in Claude Code for step-by-step instructions."
    echo ""
fi

# 4. Register MCP server with Claude Code
echo "Registering MCP server with Claude Code..."
claude mcp add \
    --transport stdio \
    --scope project \
    -e DOTENV_PATH="$SCRIPT_DIR/.env" \
    office-assistant -- \
    uv run --directory "$SCRIPT_DIR" python -m office_assistant

echo ""
echo "=== Setup Complete ==="
echo ""
echo "Next steps:"
echo "  1. Add your Azure credentials to .env (run /calendar-setup for help)"
echo "  2. Start Claude Code in this directory"
echo "  3. Type /calendar-setup to authenticate"
echo "  4. Type /calendar to start managing calendars"
