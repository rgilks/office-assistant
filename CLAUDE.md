# Office Assistant

Office 365 calendar assistant powered by Claude Code and Microsoft Graph API.

## Architecture

- **MCP server** (`src/office_assistant/server.py`): FastMCP over stdio, provides calendar tools
- **Skills** (`.claude/skills/`): User-facing `/slash-commands`
- **Auth** (`src/office_assistant/auth.py`): MSAL device code flow, tokens cached at `~/.office-assistant/`

## Commands

- `/calendar` - Natural language calendar queries (auto-invoked)
- `/schedule-meeting` - Create meetings
- `/find-time` - Find available slots across people
- `/check-availability` - Free/busy lookup
- `/calendar-setup` - Authentication setup

## Development

- Python 3.11+, managed with uv, dependencies in `pyproject.toml`
- Setup: `./setup.sh`
- Install deps: `uv sync --extra dev`
- Tests: `uv run pytest`
- MCP server directly: `uv run python -m office_assistant`
