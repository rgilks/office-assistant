# Office Assistant

Office 365 calendar assistant powered by Claude Code and Microsoft Graph API.

## Architecture

- **MCP server** (`src/office_assistant/server.py`): FastMCP over stdio, provides calendar tools
- **App context** (`src/office_assistant/app.py`): Shared `mcp` instance + lifespan that manages the `GraphClient`
- **Graph client** (`src/office_assistant/graph_client.py`): Async HTTP client with retry, pagination, and error normalization
- **Auth** (`src/office_assistant/auth.py`): MSAL device code flow, tokens cached at `~/.office-assistant/`
- **Tools** (`src/office_assistant/tools/`): MCP tool modules — `events.py`, `calendars.py`, `rooms.py`, `availability.py`
- **Helpers** (`src/office_assistant/tools/_helpers.py`): Shared validation (emails, datetimes, timezones) and error response formatting
- **Skills** (`.claude/skills/`): User-facing `/slash-commands` that map to tool calls

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
- Lint: `uv run ruff check src/ tests/`
- Format: `uv run ruff format src/ tests/`
- Type check: `uv run mypy src/`
- All checks: `uv run pytest -q --cov && uv run ruff check . && uv run ruff format --check . && uv run mypy src/`
- MCP server directly: `uv run python -m office_assistant`

## Testing

- Unit tests mock `GraphClient` methods (`get`, `get_all`, `post`, `patch`, `delete`) via `conftest.py`
- Integration tests use `respx` to mock HTTP responses with the real `GraphClient`
- List endpoints (`list_events`, `list_rooms`, `list_calendars`) use `mock_graph.get_all` (paginated)
- Single-resource endpoints (`update_event` fetch) use `mock_graph.get`
- Coverage minimum: 80% (enforced in CI), currently ~98%

## Key patterns

- All Graph API errors are normalized to `GraphApiError` dataclass with `status_code`, `code`, `message`, `request_id`
- Tool error responses use `graph_error_response()` for consistent shape: `{error, errorType, statusCode, ...}`
- Transient errors (429, 503, 504) are retried with exponential backoff (max 3 attempts)
- Pagination via `graph.get_all()` follows `@odata.nextLink` automatically
- Delegate calendar access uses `/users/{email}` path prefix instead of `/me`
- Personal vs org accounts: `TENANT_ID=consumers` for personal, org tenant GUID for work/school

## Adding a new tool

1. Add the function to the appropriate module in `src/office_assistant/tools/`
2. Decorate with `@mcp.tool()` (import `mcp` from `office_assistant.app`)
3. Use `get_graph(ctx)` to get the `GraphClient`
4. Wrap Graph calls in `try/except GraphApiError` and return `graph_error_response(exc)`
5. Add tests in `tests/test_<module>.py`
6. The tool is auto-registered — no need to update `server.py`

## Common issues

- **"CLIENT_ID and TENANT_ID must be set"**: Run `/calendar-setup` or check `.env` file
- **403 on rooms/free-busy/delegate**: These require a work/school account with appropriate permissions
- **Token expired**: Refresh tokens last ~90 days; after that, re-run `/calendar-setup`
- **Tests fail on `mock_graph.get` vs `get_all`**: List endpoints use `get_all` (paginated), single-resource fetches use `get`
