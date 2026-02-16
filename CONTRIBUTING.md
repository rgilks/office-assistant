# Contributing

Thanks for your interest in contributing to Office Assistant!

## Setup

```bash
git clone https://github.com/rgilks/office-assistant.git
cd office-assistant
uv sync --extra dev
pre-commit install
```

## Development workflow

1. Create a branch for your change
2. Write code and tests
3. Run all checks before committing:

```bash
uv run pytest -q --cov
uv run ruff check src/ tests/
uv run ruff format src/ tests/
uv run mypy src/
```

Pre-commit hooks run ruff lint, ruff format, and pytest automatically on each commit.

## Testing

- Unit tests go in `tests/test_<module>.py`
- Integration tests (with `respx` HTTP mocking) go in `tests/test_integration_flows.py`
- All new tools need tests for both success and error paths
- Coverage minimum is 80% (enforced in CI)

## Adding a new tool

1. Add your function to `src/office_assistant/tools/` with the `@mcp.tool()` decorator
2. Use `get_graph(ctx)` for the Graph client, wrap calls in `try/except GraphApiError`
3. Return errors via `graph_error_response()` for consistent formatting
4. Add unit tests and (optionally) an integration test
5. Tools are auto-registered via the decorator

## Code style

- Ruff handles linting and formatting (configured in `pyproject.toml`)
- Type annotations are required on all functions (`mypy --disallow-untyped-defs`)
- Keep functions focused; use helpers in `_helpers.py` for shared logic

## Pull requests

- Keep PRs focused on a single change
- Ensure all checks pass (tests, lint, format, mypy)
- Add or update tests for any changed behavior
