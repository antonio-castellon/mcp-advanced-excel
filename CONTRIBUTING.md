# Contributing

mcp-advanced-excel stays small on purpose. It reads lab spreadsheets for an MCP host. It is read-only unless writing is turned on, and every path stays inside an allowlist.

A bug you can show, a header the reader misses, or a change that keeps that boundary: those fit. A change that opens a port, evaluates formulas, or reads outside the allowlist does not.

To propose a change:

1. Open an issue first if the problem is not clear yet.
2. Branch from `main`.
3. Do not commit secrets, tokens, keys, or a `.env`.
4. Keep the change small and safe. Tests and CI must stay green.

```bash
pip install -e ".[dev]"
ruff check src tests
ruff format --check src tests
pytest
```

CI runs ruff, pytest on Python 3.11, 3.12, and 3.13, and `pip-audit`.
