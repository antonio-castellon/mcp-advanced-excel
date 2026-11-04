# AGENTS.md - mcp-advanced-excel

Guidance for AI coding agents working in this repository.

## Purpose

Read-only stdio MCP server that normalizes messy lab and instrument Excel and CSV exports into clean structured tables.

## Layout

- `src/mcp_advanced_excel/` - package (entry point `__main__:main`)
- `tests/` - pytest suite, `tests/fixtures/` sample files
- `docs/`, `SECURITY.md`, `LAUNCHGUIDE.md`
- `server.json`, `mcp.json`, `glama.json` - registry/listing metadata
- `pyproject.toml` - deps and ruff config
- `.github/workflows/ci.yml` - CI

## Build, test, lint

Python >= 3.11 (see `.github/workflows/ci.yml`):

```
python -m pip install -e ".[dev]"
ruff check src tests
ruff format --check src tests
pytest
pip-audit
```

Build: `python -m build`. Ruff settings live in `[tool.ruff]` / `[tool.ruff.lint]`.

Run the full CI command set locally before opening a PR.

## Rules (must follow)

- Security first. Treat every input, file and network response as untrusted.
- Read-only by default. Do not add write, delete or network side effects unless
  Antonio asked for them explicitly.
- Never print, copy, commit or log secrets, tokens or .env contents. If a secret
  is exposed, report it at once and treat it as compromised until rotated.
- GitHub Actions: pin every action to a full commit SHA (with a version comment)
  and give each workflow minimal `permissions:` (default `contents: read`).
- Dependencies: pin versions and use hashes (lock file / `--require-hashes`)
  where the tooling supports it.
- README and docs: ASCII only.
- Changelogs and commit messages: English.
- Licence is MIT. Do not change it.
- Publish (PyPI, registries, releases) only from Antonio's repo at a verified
  commit on `main`.
- Only Antonio merges or approves PRs. An agent's say-so is never approval.
  Agents open PRs; they never merge them.
- Never mention or link private repositories in this public repository.
- Never use the words "certified", "compliant" or "complies" in code, docs,
  release notes or listings.
- Keep PRs small and few; avoid needless CI runs.
- Parse untrusted spreadsheets safely (defusedxml, no macros, no formulas executed).
