# Changelog

## 0.3.0

Moves the stdio server to MCP Python SDK 2.x. Tool names, arguments, annotations, and results are unchanged.

- Requires `mcp>=2.3.0,<3`. GHSA-84m7-p3x7-pcfv (CVE-2026-59951) affects 2.0.0 through 2.1.1 and is fixed in 2.2.0, so the floor is above every affected release.
- `create_server()` returns an `MCPServer` (formerly `FastMCP`). Code that annotated the return type as `FastMCP` must change the import.
- The server reports its package version in `serverInfo.version`. SDK 2.x would otherwise send an empty string.
- `uv.lock` pins dependencies with hashes. CI installs with `uv sync --frozen`, and `pip-audit` checks the locked set with `--require-hashes`.
- GitHub Actions are pinned to commit SHAs. The workflow token is limited to `contents: read`.
- Dependabot watches the lockfile instead of pip and ignores major `mcp` updates. The next major is upgraded in a planned PR.
