# Security model

`mcp-advanced-excel` reads local spreadsheets for an MCP host. It is not a
network service. The threat model is an agent, or a prompt, that is talked
into opening a path it should not open, writing a cleaned file somewhere it
should not write, or repeating a secret that was sitting in a cell.

## Transport

The process speaks MCP over stdio only (`FastMCP.run(transport="stdio")`).
It does not start the Streamable HTTP or SSE transports. The dependency floor
is `mcp>=1.30.0,<2`.

That floor is the first 1.x release that fixes GHSA-84m7-p3x7-pcfv /
CVE-2026-59951 (unbounded Streamable HTTP session retention, fixed in 1.30.0
and in 2.2.0). Versions 2.0.0 through 2.1.1 are also affected, and 2.x changes
the server API, so this package stays on the patched 1.x line. Stdio itself
was not the vulnerable transport; the pin is there so a scanner, or a later
edit that enables HTTP, is not sitting on an affected release.

The server never opens a socket of its own. The MCP SDK brings HTTP libraries
for transports this process does not start.

## Allowlist

`EXCEL_ALLOWED_ROOTS` is a path-separator list of directories (`:` on Linux
and macOS, `;` on Windows). When it is empty or unset, every read is refused.

A path is accepted only when all of the following hold:

- It is not a URL and it contains no NUL byte.
- `Path.resolve` succeeds. Symlinks are followed.
- The resolved path is a regular file.
- The resolved path is inside one of the resolved roots (`Path.relative_to`).
  A root of `/data/lab` does not allow `/data/lab-archive`.
- `..` segments are resolved before the check, so
  `/data/lab/../../etc/passwd` is rejected.
- A symlink that lives inside a root and points outside it is rejected.
- The file size is at most `EXCEL_MAX_FILE_BYTES` (default 25 MiB, ceiling
  100 MiB).

The opened path is the resolved path, not the original string.

This is a local control against confused-agent and path-traversal mistakes.
It is not a defence against another user on the same machine who can rewrite
a directory you already allowed, or against hard links.

## Writes

`excel_write_csv` is registered so the host can see it, and it returns
`write_disabled` unless `EXCEL_ALLOW_WRITE` is `1`, `true`, `yes`, or `on`.
The output path must resolve inside `EXCEL_OUTPUT_ROOTS`, must end in `.csv`,
and must not be the source file. Existing files are left alone unless
`overwrite` is true. The parent directory must already exist; the server does
not create directories. Cells that start with `=`, `+`, `@`, tab, or carriage
return are prefixed with a quote so a later spreadsheet open does not treat
them as formulas.

## Workbook contents

- Macros are not loaded (`keep_vba=False` for xlsx/xlsm). They are never
  executed.
- Formulas are returned as text. Nothing in this package evaluates them.
- External links are not followed (`keep_links=False`).
- There is no password or decryption parameter. Encrypted workbooks fail
  the read.
- Scan limits: `EXCEL_MAX_ROWS` (default 5000), `EXCEL_MAX_COLUMNS` (default
  128), `EXCEL_MAX_SCAN_ROWS` (default 10000). Higher settings are clamped
  (100000 rows, 512 columns, 100 MiB).

## Scrubbing

Before header detection, cell text is scrubbed for common token shapes
(`ghp_`, `sk-`, Slack `xox`, AWS `AKIA`, Telegram-style bot tokens, webhook
secrets), labeled passwords, secret-like environment assignments, and
US SSN-shaped numbers. This is a last resort. Do not put secrets in
workbooks the agent is allowed to read.

## Units file

`EXCEL_UNITS_FILE`, when set, must itself resolve inside the read allowlist.
It is a JSON document. It is not code and it is not executed.
