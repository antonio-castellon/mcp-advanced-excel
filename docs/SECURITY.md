# Security model

To report a vulnerability, use the private Security tab. See [SECURITY.md](../SECURITY.md).

`mcp-advanced-excel` reads local spreadsheets for an MCP host. It is not a
network service. The threat model is an agent, or a prompt, that is talked
into opening a path it should not open, writing a cleaned file somewhere it
should not write, or repeating a secret that was sitting in a cell.

## Transport

The process speaks MCP over stdio only (`MCPServer.run(transport="stdio")`).
It does not start the Streamable HTTP or SSE transports. The dependency floor
is `mcp>=2.3.0,<3`.

That floor is above every release affected by GHSA-84m7-p3x7-pcfv /
CVE-2026-59951 (unbounded Streamable HTTP session retention, fixed in 1.30.0
and in 2.2.0). On the 2.x line, versions 2.0.0 through 2.1.1 are affected. Stdio itself
was not the vulnerable transport; the pin is there so a scanner, or a later
edit that enables HTTP, is not sitting on an affected release.

The server never opens a socket of its own. The MCP SDK brings HTTP libraries
for transports this process does not start.

## Allowlist

`EXCEL_ALLOWED_ROOTS` is a path-separator list of directories (`:` on Linux
and macOS, `;` on Windows). When it is empty or unset, every read is refused.

A path is accepted only when all of the following hold:

- It is absolute. Relative paths are rejected so the server cwd is not a root.
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
- For xlsx/xlsm, the zip's declared uncompressed size is at most 200 MiB and
  the compression ratio is at most 100. A zip that fails either check is
  refused before the sheet is parsed.

The opened path is the resolved path, not the original string.

This is a local control against confused-agent and path-traversal mistakes.
It is not a defence against another user on the same machine who can rewrite
a directory you already allowed, or against hard links.

## Writes

`excel_write_csv` is registered so the host can see it, and it returns
`write_disabled` unless `EXCEL_ALLOW_WRITE` is `1`, `true`, `yes`, or `on`.
The output path must resolve inside `EXCEL_OUTPUT_ROOTS`, must end in `.csv`,
and must not be the source file. The filename must end in `.csv` or `.json`.
Existing files are left alone unless `overwrite` is true. A symlink is refused
even with `overwrite`. The write is a temp file in the output directory,
fsync, then `os.replace`, so the final path is not opened by following a
link. The parent directory must already exist; the server does not create
directories. Cells whose first non-space character is `=`, `+`, `-`, `@`,
tab, carriage return, or a fullwidth equals sign (U+FF1D) are prefixed with
a quote so a later spreadsheet open does not treat them as formulas.

## Workbook contents

- xlsx and xlsm are read with defusedxml. The sheet stream stops after the
  rows the call asked for. VBA is not loaded and is never executed. External
  links are not followed.
- Cached formula values are returned when the file stored them. A formula
  with no cached value is null, plus a warning. Formula text is included
  only when `include_formulas` is true. Nothing in this package evaluates
  a formula.
- There is no password or decryption parameter. Encrypted workbooks fail
  the read.
- Every payload that echoes workbook content sets `content_trust` to
  `untrusted`. Hidden sheets are flagged and still readable. Cell text is
  data, not instructions.
- Page size: `EXCEL_MAX_ROWS` (default 5000). Width: `EXCEL_MAX_COLUMNS`
  (default 128). A call will not stream past `EXCEL_MAX_SCAN_ROWS`
  (default 100000). Higher settings are clamped (100000 rows, 512 columns,
  100 MiB). Detection looks at a prefix of 2000 rows, not the whole sheet.
- The process does not load a `.env` from the working directory. Set
  `EXCEL_DOTENV` to an explicit file if you want one.

## Scrubbing

Scrubbing is off unless `EXCEL_REDACT=secrets`. That mode replaces token
shapes (`ghp_`, `sk-`, Slack `xox`, AWS `AKIA`, Telegram-style bot tokens,
webhook secrets), labeled passwords, and secret-like environment assignments.
The assignment name is kept (`API_KEY=[REDACTED]`). Coordinates and the kind
of match are reported. The secret is not. US SSN-shaped numbers are not
redacted, because the same pattern is a common lab id. This is a last resort.
Do not put secrets in workbooks the agent is allowed to read.

## Units file

`EXCEL_UNITS_FILE`, when set, must itself resolve inside the read allowlist.
It is a JSON document. It is not code and it is not executed.
