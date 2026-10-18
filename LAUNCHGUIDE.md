# mcp-advanced-excel

## Tagline
Read-only stdio MCP server that turns messy lab Excel/CSV exports into clean tables.

## Description
Normalizes the spreadsheets that lab instruments and LIMS exports actually produce (HPLC, ICP, balances, plate readers and similar): title blocks, blank lines, multi-row headers, unit rows, merged cells, footer notes, decimal commas. A small stdio Model Context Protocol server that finds the real header and data rectangle, gives columns stable snake_case names, extracts and normalizes units (ug/mL to mg/L, K and degF to degC), and returns JSON rows plus column and unit metadata.

Read-only by default. File access is limited to an explicit directory allowlist (symlinks resolved, `..` rejected), with size and row/column caps. Macros never run and formulas are never evaluated. CSV export exists but stays off until you set both an output directory and an explicit flag. Runs on your machine; no network access, no hosted service.

## Setup Requirements
- `EXCEL_ALLOWED_ROOTS` (required): Directories the server may read (`:` separated on Linux/macOS, `;` on Windows). Empty refuses every path.
- `EXCEL_ALLOW_WRITE` (optional): Set to `1` / `true` / `yes` / `on` to enable `excel_write_csv`. Default is off.
- `EXCEL_OUTPUT_ROOTS` (optional): Existing directories `excel_write_csv` may write to. Needed together with `EXCEL_ALLOW_WRITE`.
- `EXCEL_MAX_FILE_BYTES`, `EXCEL_MAX_ROWS`, `EXCEL_MAX_COLUMNS`, `EXCEL_MAX_SCAN_ROWS` (optional): Caps, clamped to built-in ceilings.
- `EXCEL_UNITS_FILE` (optional): JSON unit overlay, must sit inside the read allowlist.

Install from a checkout (`pip install -e .`). This package is not on PyPI yet, so `uvx --from mcp-advanced-excel` will not work until a publish. Requires Python 3.11+.

## Category
Data & Analytics

## Use Cases
Lab instrument data cleanup, LIMS/ELN import prep, QC result review, Unit harmonization, Spreadsheet-to-JSON for agents

## Features
- Stdio MCP server for any local MCP host
- Reads .xlsx, .xlsm, .xls, .csv, .tsv
- Header and table detection (metadata, blank rows, footers, multi-row and merged headers)
- Stable snake_case column names with unit metadata
- Unit parsing, with numeric conversion only when requested
- Decimal comma and thousands separator parsing; UTF-8/UTF-16/cp1252 input
- `<LOD` / `<LOQ` kept as text and flagged, never invented as a number
- Directory allowlist, symlink and traversal protection, zip-bomb guard, size caps
- No macros, cached formula values only, no network
- Optional, off-by-default CSV or JSON export
- Optional scrubbing of token-shaped secrets (off by default, no SSN rule)

## Getting Started
- "What sheets are in this instrument export?" -> `excel_list_sheets`
- "Where is the real data table on sheet 1?" -> `excel_inspect_layout`
- "Give me the results as clean rows with units in mg/L" -> `excel_read_table`
- "Save the cleaned table as CSV" -> `excel_write_csv` (only when enabled)

## Tags
excel, csv, laboratory, lims, instrument, units, normalization, mcp, stdio, read-only, allowlist, local

## Documentation URL
https://github.com/antonio-castellon/mcp-advanced-excel
