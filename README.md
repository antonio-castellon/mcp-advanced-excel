# mcp-advanced-excel

**Read-only MCP server that normalizes messy lab and instrument spreadsheets.**

A small stdio [Model Context Protocol](https://modelcontextprotocol.io/) server.
Point it at an `.xlsx`, `.xlsm`, `.xls`, `.csv`, or `.tsv` export and it returns
a clean table: JSON rows, stable column names, and unit metadata. It is meant
for HPLC, ICP, balances, and the other instruments that write a title block,
a blank line, a header, the numbers, and a footer note into one sheet.

The process runs on your machine. It does not open a port, does not fetch
URLs, and does not evaluate formulas or macros. CSV export exists and is off
until you set both an output directory and an explicit flag.

Version 0.1.0. Available on [PyPI](https://pypi.org/project/mcp-advanced-excel/)
and the [official MCP Registry](https://registry.modelcontextprotocol.io/)
as `io.github.antonio-castellon/mcp-advanced-excel`.

## What it does

- Lists sheets.
- Finds the real header and the data rectangle. Metadata lines, blank rows,
  and footer notes stay out of the table.
- Joins a second header row when it is a unit row or a merged group label
  (`Concentration` over `mg/L` and `ug/mL`).
- Turns headers into snake_case names and pulls units out of parentheses,
  brackets, or a unit row.
- Converts values into a small canonical unit table. `1 ug/mL` and `1 mg/L`
  are the same concentration. Kelvin and Fahrenheit become `degC`.
- Parses decimal commas (`12,5`, `1.234,56`) and thousands separators
  (space, apostrophe, or the other mark).
- Reads UTF-8, UTF-8 with BOM, UTF-16, and Windows cp1252 / Latin-1.
- Fills merged cells from the top-left cell.
- Scrubs token-shaped secrets and labeled passwords before the text is returned.

## Architecture

```
  +---------------------------+
  |  MCP host                 |  Cursor, Claude Desktop, or another stdio client
  +-------------+-------------+
                |  MCP (stdio)
                v
  +---------------------------+
  |  mcp-advanced-excel       |  allowlist, limits, parse, units
  +-------------+-------------+
                |  local files only
                v
         .xlsx / .xlsm / .xls / .csv / .tsv
```

## Install

Python 3.11 or newer.

From PyPI with uv:

```bash
uvx --from mcp-advanced-excel mcp-advanced-excel
```

From a checkout:

```bash
git clone https://github.com/antonio-castellon/mcp-advanced-excel.git
cd mcp-advanced-excel
python -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
```

`pipx run --spec . mcp-advanced-excel` works from the same checkout.

Smoke without a host (the process waits on stdin for MCP JSON-RPC):

```bash
python -m mcp_advanced_excel
```

## Configure

Set `EXCEL_ALLOWED_ROOTS` to the directories the server may read. The
separator is `:` on Linux and macOS and `;` on Windows. An empty value
refuses every path. That is the default.

CSV export needs both of the following. Either one alone leaves the tool
disabled:

- `EXCEL_ALLOW_WRITE=1` (also `true`, `yes`, or `on`)
- `EXCEL_OUTPUT_ROOTS` set to a directory that already exists

Optional caps: `EXCEL_MAX_FILE_BYTES` (default 26214400), `EXCEL_MAX_ROWS`
(default 5000), `EXCEL_MAX_COLUMNS` (default 128), `EXCEL_MAX_SCAN_ROWS`
(default 10000). Values above the built-in ceilings are clamped. See
[docs/SECURITY.md](docs/SECURITY.md).

Optional unit overlay: `EXCEL_UNITS_FILE` points at a JSON file that itself
sits inside the read allowlist. Shape and builtin units:
[docs/UNITS.md](docs/UNITS.md).

### Cursor / Claude Desktop

Cursor (`~/.cursor/mcp.json`) or Claude Desktop (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "advanced-excel": {
      "command": "uvx",
      "args": ["--from", "mcp-advanced-excel", "mcp-advanced-excel"],
      "env": {
        "EXCEL_ALLOWED_ROOTS": "/data/instruments",
        "EXCEL_ALLOW_WRITE": "0",
        "EXCEL_OUTPUT_ROOTS": "/data/cleaned"
      }
    }
  }
}
```

From a checkout, point `command` at the virtualenv interpreter and drop `args`,
or keep `args` as `["-m", "mcp_advanced_excel"]`:

```json
{
  "mcpServers": {
    "advanced-excel": {
      "command": "/path/to/mcp-advanced-excel/.venv/bin/python",
      "args": ["-m", "mcp_advanced_excel"],
      "env": {
        "EXCEL_ALLOWED_ROOTS": "/data/instruments",
        "EXCEL_ALLOW_WRITE": "0"
      }
    }
  }
}
```

A copy of the uvx block lives in `mcp.json` at the repo root. `.env.example`
lists every variable. Do not commit a `.env`.

## Tools

| Tool | Purpose |
| --- | --- |
| `excel_list_sheets` | Sheet names, or the one logical sheet of a CSV |
| `excel_inspect_layout` | Header rows, data bounds, units, and 3 sample rows |
| `excel_read_table` | Full cleaned rows plus column metadata |
| `excel_write_csv` | Write a cleaned CSV. Off unless write is explicitly enabled |

Row numbers in every payload are 1-based sheet rows. `table_index` is 0-based
and follows top-to-bottom order. Omit it and the server returns the table
with the most data rows.

### excel_list_sheets

| Argument | Required | Meaning |
| --- | --- | --- |
| `path` | yes | Local file. Resolved path must sit inside `EXCEL_ALLOWED_ROOTS`. |

```json
{
  "ok": true,
  "path": "/data/instruments/hplc.xlsx",
  "format": "xlsx",
  "sheets": [
    {"name": "Results", "index": 0}
  ]
}
```

`format` is `xlsx`, `xlsm`, `xls`, `csv`, or `tsv`.

### excel_inspect_layout

| Argument | Required | Default | Meaning |
| --- | --- | --- | --- |
| `path` | yes | | Workbook or CSV |
| `sheet` | no | first sheet | Exact sheet name |
| `sheet_index` | no | 0 | Used when `sheet` is omitted |
| `table_index` | no | largest table | 0-based index in sheet order |
| `header_row` | no | auto | 1-based row to use as the only header |
| `decimal` | no | `auto` | `auto`, `comma`, or `dot` |

`header_row` skips detection. Pass `table_index` only as 0 or omit it when
you set `header_row`.

The response has the same shape as `excel_read_table` except the full `rows`
array is replaced by `sample_rows` (at most 3) and `sample_source_rows`.
`row_count` is still the full detected count, subject to the row cap.

```json
{
  "ok": true,
  "sheet": "Wide",
  "header_rows": [3, 4],
  "data_start_row": 5,
  "data_end_row": 5,
  "merged_ranges": 1,
  "metadata": [{"row": 1, "text": "Project: Demo"}],
  "notes": [],
  "columns": [
    {
      "name": "concentration",
      "source_header": "Concentration",
      "source_column": 2,
      "unit": "mg/L",
      "source_unit": "mg/L",
      "dimension": "concentration",
      "unit_converted": true,
      "unit_note": null
    },
    {
      "name": "concentration_2",
      "source_header": "Concentration",
      "source_column": 3,
      "unit": "mg/L",
      "source_unit": "ug/mL",
      "dimension": "concentration",
      "unit_converted": true,
      "unit_note": null
    }
  ],
  "sample_rows": [
    {"sample_id": "S-1", "concentration": 10, "concentration_2": 10, "temperature": 25}
  ]
}
```

That sample is a merged `Concentration` header over `mg/L` and `ug/mL`.
Both values are already in `mg/L` (`1 ug/mL = 1 mg/L`).

### excel_read_table

| Argument | Required | Default | Meaning |
| --- | --- | --- | --- |
| `path` | yes | | Workbook or CSV |
| `sheet` | no | first sheet | Exact sheet name |
| `sheet_index` | no | 0 | Used when `sheet` is omitted |
| `table_index` | no | largest table | 0-based index in sheet order |
| `header_row` | no | auto | 1-based header override |
| `max_rows` | no | `EXCEL_MAX_ROWS` | Cap for this call. Cannot raise the server cap |
| `normalize_units` | no | true | When false, numbers stay in the source unit |
| `decimal` | no | `auto` | `auto`, `comma`, or `dot` |

Example source, stored as text in the sheet so the comma is really a comma:

```
Instrument: Agilent 1260
Method: HPLC-UV

Sample ID | Compound | Conc. (ug/mL) | Temp (K)
S-001     | Caffeine | 12,5          | 298,15
S-002     | Caffeine | 1.234,5       | 300

End of report. Do not distribute.
```

```json
{
  "ok": true,
  "path": "/data/instruments/hplc.xlsx",
  "format": "xlsx",
  "sheet": "Results",
  "sheet_index": 0,
  "table_index": 0,
  "table_count": 1,
  "encoding": null,
  "decimal": "comma",
  "metadata": [
    {"row": 1, "text": "Instrument: Agilent 1260"},
    {"row": 2, "text": "Method: HPLC-UV"}
  ],
  "header_rows": [4],
  "data_start_row": 5,
  "data_end_row": 6,
  "notes": [
    {"row": 8, "text": "End of report. Do not distribute."}
  ],
  "columns": [
    {
      "index": 2,
      "source_column": 3,
      "name": "conc",
      "source_header": "Conc.",
      "unit": "mg/L",
      "source_unit": "ug/mL",
      "dimension": "concentration",
      "unit_converted": true,
      "unit_note": null
    },
    {
      "index": 3,
      "source_column": 4,
      "name": "temp",
      "source_header": "Temp",
      "unit": "degC",
      "source_unit": "K",
      "dimension": "temperature",
      "unit_converted": true,
      "unit_note": null
    }
  ],
  "row_count": 2,
  "truncated": false,
  "redactions": 0,
  "merged_ranges": 0,
  "formula_cells": 0,
  "warnings": [],
  "rows": [
    {"sample_id": "S-001", "compound": "Caffeine", "conc": 12.5, "temp": 25},
    {"sample_id": "S-002", "compound": "Caffeine", "conc": 1234.5, "temp": 26.85}
  ],
  "source_rows": [5, 6]
}
```

`298.15 K` is `25 degC`. `300 K` is `26.85 degC`. `12,5 ug/mL` is `12.5 mg/L`.
`1.234,5` is one thousand two hundred thirty-four and a half, not `1.234`.

A refused path:

```json
{"ok": false, "error": "outside_allowlist", "message": "Path is outside EXCEL_ALLOWED_ROOTS."}
```

Other `error` codes: `roots_unset`, `invalid_path`, `not_a_file`,
`file_too_large`, `unsupported_type`, `sheet_not_found`, `table_not_found`,
`bad_argument`, `read_failed`, `internal_error`.

### excel_write_csv

Same read arguments as `excel_read_table`, plus:

| Argument | Required | Default | Meaning |
| --- | --- | --- | --- |
| `output_path` | yes | | Destination `.csv` inside `EXCEL_OUTPUT_ROOTS` |
| `overwrite` | no | false | Replace an existing file in the output root |

The header row is the clean name, with the canonical unit in brackets when
there is one (`conc [mg/L]`). Text that starts with `=`, `+`, `@`, tab, or
carriage return is prefixed with a quote.

Off by default:

```json
{
  "ok": false,
  "error": "write_disabled",
  "message": "CSV export is off. Set EXCEL_ALLOW_WRITE=1 and EXCEL_OUTPUT_ROOTS to an explicit directory."
}
```

Enabled:

```json
{
  "ok": true,
  "output_path": "/data/cleaned/hplc.csv",
  "rows_written": 2,
  "columns": ["sample_id", "compound", "conc [mg/L]", "temp [degC]"],
  "warnings": []
}
```

Write errors also use `outside_output_root`, `output_exists`,
`output_dir_missing`, `bad_extension`, and `invalid_path` (this last one
covers an attempt to overwrite the source file).

## How detection works

Blank rows split the sheet into blocks, the same idea as
`get_all_tables` in [AdvancedExcel](https://github.com/antonio-castellon/AdvancedExcel).
Inside a block the header is the text row directly above the longest run of
rows that contain a measurement. A second row above that is kept when the
lower row is mostly units, or the upper row is a merged group label
(adjacent cells share the same text after merges are filled).

Rows above the header are `metadata`. Rows after the data run, including a
footer like "End of report", are `notes`. A one-row key/value line
(`Dilution | 10`) is metadata, not a table. When several real tables share
a sheet, the default is the one with the most data rows. Pass `table_index`
for another. `tables` in the payload lists every detected block.

Column names are snake_case. AdvancedExcel uppercases and replaces spaces;
this server lowercases, which is easier to use as a JSON key. Empty headers
become `column_N`. Duplicate names get `_2`, `_3`, and so on.

Numbers with a leading zero (`007`) stay text so instrument codes are not
stripped. `NA`, `N/A`, `ND`, `NaN`, and similar tokens become JSON `null`.
A leading `<`, `>`, `<=`, or `>=` is kept on the converted number (`<0.5`).

`decimal=auto` counts comma-decimal cells against dot-decimal cells. A lone
group of three (`1,234` or `1.234`) does not vote. The default when the vote
ties is dot. Force `comma` or `dot` when you know the instrument locale.

## Security model

Details and the threat model: [docs/SECURITY.md](docs/SECURITY.md).

- Stdio only. No HTTP MCP transport is started.
- Read roots are an allowlist. Symlinks are resolved before the check, so a
  link inside a root that points outside is denied. `..` is denied the same
  way. URLs are denied.
- File size, row, column, and scan caps are always on.
- Macros are not loaded (`keep_vba=False`). Formulas are copied out as text
  and are not evaluated. External workbook links are not followed.
- Writes are off by default and can only create `.csv` files inside
  `EXCEL_OUTPUT_ROOTS`.
- Obvious secrets in cells are replaced with `[REDACTED]` before detection
  and before the response is built. `redactions` counts them.

Dependency floor: `mcp>=1.30.0,<2`. Release 1.30.0 is the first 1.x package
that fixes GHSA-84m7-p3x7-pcfv / CVE-2026-59951 (Streamable HTTP sessions were
kept without a limit). The 2.0.0 through 2.1.1 line is affected as well, and
2.x changes the server API, so this project stays on patched 1.x. Stdio was
not the vulnerable transport; the floor is there so the installed SDK is not
an affected version.

## Limits

| Limit | Default | Ceiling |
| --- | --- | --- |
| File size | 25 MiB | 100 MiB |
| Returned rows | 5000 | 100000 |
| Columns | 128 | 512 |
| Rows scanned | 10000 | 100000 |
| Header rows joined | 2 | 2 |

Supported types are `.xlsx`, `.xlsm`, `.xls`, `.csv`, and `.tsv`. `.xlsb` is
not. Encrypted workbooks fail the read. A two-column block with no numbers
and fewer than three rows is treated as metadata; pass `header_row` to force
it. `ppm` is converted as `1 ppm = 1 mg/L` (dilute aqueous approximation) and
the column carries `unit_note`.

## Relation to AdvancedExcel

[AdvancedExcel](https://github.com/antonio-castellon/AdvancedExcel) is the
pandas library this server borrows from. What was reused, as ideas rather
than as a dependency:

- Contiguous blocks separated by empty rows (`get_all_tables`).
- The first text row of a block as the header, extended here with one parent
  row for units or merged group labels (`merge_rows` is explicit in
  AdvancedExcel; here that second row is automatic when it looks like units).
- Header cleanup into stable names (there: upper snake case; here: lower).
- `NA` / `NAN` / `na` style tokens treated as missing (`replace_nan_strings`).
- Stripping whitespace and dropping empty rows and columns.
- Folding the micro sign to `u` before a unit lookup
  (`replace_by_dictionary` does the same fold).

pandas, numpy, and pyxlsb are not dependencies. openpyxl reads `.xlsx` and
`.xlsm`. xlrd reads `.xls`. CSV uses the standard library. The MCP surface
is new; AdvancedExcel does not speak MCP.

## Development

```bash
pip install -e ".[dev]"
pytest
ruff check src tests
ruff format --check src tests
```

Regenerate the synthetic workbooks under `tests/fixtures/`:

```bash
python tests/fixtures/build_fixtures.py
```

GitHub Actions runs ruff and pytest on Python 3.11, 3.12, and 3.13.

## MCP Registry

Canonical name: `io.github.antonio-castellon/mcp-advanced-excel`

`server.json` is the official registry document for that name. The
`mcp-name` marker below is the PyPI ownership string, same pattern as
[mcp-telegram-bridge](https://github.com/antonio-castellon/mcp-telegram-bridge).

<!-- mcp-name: io.github.antonio-castellon/mcp-advanced-excel -->

## Roadmap

- `.xlsb` if a small reader is worth it.
- Streaming past the scan cap without holding the whole grid.
- A caller-supplied header map for sheets the detector still misses.
- More dimensions in the unit table, still as data rather than code, via
  `EXCEL_UNITS_FILE`.

## License

MIT (c) Antonio Castellon / Castellon.CH
