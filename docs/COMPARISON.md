# Comparison

Checked 18 Oct 2026 against the default-branch READMEs and tool lists.
Star counts are the GitHub stargazer counts on that date. This file does not
claim anything those repositories do not say.

| | mcp-advanced-excel 0.2.0 (this repo) | [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) | [negokaz/excel-mcp-server](https://github.com/negokaz/excel-mcp-server) |
| --- | --- | --- | --- |
| Stars on 18 Oct 2026 | no public star count on that date | 4247 | 1041 |
| Last push seen | this release | 2026-10-18 | 2025-07-19 |
| Language | Python | Python | Node.js |
| Tools | 6 | 46 (TOOLS.md) | 7 (README) |
| Job it is built for | Messy lab and instrument exports | Create and edit workbooks, charts, PivotTables, slicers, formulas | Read and edit xlsx, plus a Windows screenshot |
| Editing, charts, pivots, formatting | No | Yes | Write, copy, format. Screenshot is Windows only |
| Formula calculation | No. Cached value, or null | Yes (`trace_formula`, `goal_seek`) | Values by default (`showFormula` defaults to false) |
| Lab table detection | Yes. Metadata, blank rows, footers, two-row unit headers | Range reads and sheet description. No lab-header detector described | Sheet description and paged reads. No lab-header detector described |
| Decimal commas | Yes | README documents a decimal option | Not described in the README |
| `<LOD` / `<LOQ` | Kept as text and flagged. A number is not invented | Not described | Not described |
| Unit conversion | Opt-in, small built-in table | Not described | Not described |
| Default file access | Deny all until `EXCEL_ALLOWED_ROOTS` is set. Paths must be absolute | Path sandbox relative to an allowed folder | README does not describe a directory allowlist |
| Writes | Off until `EXCEL_ALLOW_WRITE` and `EXCEL_OUTPUT_ROOTS` | Creating and editing workbooks is the product | Write tools are part of the server |
| Transport | stdio only | stdio | stdio |
| Cell text | Marked `content_trust: untrusted` | README says cell contents are untrusted | Not described in the portion checked |

Where the other servers are stronger: general editing, charts, formatting,
PivotTables, slicers, and a much larger tool surface (haris-musa), a formula
calculation engine (haris-musa), and a live Windows screenshot (negokaz).

Where this server is the one to use: an instrument export with a title block,
a unit row, decimal commas, and censored limits, read by an agent that should
not be able to rewrite the workbook.

Sources:

- https://github.com/haris-musa/excel-mcp-server (README and TOOLS.md, 46 tools)
- https://github.com/negokaz/excel-mcp-server (README tool list: `excel_describe_sheets`, `excel_read_sheet`, `excel_screen_capture`, `excel_write_to_sheet`, `excel_create_table`, `excel_copy_sheet`, `excel_format_range`)
