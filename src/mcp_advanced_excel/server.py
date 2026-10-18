"""MCP server: read-only lab spreadsheet normalization. Writes are opt-in."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from .config import Settings, load_dotenv_if_present
from .errors import ToolFailure
from .normalize import list_sheets, normalize_file, write_cleaned_csv

INSTRUCTIONS = """\
Read-only stdio MCP server that normalizes messy lab and instrument
spreadsheets (.xlsx, .xlsm, .xls, .csv, .tsv).

It lists sheets, finds the real header and data block (skipping metadata,
blank rows, and footer notes), cleans column names, and normalizes units
with a small built-in table. Decimal commas and thousands separators are
handled. Formulas are returned as text and are never evaluated. Macros are
not loaded.

File access is denied unless the path resolves inside EXCEL_ALLOWED_ROOTS.
excel_write_csv is registered but refuses to run unless EXCEL_ALLOW_WRITE=1
and EXCEL_OUTPUT_ROOTS names an explicit output directory.
"""


def create_server(settings: Settings | None = None) -> FastMCP:
    """Build the FastMCP app. Pass settings in tests; main() loads the environment."""
    cfg = settings if settings is not None else Settings.from_env()

    mcp = FastMCP(
        name="mcp-advanced-excel",
        instructions=INSTRUCTIONS,
        log_level="WARNING",
    )

    @mcp.tool(
        name="excel_list_sheets",
        description=(
            "List sheet names in an Excel workbook or the single logical sheet "
            "of a CSV/TSV file. The path must resolve inside EXCEL_ALLOWED_ROOTS."
        ),
    )
    def excel_list_sheets(path: str) -> dict[str, Any]:
        return _call(list_sheets, path, cfg)

    @mcp.tool(
        name="excel_inspect_layout",
        description=(
            "Detect metadata, header rows, the data rectangle, footer notes, "
            "and column units without returning every data row. "
            "Includes up to 3 sample rows."
        ),
    )
    def excel_inspect_layout(
        path: str,
        sheet: str | None = None,
        sheet_index: int | None = None,
        table_index: int | None = None,
        header_row: int | None = None,
        decimal: str = "auto",
    ) -> dict[str, Any]:
        return _call(
            _inspect,
            path,
            cfg,
            sheet=sheet,
            sheet_index=sheet_index,
            table_index=table_index,
            header_row=header_row,
            decimal=decimal,
        )

    @mcp.tool(
        name="excel_read_table",
        description=(
            "Return a cleaned table as JSON rows plus column and unit metadata. "
            "Skips metadata blocks, blank rows, and footer notes. "
            "Normalizes units (mg/L, ug/mL, degC, K, and the built-in table) "
            "and decimal commas. Set normalize_units=false to keep raw numbers."
        ),
    )
    def excel_read_table(
        path: str,
        sheet: str | None = None,
        sheet_index: int | None = None,
        table_index: int | None = None,
        header_row: int | None = None,
        max_rows: int | None = None,
        normalize_units: bool = True,
        decimal: str = "auto",
    ) -> dict[str, Any]:
        return _call(
            _read,
            path,
            cfg,
            sheet=sheet,
            sheet_index=sheet_index,
            table_index=table_index,
            header_row=header_row,
            max_rows=max_rows,
            normalize_units=normalize_units,
            decimal=decimal,
        )

    @mcp.tool(
        name="excel_write_csv",
        description=(
            "Write the cleaned table to a .csv file inside EXCEL_OUTPUT_ROOTS. "
            "Disabled unless EXCEL_ALLOW_WRITE=1. Never writes outside that "
            "directory and never overwrites the source file."
        ),
    )
    def excel_write_csv(
        path: str,
        output_path: str,
        sheet: str | None = None,
        sheet_index: int | None = None,
        table_index: int | None = None,
        header_row: int | None = None,
        max_rows: int | None = None,
        normalize_units: bool = True,
        decimal: str = "auto",
        overwrite: bool = False,
    ) -> dict[str, Any]:
        return _call(
            write_cleaned_csv,
            path,
            output_path,
            cfg,
            overwrite=overwrite,
            sheet=sheet,
            sheet_index=sheet_index,
            table_index=table_index,
            header_row=header_row,
            max_rows=max_rows,
            normalize_units=normalize_units,
            decimal=decimal,
        )

    mcp._settings = cfg  # type: ignore[attr-defined]
    return mcp


def _inspect(path: str, settings: Settings, **kwargs: Any) -> dict[str, Any]:
    table = normalize_file(path, settings, normalize_units=True, **kwargs)
    return table.as_dict(include_rows=False, sample_limit=3)


def _read(path: str, settings: Settings, **kwargs: Any) -> dict[str, Any]:
    table = normalize_file(path, settings, **kwargs)
    return table.as_dict(include_rows=True)


def _call(func, *args: Any, **kwargs: Any) -> dict[str, Any]:
    try:
        result = func(*args, **kwargs)
    except ToolFailure as exc:
        return {"ok": False, "error": exc.code, "message": str(exc)}
    except Exception as exc:
        detail = str(exc).strip().splitlines()[0] if str(exc).strip() else exc.__class__.__name__
        if len(detail) > 180:
            detail = detail[:180] + "..."
        return {
            "ok": False,
            "error": "internal_error",
            "message": f"{exc.__class__.__name__}: {detail}",
        }
    return result


def list_tool_names(server: FastMCP | None = None) -> list[str]:
    """Return registered tool names without starting stdio."""
    app = server or create_server(Settings())
    tools = app._tool_manager.list_tools()
    return sorted(tool.name for tool in tools)


def main() -> None:
    """Entrypoint: stdio MCP server. Does not open a network port."""
    load_dotenv_if_present()
    settings = Settings.from_env()
    app = create_server(settings)
    app.run(transport="stdio")
