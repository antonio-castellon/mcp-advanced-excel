"""Turn a spreadsheet path into JSON-ready rows and column metadata."""

from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .config import Settings
from .detect import DetectedTable, default_table_index, detect_layout, is_footer_text
from .errors import ToolFailure
from .grid import SheetGrid, check_file_size, file_format, list_sheet_names, read_sheet
from .numbers import DecimalStyle, parse_number, sniff_decimal
from .paths import resolve_read_file, resolve_write_file
from .safety import scrub_text
from .units import UnitSpec, lookup_unit, split_header_cell

_NAME_CLEAN = re.compile(r"[^a-z0-9]+")


@dataclass
class ColumnView:
    index: int
    source_column: int
    name: str
    source_header: str
    unit: str | None
    source_unit: str | None
    dimension: str | None
    unit_converted: bool
    unit_note: str | None = None
    spec: UnitSpec | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "source_column": self.source_column,
            "name": self.name,
            "source_header": self.source_header,
            "unit": self.unit,
            "source_unit": self.source_unit,
            "dimension": self.dimension,
            "unit_converted": self.unit_converted,
            "unit_note": self.unit_note,
        }


@dataclass
class NormalizedTable:
    path: Path
    file_format: str
    sheet: str
    sheet_index: int
    table_index: int | None
    table_count: int
    tables: list[dict[str, Any]]
    encoding: str | None
    decimal: str
    metadata: list[dict[str, Any]]
    header_rows: list[int]
    data_start_row: int | None
    data_end_row: int | None
    notes: list[dict[str, Any]]
    columns: list[ColumnView]
    rows: list[dict[str, Any]]
    source_rows: list[int]
    truncated: bool
    redactions: int
    merged_ranges: int
    formula_cells: int
    warnings: list[str] = field(default_factory=list)

    def as_dict(self, *, include_rows: bool, sample_limit: int = 0) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "ok": True,
            "path": str(self.path),
            "format": self.file_format,
            "sheet": self.sheet,
            "sheet_index": self.sheet_index,
            "table_index": self.table_index,
            "table_count": self.table_count,
            "tables": self.tables,
            "encoding": self.encoding,
            "decimal": self.decimal,
            "metadata": self.metadata,
            "header_rows": self.header_rows,
            "data_start_row": self.data_start_row,
            "data_end_row": self.data_end_row,
            "notes": self.notes,
            "columns": [column.as_dict() for column in self.columns],
            "row_count": len(self.rows),
            "truncated": self.truncated,
            "redactions": self.redactions,
            "merged_ranges": self.merged_ranges,
            "formula_cells": self.formula_cells,
            "warnings": self.warnings,
        }
        if include_rows:
            payload["rows"] = self.rows
            payload["source_rows"] = self.source_rows
        elif sample_limit > 0:
            payload["sample_rows"] = self.rows[:sample_limit]
            payload["sample_source_rows"] = self.source_rows[:sample_limit]
        return payload


def list_sheets(path: str, settings: Settings) -> dict[str, Any]:
    resolved = resolve_read_file(path, settings.allowed_roots)
    check_file_size(resolved, settings.max_file_bytes)
    kind = file_format(resolved)
    names = list_sheet_names(resolved)
    return {
        "ok": True,
        "path": str(resolved),
        "format": kind,
        "sheets": [{"name": name, "index": index} for index, name in enumerate(names)],
    }


def normalize_file(
    path: str,
    settings: Settings,
    *,
    sheet: str | None = None,
    sheet_index: int | None = None,
    table_index: int | None = None,
    header_row: int | None = None,
    max_rows: int | None = None,
    normalize_units: bool = True,
    decimal: str = "auto",
) -> NormalizedTable:
    if decimal not in {"auto", "comma", "dot"}:
        raise ToolFailure("bad_argument", "decimal must be auto, comma, or dot.")
    if header_row is not None and header_row < 1:
        raise ToolFailure("bad_argument", "header_row is a 1-based sheet row.")
    if header_row is not None and table_index not in (None, 0):
        raise ToolFailure(
            "bad_argument",
            "header_row selects one table. Omit table_index or pass 0.",
        )
    if table_index is not None and table_index < 0:
        raise ToolFailure("bad_argument", "table_index must be >= 0.")
    row_cap = settings.max_rows if max_rows is None else min(max_rows, settings.max_rows)
    if row_cap < 1:
        raise ToolFailure("bad_argument", "max_rows must be >= 1.")

    resolved = resolve_read_file(path, settings.allowed_roots)
    check_file_size(resolved, settings.max_file_bytes)
    kind = file_format(resolved)
    grid = read_sheet(
        resolved,
        sheet=sheet,
        sheet_index=sheet_index,
        max_columns=settings.max_columns,
        max_scan_rows=settings.max_scan_rows,
    )
    scrubbed, redactions = _scrub_grid(grid.rows)
    if header_row is not None:
        tables, metadata_rows, selected = _layout_from_header(scrubbed, header_row)
    else:
        tables, metadata_rows = detect_layout(scrubbed)
        if table_index is None:
            selected = default_table_index(tables)
        else:
            if table_index >= len(tables):
                raise ToolFailure(
                    "table_not_found",
                    f"table_index {table_index} is outside 0..{max(len(tables) - 1, 0)}.",
                )
            selected = table_index
    return _materialize(
        resolved=resolved,
        kind=kind,
        grid=grid,
        scrubbed=scrubbed,
        tables=tables,
        metadata_rows=metadata_rows,
        selected=selected,
        redactions=redactions,
        row_cap=row_cap,
        normalize_units=normalize_units,
        decimal=decimal,
        units=settings.units,
    )


def write_cleaned_csv(
    path: str,
    output_path: str,
    settings: Settings,
    *,
    overwrite: bool = False,
    sheet: str | None = None,
    sheet_index: int | None = None,
    table_index: int | None = None,
    header_row: int | None = None,
    max_rows: int | None = None,
    normalize_units: bool = True,
    decimal: str = "auto",
) -> dict[str, Any]:
    if not settings.allow_write:
        raise ToolFailure(
            "write_disabled",
            "CSV export is off. Set EXCEL_ALLOW_WRITE=1 and EXCEL_OUTPUT_ROOTS "
            "to an explicit directory.",
        )
    table = normalize_file(
        path,
        settings,
        sheet=sheet,
        sheet_index=sheet_index,
        table_index=table_index,
        header_row=header_row,
        max_rows=max_rows,
        normalize_units=normalize_units,
        decimal=decimal,
    )
    target = resolve_write_file(output_path, settings.output_roots, overwrite=overwrite)
    if target.resolve() == table.path.resolve():
        raise ToolFailure("invalid_path", "Refusing to overwrite the source file.")
    _write_csv(target, table)
    return {
        "ok": True,
        "output_path": str(target),
        "rows_written": len(table.rows),
        "columns": [_csv_header(column) for column in table.columns],
        "warnings": table.warnings,
    }


def _materialize(
    *,
    resolved: Path,
    kind: str,
    grid: SheetGrid,
    scrubbed: list[list[str]],
    tables: list[DetectedTable],
    metadata_rows: list[int],
    selected: int | None,
    redactions: int,
    row_cap: int,
    normalize_units: bool,
    decimal: str,
    units: dict[str, UnitSpec],
) -> NormalizedTable:
    warnings: list[str] = []
    if grid.truncated_rows:
        warnings.append(
            f"Stopped scanning after {len(grid.rows)} rows (EXCEL_MAX_SCAN_ROWS / EXCEL_MAX_ROWS)."
        )
    if grid.truncated_cols:
        warnings.append("Stopped scanning extra columns at EXCEL_MAX_COLUMNS.")
    if grid.formula_cells:
        warnings.append(
            f"{grid.formula_cells} formula cell(s) were returned as text and were not evaluated."
        )
    if redactions:
        warnings.append(f"Redacted {redactions} sensitive value(s).")

    summaries = [_summary(scrubbed, table, index) for index, table in enumerate(tables)]
    metadata = [_row_record(scrubbed, index) for index in metadata_rows]
    if selected is None:
        warnings.append("No data table detected.")
        return NormalizedTable(
            path=resolved,
            file_format=kind,
            sheet=grid.name,
            sheet_index=grid.index,
            table_index=None,
            table_count=0,
            tables=[],
            encoding=grid.encoding,
            decimal="dot" if decimal == "auto" else decimal,
            metadata=metadata,
            header_rows=[],
            data_start_row=None,
            data_end_row=None,
            notes=[],
            columns=[],
            rows=[],
            source_rows=[],
            truncated=grid.truncated_rows or grid.truncated_cols,
            redactions=redactions,
            merged_ranges=grid.merged_ranges,
            formula_cells=grid.formula_cells,
            warnings=warnings,
        )

    table = tables[selected]
    if len(tables) > 1:
        warnings.append(
            f"Sheet contains {len(tables)} tables; returned table_index {selected}. "
            "Pass table_index to select another."
        )
    columns, kept = _build_columns(scrubbed, table, units, warnings)
    style = _decimal_style(scrubbed, table, kept, decimal)
    rows, source_rows, truncated_cap = _build_rows(
        scrubbed,
        table,
        columns,
        kept,
        decimal=style,
        normalize_units=normalize_units,
        row_cap=row_cap,
    )
    if truncated_cap:
        warnings.append(f"Returned the first {row_cap} data rows.")
    notes = [
        _row_record(scrubbed, index) for index in table.note_rows if _row_text(scrubbed, index)
    ]
    # Preamble rows inside the block are metadata too.
    preamble = [_row_record(scrubbed, index) for index in table.preamble_rows]
    return NormalizedTable(
        path=resolved,
        file_format=kind,
        sheet=grid.name,
        sheet_index=grid.index,
        table_index=selected,
        table_count=len(tables),
        tables=summaries,
        encoding=grid.encoding,
        decimal=style,
        metadata=[*metadata, *preamble],
        header_rows=[index + 1 for index in table.header_rows],
        data_start_row=(table.data_rows[0] + 1) if table.data_rows else None,
        data_end_row=(table.data_rows[-1] + 1) if table.data_rows else None,
        notes=notes,
        columns=columns,
        rows=rows,
        source_rows=source_rows,
        truncated=grid.truncated_rows or grid.truncated_cols or truncated_cap,
        redactions=redactions,
        merged_ranges=grid.merged_ranges,
        formula_cells=grid.formula_cells,
        warnings=warnings,
    )


def _layout_from_header(
    grid: list[list[str]], header_row: int
) -> tuple[list[DetectedTable], list[int], int]:
    index = header_row - 1
    if index >= len(grid) or not any(cell.strip() for cell in grid[index]):
        raise ToolFailure("bad_argument", "header_row does not point at a non-empty row.")
    data_rows: list[int] = []
    note_rows: list[int] = []
    for row_index in range(index + 1, len(grid)):
        if not any(cell.strip() for cell in grid[row_index]):
            break
        text = " ".join(cell.strip() for cell in grid[row_index] if cell.strip())
        if is_footer_text(text) and not any(
            parse_number(cell, "dot").kind in {"number", "censored"} for cell in grid[row_index]
        ):
            note_rows.append(row_index)
            continue
        data_rows.append(row_index)
    # Footer rows after the blank line that ended the loop are notes too.
    table = DetectedTable(
        header_rows=[index],
        data_rows=data_rows,
        preamble_rows=[],
        note_rows=note_rows,
    )
    metadata = [row for row in range(index) if any(cell.strip() for cell in grid[row])]
    return [table], metadata, 0


def _build_columns(
    grid: list[list[str]],
    table: DetectedTable,
    units: dict[str, UnitSpec],
    warnings: list[str],
) -> tuple[list[ColumnView], list[int]]:
    width = max((len(grid[index]) for index in [*table.header_rows, *table.data_rows]), default=0)
    kept: list[int] = []
    for column in range(width):
        if _column_used(grid, table, column):
            kept.append(column)
    views: list[ColumnView] = []
    used_names: dict[str, int] = {}
    for column in kept:
        source_parts: list[str] = []
        source_unit: str | None = None
        for row_index in table.header_rows:
            cell = _cell(grid, row_index, column)
            if not cell:
                continue
            name_part, unit_part = split_header_cell(cell, units)
            if unit_part and source_unit is None:
                source_unit = unit_part
            if name_part:
                source_parts.append(name_part)
        source_header = " ".join(source_parts).strip()
        base = _clean_name(source_header) or f"column_{column + 1}"
        seen = used_names.get(base, 0) + 1
        used_names[base] = seen
        name = base if seen == 1 else f"{base}_{seen}"
        spec = lookup_unit(source_unit, units) if source_unit else None
        if source_unit and spec is None:
            warnings.append(f"Column {name!r} has unrecognized unit {source_unit!r}.")
        views.append(
            ColumnView(
                index=len(views),
                source_column=column + 1,
                name=name,
                source_header=source_header,
                unit=spec.canonical if spec else source_unit,
                source_unit=source_unit,
                dimension=spec.dimension if spec else None,
                unit_converted=False,
                unit_note=spec.note if spec else None,
                spec=spec,
            )
        )
    return views, kept


def _build_rows(
    grid: list[list[str]],
    table: DetectedTable,
    columns: list[ColumnView],
    kept: list[int],
    *,
    decimal: DecimalStyle,
    normalize_units: bool,
    row_cap: int,
) -> tuple[list[dict[str, Any]], list[int], bool]:
    rows: list[dict[str, Any]] = []
    source_rows: list[int] = []
    limited = table.data_rows[:row_cap]
    for row_index in limited:
        record: dict[str, Any] = {}
        for column, source_column in zip(columns, kept, strict=True):
            record[column.name] = _cell_value(
                _cell(grid, row_index, source_column),
                decimal=decimal,
                normalize_units=normalize_units,
                column=column,
            )
        rows.append(record)
        source_rows.append(row_index + 1)
    return rows, source_rows, len(table.data_rows) > row_cap


def _cell_value(
    text: str,
    *,
    decimal: DecimalStyle,
    normalize_units: bool,
    column: ColumnView,
) -> Any:
    if text.startswith("="):
        return text
    parsed = parse_number(text, decimal)
    if parsed.kind in {"empty", "null"}:
        return None
    if parsed.kind == "text" or parsed.number is None:
        return parsed.text if parsed.kind == "text" else text
    value: int | float = parsed.number
    spec = column.spec
    if normalize_units and spec is not None:
        if spec.factor != 1.0 or spec.offset != 0.0:
            value = spec.convert(float(value))
        column.unit_converted = True
    if isinstance(value, float) and value.is_integer() and abs(value) < 10**15:
        value = int(value)
    if parsed.qualifier:
        return f"{parsed.qualifier}{_format_number(value)}"
    if parsed.percent and column.unit is None:
        column.unit = "%"
        column.dimension = column.dimension or "fraction"
    return value


def _decimal_style(
    grid: list[list[str]],
    table: DetectedTable,
    kept: list[int],
    requested: str,
) -> DecimalStyle:
    if requested in {"comma", "dot"}:
        return requested  # type: ignore[return-value]
    samples: list[str] = []
    for row_index in table.data_rows:
        for column in kept:
            samples.append(_cell(grid, row_index, column))
    return sniff_decimal(samples)


def _summary(grid: list[list[str]], table: DetectedTable, index: int) -> dict[str, Any]:
    width = 0
    if table.header_rows or table.data_rows:
        width = max(len(grid[row]) for row in [*table.header_rows, *table.data_rows])
    return {
        "index": index,
        "header_rows": [row + 1 for row in table.header_rows],
        "data_start_row": (table.data_rows[0] + 1) if table.data_rows else None,
        "data_end_row": (table.data_rows[-1] + 1) if table.data_rows else None,
        "data_rows": len(table.data_rows),
        "columns": width,
    }


def _scrub_grid(grid: list[list[str]]) -> tuple[list[list[str]], int]:
    total = 0
    scrubbed: list[list[str]] = []
    for row in grid:
        cleaned: list[str] = []
        for cell in row:
            text, count = scrub_text(cell)
            total += count
            cleaned.append(text)
        scrubbed.append(cleaned)
    return scrubbed, total


def _row_record(grid: list[list[str]], index: int) -> dict[str, Any]:
    return {"row": index + 1, "text": _row_text(grid, index)}


def _row_text(grid: list[list[str]], index: int) -> str:
    if index < 0 or index >= len(grid):
        return ""
    return " | ".join(cell.strip() for cell in grid[index] if cell and cell.strip())


def _column_used(grid: list[list[str]], table: DetectedTable, column: int) -> bool:
    for row_index in [*table.header_rows, *table.data_rows]:
        if _cell(grid, row_index, column):
            return True
    return False


def _cell(grid: list[list[str]], row: int, column: int) -> str:
    if row < 0 or row >= len(grid):
        return ""
    if column < 0 or column >= len(grid[row]):
        return ""
    return grid[row][column].strip()


def _clean_name(text: str) -> str:
    folded = text.strip().lower().replace("\u00b5", "u").replace("\u03bc", "u")
    folded = folded.replace("\u00b0", "deg")
    cleaned = _NAME_CLEAN.sub("_", folded)
    cleaned = cleaned.strip("_")
    if not cleaned:
        return ""
    if cleaned[0].isdigit():
        cleaned = "col_" + cleaned
    return cleaned


def _format_number(value: int | float) -> str:
    if isinstance(value, int):
        return str(value)
    text = f"{value:.10f}".rstrip("0").rstrip(".")
    return text or "0"


def _csv_header(column: ColumnView) -> str:
    if column.unit:
        return f"{column.name} [{column.unit}]"
    return column.name


def _write_csv(path: Path, table: NormalizedTable) -> None:
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\n")
    writer.writerow([_csv_safe(_csv_header(column)) for column in table.columns])
    for row in table.rows:
        writer.writerow([_csv_safe(_csv_cell(row.get(column.name))) for column in table.columns])
    path.write_text(buffer.getvalue(), encoding="utf-8", newline="")


def _csv_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return _format_number(value)
    return str(value)


def _csv_safe(value: str) -> str:
    """Neutralize spreadsheet formula injection in exported text."""
    if value and value[0] in {"=", "+", "@", "\t", "\r"}:
        return "'" + value
    return value
