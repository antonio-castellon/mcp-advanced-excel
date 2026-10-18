"""Turn a spreadsheet path into JSON-ready rows and column metadata."""

from __future__ import annotations

import csv
import io
import json
import os
import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl.utils import range_boundaries

from .config import Settings
from .detect import DetectedTable, default_table_index, detect_layout, is_footer_text
from .errors import ToolFailure
from .grid import (
    DETECT_PREFIX_ROWS,
    SheetGrid,
    check_file_size,
    file_format,
    list_sheet_infos,
    read_window,
)
from .numbers import DecimalStyle, looks_numeric, parse_number, sniff_decimal
from .paths import resolve_read_file, resolve_write_file
from .plugins import SheetView, get_plugin_host
from .safety import scrub_text
from .units import UnitSpec, lookup_unit, split_header_cell, unit_key

_NAME_CLEAN = re.compile(r"[^a-z0-9]+")
_FORMULA_LEAD = set("=+-@\t\r")
_FULLWIDTH_EQUALS = "\uff1d"
_CONVERSION_LIST_CAP = 50
_REDACTION_LIST_CAP = 50
_OVERVIEW_PREFIX_ROWS = 40

CONTENT_TRUST = "untrusted"


@dataclass
class ColumnView:
    index: int
    source_column: int
    name: str
    source_header: str
    unit: str | None
    source_unit: str | None
    canonical_unit: str | None
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
            "canonical_unit": self.canonical_unit,
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
    sheet_state: str
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
    redaction_events: list[dict[str, Any]]
    merged_ranges: int
    formula_cells: int
    offset: int = 0
    next_offset: int | None = None
    continues: bool = False
    declared_rows: int | None = None
    declared_cols: int | None = None
    hidden_columns: list[int] = field(default_factory=list)
    censored_cells: list[dict[str, Any]] = field(default_factory=list)
    conversions: list[dict[str, Any]] = field(default_factory=list)
    conversions_count: int = 0
    formulas: list[dict[str, Any]] | None = None
    raw_preview: list[dict[str, Any]] = field(default_factory=list)
    data_row_count: int | None = None
    warnings: list[str] = field(default_factory=list)

    def as_dict(self, *, include_rows: bool, sample_limit: int = 0) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "ok": True,
            "content_trust": CONTENT_TRUST,
            "path": str(self.path),
            "format": self.file_format,
            "sheet": self.sheet,
            "sheet_index": self.sheet_index,
            "sheet_state": self.sheet_state,
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
            "data_row_count": self.data_row_count,
            "offset": self.offset,
            "next_offset": self.next_offset,
            "continues": self.continues,
            "truncated": self.truncated,
            "declared_rows": self.declared_rows,
            "declared_columns": self.declared_cols,
            "hidden_columns": self.hidden_columns,
            "redactions": self.redactions,
            "censored_cells": self.censored_cells,
            "conversions": self.conversions,
            "conversions_count": self.conversions_count,
            "merged_ranges": self.merged_ranges,
            "formula_cells": self.formula_cells,
            "warnings": self.warnings,
        }
        if self.redaction_events:
            payload["redaction_events"] = self.redaction_events
        if self.formulas is not None:
            payload["formulas"] = self.formulas
        if self.raw_preview:
            payload["raw_preview"] = self.raw_preview
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
    infos = list_sheet_infos(resolved)
    return {
        "ok": True,
        "content_trust": CONTENT_TRUST,
        "path": str(resolved),
        "format": kind,
        "sheets": [{"name": info.name, "index": info.index, "state": info.state} for info in infos],
    }


def overview(path: str, settings: Settings) -> dict[str, Any]:
    """One short pass per sheet: state, declared size, detected tables, two samples."""
    resolved = resolve_read_file(path, settings.allowed_roots)
    check_file_size(resolved, settings.max_file_bytes)
    kind = file_format(resolved)
    infos = list_sheet_infos(resolved)
    sheets: list[dict[str, Any]] = []
    for info in infos:
        table = normalize_file(
            str(resolved),
            settings,
            sheet=info.name,
            max_rows=2,
            prefix_rows=_OVERVIEW_PREFIX_ROWS,
            normalize_units=False,
        )
        sheets.append(
            {
                "name": table.sheet,
                "index": table.sheet_index,
                "state": table.sheet_state,
                "declared_rows": table.declared_rows,
                "declared_columns": table.declared_cols,
                "hidden_columns": table.hidden_columns,
                "table_count": table.table_count,
                "tables": table.tables,
                "continues": table.continues,
                "sample_rows": table.rows[:2],
                "warnings": table.warnings,
            }
        )
    return {
        "ok": True,
        "content_trust": CONTENT_TRUST,
        "path": str(resolved),
        "format": kind,
        "sheets": sheets,
    }


def read_range(
    path: str,
    settings: Settings,
    *,
    sheet: str | None = None,
    sheet_index: int | None = None,
    cell_range: str,
    include_formulas: bool = False,
) -> dict[str, Any]:
    """Return raw cells for an A1 range. No header detection and no unit conversion."""
    _reject_both_sheet_selectors(sheet, sheet_index)
    bounds = _parse_range(cell_range)
    min_col, min_row, max_col, max_row = bounds
    row_count = max_row - min_row + 1
    col_count = max_col - min_col + 1
    warnings: list[str] = []
    if row_count > settings.max_rows:
        max_row = min_row + settings.max_rows - 1
        row_count = settings.max_rows
        warnings.append(f"Range height was capped at EXCEL_MAX_ROWS ({settings.max_rows}).")
    if col_count > settings.max_columns:
        max_col = min_col + settings.max_columns - 1
        warnings.append(f"Range width was capped at EXCEL_MAX_COLUMNS ({settings.max_columns}).")
    resolved = resolve_read_file(path, settings.allowed_roots)
    check_file_size(resolved, settings.max_file_bytes)
    kind = file_format(resolved)
    grid = read_window(
        resolved,
        sheet=sheet,
        sheet_index=sheet_index,
        max_columns=max_col,
        start_row=min_row,
        row_limit=row_count,
    )
    scrubbed, events, redactions = _maybe_scrub(grid.rows, settings, grid.start_row)
    rows: list[list[Any]] = []
    formulas: list[dict[str, Any]] = []
    uncached = 0
    for offset, cells in enumerate(scrubbed):
        sheet_row = grid.start_row + offset
        view: list[Any] = []
        for col in range(min_col, max_col + 1):
            text = cells[col - 1] if col - 1 < len(cells) else ""
            formula = grid.formulas.get((sheet_row, col))
            if formula and not text:
                uncached += 1
                view.append(None)
            else:
                view.append(text or None)
            if include_formulas and formula:
                formulas.append({"row": sheet_row, "column": col, "formula": formula})
        rows.append(view)
    if uncached:
        warnings.append(
            f"{uncached} formula cell(s) had no cached value and were returned as null."
        )
    warnings.extend(grid.warnings)
    if redactions:
        warnings.append(f"Redacted {redactions} sensitive value(s).")
    payload: dict[str, Any] = {
        "ok": True,
        "content_trust": CONTENT_TRUST,
        "path": str(resolved),
        "format": kind,
        "sheet": grid.name,
        "sheet_index": grid.index,
        "sheet_state": grid.state,
        "range": cell_range,
        "start_row": min_row,
        "start_column": min_col,
        "rows": rows,
        "declared_rows": grid.declared_rows,
        "declared_columns": grid.declared_cols,
        "hidden_columns": grid.hidden_columns,
        "redactions": redactions,
        "warnings": warnings,
    }
    if events:
        payload["redaction_events"] = events
    if include_formulas:
        payload["formulas"] = formulas
    return payload


def normalize_file(
    path: str,
    settings: Settings,
    *,
    sheet: str | None = None,
    sheet_index: int | None = None,
    table_index: int | None = None,
    header_row: int | None = None,
    max_rows: int | None = None,
    offset: int = 0,
    normalize_units: bool = False,
    include_formulas: bool = False,
    decimal: str = "auto",
    prefix_rows: int | None = None,
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
    if offset < 0:
        raise ToolFailure("bad_argument", "offset must be >= 0.")
    _reject_both_sheet_selectors(sheet, sheet_index)
    row_cap = settings.max_rows if max_rows is None else min(max_rows, settings.max_rows)
    if row_cap < 1:
        raise ToolFailure("bad_argument", "max_rows must be >= 1.")

    resolved = resolve_read_file(path, settings.allowed_roots)
    check_file_size(resolved, settings.max_file_bytes)
    kind = file_format(resolved)
    prefix = DETECT_PREFIX_ROWS if prefix_rows is None else max(1, prefix_rows)
    if header_row is not None:
        prefix = max(prefix, header_row + row_cap)
    prefix = min(prefix, settings.max_scan_rows)
    grid = read_window(
        resolved,
        sheet=sheet,
        sheet_index=sheet_index,
        max_columns=settings.max_columns,
        start_row=1,
        row_limit=prefix,
    )
    scrubbed, redaction_events, redactions = _maybe_scrub(grid.rows, settings, grid.start_row)
    hinted = _template_header(resolved, grid, scrubbed)
    if header_row is None and hinted is not None:
        header_row = hinted
    if header_row is not None:
        tables, metadata_rows, selected = _layout_from_header(scrubbed, header_row, grid.start_row)
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
    page_rows, source_rows, next_offset, continues = _select_page(
        resolved,
        settings,
        grid,
        scrubbed,
        tables[selected] if selected is not None else None,
        offset=offset,
        row_cap=row_cap,
        sheet=sheet,
        sheet_index=sheet_index,
    )
    return _materialize(
        resolved=resolved,
        kind=kind,
        grid=grid,
        scrubbed=scrubbed,
        page_rows=page_rows,
        page_source_rows=source_rows,
        tables=tables,
        metadata_rows=metadata_rows,
        selected=selected,
        redactions=redactions,
        redaction_events=redaction_events,
        normalize_units=normalize_units,
        include_formulas=include_formulas,
        decimal=decimal,
        units=settings.units,
        offset=offset,
        next_offset=next_offset,
        continues=continues,
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
    offset: int = 0,
    normalize_units: bool = False,
    include_formulas: bool = False,
    decimal: str = "auto",
) -> dict[str, Any]:
    if not settings.allow_write:
        raise ToolFailure(
            "write_disabled",
            "Export is off. Set EXCEL_ALLOW_WRITE=1 and EXCEL_OUTPUT_ROOTS "
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
        offset=offset,
        normalize_units=normalize_units,
        include_formulas=include_formulas,
        decimal=decimal,
    )
    target = resolve_write_file(output_path, settings.output_roots, overwrite=overwrite)
    if target.resolve() == table.path.resolve():
        raise ToolFailure("invalid_path", "Refusing to overwrite the source file.")
    if target.suffix.lower() == ".json":
        _atomic_write(target, _json_body(table))
    else:
        _atomic_write(target, _csv_body(table))
    return {
        "ok": True,
        "content_trust": CONTENT_TRUST,
        "output_path": str(target),
        "rows_written": len(table.rows),
        "columns": [_csv_header(column) for column in table.columns],
        "warnings": table.warnings,
    }


def _template_header(path: Path, grid: SheetGrid, rows: list[list[str]]) -> int | None:
    preview: list[dict[str, Any]] = []
    for index, row in enumerate(rows[:15]):
        text = " | ".join(cell.strip() for cell in row if cell and cell.strip())
        if text:
            preview.append({"row": grid.start_row + index, "text": text})
    view = SheetView(
        path=str(path),
        sheet=grid.name,
        sheet_index=grid.index,
        preview_rows=tuple(preview),
    )
    return get_plugin_host().template_header_row(view)


def _select_page(
    path: Path,
    settings: Settings,
    grid: SheetGrid,
    scrubbed: list[list[str]],
    table: DetectedTable | None,
    *,
    offset: int,
    row_cap: int,
    sheet: str | None,
    sheet_index: int | None,
) -> tuple[list[list[str]], list[int], int | None, bool]:
    if table is None or not table.data_rows:
        return [], [], None, False
    continues = _table_continues(grid, table)
    selected: list[list[str]] = []
    source_rows: list[int] = []
    for pos, index in enumerate(table.data_rows):
        if pos < offset:
            continue
        if len(selected) >= row_cap:
            break
        selected.append(scrubbed[index])
        source_rows.append(grid.start_row + index)
    more_in_prefix = offset + len(selected) < len(table.data_rows)
    if len(selected) >= row_cap:
        more = more_in_prefix or continues
        return selected, source_rows, (offset + row_cap if more else None), continues
    if not continues:
        return selected, source_rows, None, False
    skip = max(0, offset - len(table.data_rows))
    need = row_cap - len(selected)
    tail_rows, tail_sources, tail_more = _read_tail(
        path,
        settings,
        sheet=sheet,
        sheet_index=sheet_index,
        start_row=grid.start_row + len(grid.rows),
        skip=skip,
        need=need,
    )
    if not tail_rows and not tail_more and skip == 0:
        # The next sheet row was blank, so the prefix already held the table.
        return selected, source_rows, None, False
    selected.extend(tail_rows)
    source_rows.extend(tail_sources)
    next_offset = offset + len(selected) if tail_more else None
    return selected, source_rows, next_offset, bool(tail_rows or tail_more)


def _table_continues(grid: SheetGrid, table: DetectedTable) -> bool:
    if not grid.truncated_rows or not table.data_rows:
        return False
    return table.data_rows[-1] >= len(grid.rows) - 1


def _read_tail(
    path: Path,
    settings: Settings,
    *,
    sheet: str | None,
    sheet_index: int | None,
    start_row: int,
    skip: int,
    need: int,
) -> tuple[list[list[str]], list[int], bool]:
    """Stream data rows past the detection prefix. Skipped rows are not kept."""
    kept: list[list[str]] = []
    sources: list[int] = []
    row = start_row
    ceiling = settings.max_scan_rows
    extra = False
    while row <= ceiling and (skip > 0 or len(kept) < need):
        chunk_limit = min(500, ceiling - row + 1)
        if chunk_limit < 1:
            break
        chunk = read_window(
            path,
            sheet=sheet,
            sheet_index=sheet_index,
            max_columns=settings.max_columns,
            start_row=row,
            row_limit=chunk_limit,
        )
        if not chunk.rows:
            break
        scrubbed, _events, _count = _maybe_scrub(chunk.rows, settings, chunk.start_row)
        stop = False
        for offset, cells in enumerate(scrubbed):
            if _ends_table(cells):
                stop = True
                break
            if skip > 0:
                skip -= 1
                continue
            if len(kept) < need:
                kept.append(cells)
                sources.append(chunk.start_row + offset)
                continue
            extra = True
            stop = True
            break
        if stop or not chunk.truncated_rows:
            break
        row = chunk.start_row + len(chunk.rows)
    else:
        if row > ceiling and (skip > 0 or len(kept) >= need):
            extra = False
    return kept, sources, extra


def _ends_table(cells: list[str]) -> bool:
    if not any(cell.strip() for cell in cells):
        return True
    if any(looks_numeric(cell) for cell in cells):
        return False
    text = " ".join(cell.strip() for cell in cells if cell.strip())
    return is_footer_text(text)


def _materialize(
    *,
    resolved: Path,
    kind: str,
    grid: SheetGrid,
    scrubbed: list[list[str]],
    page_rows: list[list[str]],
    page_source_rows: list[int],
    tables: list[DetectedTable],
    metadata_rows: list[int],
    selected: int | None,
    redactions: int,
    redaction_events: list[dict[str, Any]],
    normalize_units: bool,
    include_formulas: bool,
    decimal: str,
    units: dict[str, UnitSpec],
    offset: int,
    next_offset: int | None,
    continues: bool,
) -> NormalizedTable:
    warnings: list[str] = list(grid.warnings)
    if grid.truncated_cols:
        warnings.append("Stopped scanning extra columns at EXCEL_MAX_COLUMNS.")
    if redactions:
        warnings.append(f"Redacted {redactions} sensitive value(s).")
    if grid.declared_rows and grid.declared_rows > DETECT_PREFIX_ROWS and not grid.rows:
        warnings.append(
            "The used range extends past the detection prefix and no table was found. "
            "Use excel_read_range with an explicit range."
        )

    summaries = [
        _summary(scrubbed, table, index, grid.start_row) for index, table in enumerate(tables)
    ]
    metadata = [_row_record(scrubbed, index, grid.start_row) for index in metadata_rows]
    if selected is None:
        warnings.append(
            "No data table detected in the scanned prefix. Pass header_row or use excel_read_range."
        )
        return NormalizedTable(
            path=resolved,
            file_format=kind,
            sheet=grid.name,
            sheet_index=grid.index,
            sheet_state=grid.state,
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
            redaction_events=redaction_events,
            merged_ranges=grid.merged_ranges,
            formula_cells=0,
            offset=offset,
            next_offset=None,
            continues=False,
            declared_rows=grid.declared_rows,
            declared_cols=grid.declared_cols,
            hidden_columns=list(grid.hidden_columns),
            raw_preview=_raw_preview(scrubbed, grid.start_row),
            warnings=warnings,
        )

    table = tables[selected]
    if len(tables) > 1:
        warnings.append(
            f"Sheet contains {len(tables)} tables; returned table_index {selected}. "
            "Pass table_index to select another."
        )
    columns, kept = _build_columns(
        scrubbed, table, units, warnings, normalize_units, grid.start_row
    )
    style = _decimal_style(scrubbed, table, kept, decimal)
    page_formulas = _formula_payload(grid, page_source_rows, columns)
    rows, censored, conversions, conversion_count, formula_hits = _build_rows(
        page_rows,
        page_source_rows,
        columns,
        kept,
        decimal=style,
        normalize_units=normalize_units,
        formulas=grid.formulas,
        include_formulas=include_formulas,
    )
    if formula_hits:
        warnings.append(
            f"{formula_hits} formula cell(s) had no cached value and were returned as null."
        )
    if next_offset is not None:
        warnings.append(
            f"Returned data rows {offset}..{offset + len(rows) - 1}. "
            f"Pass offset={next_offset} for the next page."
        )
    notes = [
        _row_record(scrubbed, index, grid.start_row)
        for index in table.note_rows
        if _row_text(scrubbed, index)
    ]
    preamble = [_row_record(scrubbed, index, grid.start_row) for index in table.preamble_rows]
    data_row_count = None if continues else len(table.data_rows)
    return NormalizedTable(
        path=resolved,
        file_format=kind,
        sheet=grid.name,
        sheet_index=grid.index,
        sheet_state=grid.state,
        table_index=selected,
        table_count=len(tables),
        tables=summaries,
        encoding=grid.encoding,
        decimal=style,
        metadata=[*metadata, *preamble],
        header_rows=[grid.start_row + index for index in table.header_rows],
        data_start_row=(grid.start_row + table.data_rows[0]) if table.data_rows else None,
        data_end_row=page_source_rows[-1] if page_source_rows else None,
        notes=notes,
        columns=columns,
        rows=rows,
        source_rows=page_source_rows,
        truncated=bool(grid.truncated_rows or grid.truncated_cols or next_offset is not None),
        redactions=redactions,
        redaction_events=redaction_events,
        merged_ranges=grid.merged_ranges,
        formula_cells=len(page_formulas),
        offset=offset,
        next_offset=next_offset,
        continues=continues,
        declared_rows=grid.declared_rows,
        declared_cols=grid.declared_cols,
        hidden_columns=list(grid.hidden_columns),
        censored_cells=censored,
        conversions=conversions,
        conversions_count=conversion_count,
        formulas=page_formulas if include_formulas else None,
        data_row_count=data_row_count,
        warnings=warnings,
    )


def _formula_payload(
    grid: SheetGrid, source_rows: list[int], columns: list[ColumnView]
) -> list[dict[str, Any]]:
    names = {column.source_column: column.name for column in columns}
    wanted = set(source_rows)
    payload: list[dict[str, Any]] = []
    for (sheet_row, column), formula in sorted(grid.formulas.items()):
        if sheet_row not in wanted:
            continue
        payload.append(
            {
                "row": sheet_row,
                "column": names.get(column, column),
                "formula": formula,
            }
        )
    return payload


def _layout_from_header(
    grid: list[list[str]], header_row: int, start_row: int
) -> tuple[list[DetectedTable], list[int], int]:
    index = header_row - start_row
    if index < 0 or index >= len(grid) or not any(cell.strip() for cell in grid[index]):
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
    normalize_units: bool,
    start_row: int,
) -> tuple[list[ColumnView], list[int]]:
    del start_row
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
        same_label = bool(
            spec and source_unit and unit_key(source_unit) == unit_key(spec.canonical)
        )
        if spec and (normalize_units or same_label):
            display = spec.canonical
        else:
            display = source_unit
        converted = bool(
            normalize_units
            and spec
            and (not same_label or spec.factor != 1.0 or spec.offset != 0.0)
        )
        views.append(
            ColumnView(
                index=len(views),
                source_column=column + 1,
                name=name,
                source_header=source_header,
                unit=display,
                source_unit=source_unit,
                canonical_unit=spec.canonical if spec else None,
                dimension=spec.dimension if spec else None,
                unit_converted=converted,
                unit_note=spec.note if spec else None,
                spec=spec,
            )
        )
    return views, kept


def _build_rows(
    page_rows: list[list[str]],
    source_rows: list[int],
    columns: list[ColumnView],
    kept: list[int],
    *,
    decimal: DecimalStyle,
    normalize_units: bool,
    formulas: dict[tuple[int, int], str],
    include_formulas: bool,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], int, int]:
    del include_formulas
    rows: list[dict[str, Any]] = []
    censored: list[dict[str, Any]] = []
    conversions: list[dict[str, Any]] = []
    conversion_count = 0
    formula_hits = 0
    for cells, sheet_row in zip(page_rows, source_rows, strict=True):
        record: dict[str, Any] = {}
        for column, source_column in zip(columns, kept, strict=True):
            text = cells[source_column].strip() if source_column < len(cells) else ""
            formula = formulas.get((sheet_row, source_column + 1))
            value, kind = _cell_value(
                text,
                decimal=decimal,
                normalize_units=normalize_units,
                column=column,
                formula=formula,
            )
            if kind == "formula":
                formula_hits += 1
            elif kind == "censored":
                censored.append(
                    {
                        "row": sheet_row,
                        "column": column.name,
                        "qualifier": text,
                    }
                )
            elif kind == "converted":
                conversion_count += 1
                if len(conversions) < _CONVERSION_LIST_CAP:
                    conversions.append({"row": sheet_row, "column": column.name})
            record[column.name] = value
        rows.append(record)
    return rows, censored, conversions, conversion_count, formula_hits


def _cell_value(
    text: str,
    *,
    decimal: DecimalStyle,
    normalize_units: bool,
    column: ColumnView,
    formula: str | None,
) -> tuple[Any, str]:
    if formula and (not text or text == formula):
        return None, "formula"
    parsed = parse_number(text, decimal)
    if parsed.kind in {"empty", "null"}:
        return None, "empty"
    if parsed.kind == "censored":
        return parsed.text if parsed.text is not None else text, "censored"
    if parsed.kind == "text" or parsed.number is None:
        return parsed.text if parsed.kind == "text" else text, "text"
    value: int | float = parsed.number
    spec = column.spec
    kind = "number"
    if normalize_units and spec is not None and (spec.factor != 1.0 or spec.offset != 0.0):
        value = spec.convert(float(value))
        kind = "converted"
    if isinstance(value, float) and value.is_integer() and abs(value) < 10**15:
        value = int(value)
    if parsed.percent and column.unit is None:
        column.unit = "%"
        column.dimension = column.dimension or "fraction"
    return value, kind


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


def _summary(
    grid: list[list[str]], table: DetectedTable, index: int, start_row: int
) -> dict[str, Any]:
    width = 0
    if table.header_rows or table.data_rows:
        width = max(len(grid[row]) for row in [*table.header_rows, *table.data_rows])
    return {
        "index": index,
        "header_rows": [start_row + row for row in table.header_rows],
        "data_start_row": (start_row + table.data_rows[0]) if table.data_rows else None,
        "data_end_row": (start_row + table.data_rows[-1]) if table.data_rows else None,
        "data_rows": len(table.data_rows),
        "columns": width,
    }


def _maybe_scrub(
    rows: list[list[str]], settings: Settings, start_row: int
) -> tuple[list[list[str]], list[dict[str, Any]], int]:
    if settings.redact != "secrets":
        return rows, [], 0
    events: list[dict[str, Any]] = []
    total = 0
    scrubbed: list[list[str]] = []
    for row_offset, row in enumerate(rows):
        cleaned: list[str] = []
        for col_offset, cell in enumerate(row):
            text, kinds = scrub_text(cell)
            total += len(kinds)
            for kind in kinds:
                if len(events) < _REDACTION_LIST_CAP:
                    events.append(
                        {
                            "row": start_row + row_offset,
                            "column": col_offset + 1,
                            "kind": kind,
                        }
                    )
            cleaned.append(text)
        scrubbed.append(cleaned)
    return scrubbed, events, total


def _raw_preview(grid: list[list[str]], start_row: int) -> list[dict[str, Any]]:
    preview: list[dict[str, Any]] = []
    for index, row in enumerate(grid):
        text = " | ".join(cell.strip() for cell in row if cell and cell.strip())
        if not text:
            continue
        preview.append({"row": start_row + index, "text": text})
        if len(preview) >= 5:
            break
    return preview


def _row_record(grid: list[list[str]], index: int, start_row: int) -> dict[str, Any]:
    return {"row": start_row + index, "text": _row_text(grid, index)}


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


def _csv_body(table: NormalizedTable) -> str:
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\n")
    writer.writerow([_csv_safe(_csv_header(column)) for column in table.columns])
    for row in table.rows:
        writer.writerow([_csv_safe(_csv_cell(row.get(column.name))) for column in table.columns])
    return buffer.getvalue()


def _json_body(table: NormalizedTable) -> str:
    payload = {
        "content_trust": CONTENT_TRUST,
        "columns": [column.as_dict() for column in table.columns],
        "rows": table.rows,
        "source_rows": table.source_rows,
        "warnings": table.warnings,
    }
    return json.dumps(payload, ensure_ascii=True, indent=2) + "\n"


def _atomic_write(path: Path, text: str) -> None:
    if path.is_symlink():
        raise ToolFailure("invalid_path", "Refusing to write through a symlink.")
    fd, tmp_name = tempfile.mkstemp(prefix=".excel-write-", suffix=".tmp", dir=str(path.parent))
    tmp = Path(tmp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        if path.is_symlink():
            raise ToolFailure("invalid_path", "Refusing to write through a symlink.")
        os.replace(tmp, path)
    except Exception:
        if tmp.exists():
            tmp.unlink()
        raise


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
    """Neutralize spreadsheet formula injection, including a leading minus."""
    stripped = value.lstrip(" \t")
    if stripped and (stripped[0] in _FORMULA_LEAD or stripped[0] == _FULLWIDTH_EQUALS):
        return "'" + value
    return value


def _reject_both_sheet_selectors(sheet: str | None, sheet_index: int | None) -> None:
    has_sheet = sheet is not None and str(sheet).strip() != ""
    if has_sheet and sheet_index is not None:
        raise ToolFailure("bad_argument", "Pass sheet or sheet_index, not both.")


def _parse_range(cell_range: str) -> tuple[int, int, int, int]:
    text = (cell_range or "").strip()
    if not text:
        raise ToolFailure("bad_argument", "range is required, for example A1:F50.")
    try:
        min_col, min_row, max_col, max_row = range_boundaries(text)
    except ValueError as exc:
        raise ToolFailure("bad_argument", "range must be an A1 reference such as A1:F50.") from exc
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise ToolFailure("bad_argument", "range must be an A1 reference such as A1:F50.")
    return min_col, min_row, max_col, max_row
