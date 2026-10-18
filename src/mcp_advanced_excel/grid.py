"""Load xlsx/xlsm/xls/csv into a rectangular string grid.

Formulas are kept as text. Macros are not loaded (``keep_vba=False``) and are
never executed. Merged cells are filled from the top-left cell, which is the
same idea as reading a merged header in AdvancedExcel after openpyxl expands
the cached value across the range.
"""

from __future__ import annotations

import csv
import io
import math
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from .errors import ToolFailure

_SUPPORTED = {".xlsx": "xlsx", ".xlsm": "xlsm", ".xls": "xls", ".csv": "csv", ".tsv": "tsv"}
_WHITESPACE = re.compile(r"[\t\r\n\f\v]+")


@dataclass
class SheetGrid:
    name: str
    index: int
    rows: list[list[str]]
    encoding: str | None = None
    formula_cells: int = 0
    merged_ranges: int = 0
    truncated_rows: bool = False
    truncated_cols: bool = False
    warnings: list[str] = field(default_factory=list)


def file_format(path: Path) -> str:
    kind = _SUPPORTED.get(path.suffix.lower())
    if kind is None:
        raise ToolFailure(
            "unsupported_type",
            "Supported files are .xlsx, .xlsm, .xls, .csv, and .tsv.",
        )
    return kind


def check_file_size(path: Path, max_bytes: int) -> None:
    try:
        size = path.stat().st_size
    except OSError as exc:
        raise ToolFailure("invalid_path", "File size could not be read.") from exc
    if size > max_bytes:
        raise ToolFailure(
            "file_too_large",
            f"File is {size} bytes; the limit is {max_bytes} bytes.",
        )


def list_sheet_names(path: Path) -> list[str]:
    kind = file_format(path)
    if kind in {"csv", "tsv"}:
        return [path.stem or "csv"]
    if kind == "xls":
        return _xls_sheet_names(path)
    return _xlsx_sheet_names(path)


def read_sheet(
    path: Path,
    *,
    sheet: str | None,
    sheet_index: int | None,
    max_columns: int,
    max_scan_rows: int,
) -> SheetGrid:
    kind = file_format(path)
    if kind in {"csv", "tsv"}:
        return _read_csv(path, kind=kind, max_columns=max_columns, max_scan_rows=max_scan_rows)
    if kind == "xls":
        return _read_xls(
            path,
            sheet=sheet,
            sheet_index=sheet_index,
            max_columns=max_columns,
            max_scan_rows=max_scan_rows,
        )
    return _read_xlsx(
        path,
        sheet=sheet,
        sheet_index=sheet_index,
        max_columns=max_columns,
        max_scan_rows=max_scan_rows,
    )


def tidy_cell(value: object) -> tuple[str, bool]:
    """Return display text and whether the cell was an unevaluated formula."""
    if value is None:
        return "", False
    if isinstance(value, bool):
        return "true" if value else "false", False
    if isinstance(value, datetime):
        return value.isoformat(sep=" "), False
    if isinstance(value, date):
        return value.isoformat(), False
    if isinstance(value, float):
        if not math.isfinite(value):
            return "", False
        if value.is_integer() and abs(value) < 10**15:
            return str(int(value)), False
        return format(value, ".12g"), False
    if isinstance(value, int):
        return str(value), False
    text = str(value)
    formula = text.startswith("=")
    text = _WHITESPACE.sub(" ", text).strip()
    return text, formula


def _xlsx_sheet_names(path: Path) -> list[str]:
    try:
        workbook = load_workbook(
            path, read_only=True, data_only=False, keep_vba=False, keep_links=False
        )
    except (InvalidFileException, zipfile.BadZipFile, OSError, ValueError) as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def _xls_sheet_names(path: Path) -> list[str]:
    book = _open_xls(path, formatting_info=False)
    try:
        return list(book.sheet_names())
    finally:
        book.release_resources()


def _read_xlsx(
    path: Path,
    *,
    sheet: str | None,
    sheet_index: int | None,
    max_columns: int,
    max_scan_rows: int,
) -> SheetGrid:
    try:
        workbook = load_workbook(
            path, read_only=False, data_only=False, keep_vba=False, keep_links=False
        )
    except (InvalidFileException, zipfile.BadZipFile, OSError, ValueError) as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc
    try:
        worksheet, index = _select_openpyxl(workbook, sheet, sheet_index)
        declared_rows = int(worksheet.max_row or 0)
        declared_cols = int(worksheet.max_column or 0)
        scan_rows = min(declared_rows, max_scan_rows) if declared_rows else 0
        scan_cols = min(declared_cols, max_columns) if declared_cols else 0
        grid = [["" for _ in range(scan_cols)] for _ in range(scan_rows)]
        formulas = 0
        if scan_rows and scan_cols:
            for cells in worksheet.iter_rows(min_row=1, max_row=scan_rows, max_col=scan_cols):
                for cell in cells:
                    text, formula = tidy_cell(cell.value)
                    grid[cell.row - 1][cell.column - 1] = text
                    if formula:
                        formulas += 1
        merged = _fill_openpyxl_merges(worksheet, grid, scan_rows, scan_cols)
        return SheetGrid(
            name=worksheet.title,
            index=index,
            rows=_trim_trailing_blank_rows(grid),
            formula_cells=formulas,
            merged_ranges=merged,
            truncated_rows=declared_rows > scan_rows,
            truncated_cols=declared_cols > scan_cols,
        )
    finally:
        workbook.close()


def _open_xls(path: Path, *, formatting_info: bool):
    import xlrd

    try:
        return xlrd.open_workbook(
            str(path),
            on_demand=True,
            formatting_info=formatting_info,
            logfile=io.StringIO(),
        )
    except Exception as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc


def _read_xls(
    path: Path,
    *,
    sheet: str | None,
    sheet_index: int | None,
    max_columns: int,
    max_scan_rows: int,
) -> SheetGrid:
    from xlrd import XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_ERROR, XL_CELL_NUMBER, xldate

    try:
        book = _open_xls(path, formatting_info=True)
    except ToolFailure:
        book = _open_xls(path, formatting_info=False)
    try:
        worksheet, index = _select_xlrd(book, sheet, sheet_index)
        declared_rows = int(worksheet.nrows)
        declared_cols = int(worksheet.ncols)
        scan_rows = min(declared_rows, max_scan_rows)
        scan_cols = min(declared_cols, max_columns)
        grid = [["" for _ in range(scan_cols)] for _ in range(scan_rows)]
        for row_index in range(scan_rows):
            for col_index in range(scan_cols):
                cell = worksheet.cell(row_index, col_index)
                value: object
                if cell.ctype == XL_CELL_DATE:
                    try:
                        value = xldate.xldate_as_datetime(cell.value, book.datemode)
                    except Exception:
                        value = cell.value
                elif cell.ctype == XL_CELL_NUMBER:
                    value = cell.value
                elif cell.ctype == XL_CELL_BOOLEAN:
                    value = bool(cell.value)
                elif cell.ctype == XL_CELL_ERROR:
                    value = ""
                else:
                    value = cell.value
                text, _formula = tidy_cell(value)
                grid[row_index][col_index] = text
        merged = 0
        for rlo, rhi, clo, chi in getattr(worksheet, "merged_cells", []) or []:
            if rlo >= scan_rows or clo >= scan_cols:
                continue
            value = grid[rlo][clo]
            merged += 1
            for row_index in range(rlo, min(rhi, scan_rows)):
                for col_index in range(clo, min(chi, scan_cols)):
                    if not grid[row_index][col_index]:
                        grid[row_index][col_index] = value
        return SheetGrid(
            name=worksheet.name,
            index=index,
            rows=_trim_trailing_blank_rows(grid),
            merged_ranges=merged,
            truncated_rows=declared_rows > scan_rows,
            truncated_cols=declared_cols > scan_cols,
        )
    finally:
        book.release_resources()


def _read_csv(
    path: Path,
    *,
    kind: str,
    max_columns: int,
    max_scan_rows: int,
) -> SheetGrid:
    try:
        raw = path.read_bytes()
    except OSError as exc:
        raise ToolFailure("read_failed", "File could not be read.") from exc
    text, encoding = _decode(raw)
    if kind == "tsv":
        delimiter = "\t"
    else:
        delimiter = _sniff_delimiter(text)
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows: list[list[str]] = []
    truncated_rows = False
    truncated_cols = False
    try:
        for row in reader:
            if len(rows) >= max_scan_rows:
                truncated_rows = True
                break
            cleaned: list[str] = []
            for index, cell in enumerate(row):
                if index >= max_columns:
                    truncated_cols = True
                    break
                cleaned.append(tidy_cell(cell)[0])
            rows.append(cleaned)
    except csv.Error as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc
    width = max((len(row) for row in rows), default=0)
    rectangular = [row + [""] * (width - len(row)) for row in rows]
    return SheetGrid(
        name=path.stem or "csv",
        index=0,
        rows=_trim_trailing_blank_rows(rectangular),
        encoding=encoding,
        truncated_rows=truncated_rows,
        truncated_cols=truncated_cols,
    )


def _decode(raw: bytes) -> tuple[str, str]:
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16"), "utf-16"
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "utf-8-sig"
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1"), "latin-1"


def _sniff_delimiter(text: str) -> str:
    """Pick a delimiter. European lab exports often use ';' because ',' is the decimal mark."""
    voted = _vote_delimiter(text)
    sniffed = None
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|")
        if dialect.delimiter in {",", ";", "\t", "|"}:
            sniffed = dialect.delimiter
    except csv.Error:
        sniffed = None
    if sniffed is None:
        return voted
    # A comma sniffer hit on ``0,125`` must lose to a semicolon that builds
    # a wider, stable table.
    if _typical_width(text, sniffed) >= _typical_width(text, voted):
        return sniffed
    return voted


def _vote_delimiter(text: str) -> str:
    lines = [line for line in text.splitlines() if line.strip()][:30]
    best = ","
    best_score = -1
    for delimiter in (",", ";", "\t", "|"):
        counts = [line.count(delimiter) for line in lines if delimiter in line]
        if not counts:
            continue
        common = max(set(counts), key=counts.count)
        if common < 1:
            continue
        stable = sum(1 for count in counts if count == common)
        score = stable * 10 + common
        if score > best_score:
            best_score = score
            best = delimiter
    return best


def _typical_width(text: str, delimiter: str) -> int:
    widths: list[int] = []
    try:
        reader = csv.reader(io.StringIO(text[:8192]), delimiter=delimiter)
        for index, row in enumerate(reader):
            if index >= 20:
                break
            widths.append(sum(1 for cell in row if cell.strip()))
    except csv.Error:
        return 0
    nonempty = [width for width in widths if width]
    if not nonempty:
        return 0
    return max(nonempty)


def _select_openpyxl(workbook, sheet: str | None, sheet_index: int | None):
    names = list(workbook.sheetnames)
    index = _resolve_index(names, sheet, sheet_index)
    return workbook[names[index]], index


def _select_xlrd(book, sheet: str | None, sheet_index: int | None):
    names = list(book.sheet_names())
    index = _resolve_index(names, sheet, sheet_index)
    return book.sheet_by_index(index), index


def _resolve_index(names: list[str], sheet: str | None, sheet_index: int | None) -> int:
    if sheet is not None and str(sheet).strip() != "":
        wanted = str(sheet)
        if wanted not in names:
            raise ToolFailure("sheet_not_found", f"Sheet {wanted!r} was not found.")
        return names.index(wanted)
    if sheet_index is None:
        if not names:
            raise ToolFailure("sheet_not_found", "Workbook has no sheets.")
        return 0
    if sheet_index < 0 or sheet_index >= len(names):
        raise ToolFailure(
            "sheet_not_found",
            f"sheet_index {sheet_index} is outside 0..{len(names) - 1}.",
        )
    return sheet_index


def _fill_openpyxl_merges(worksheet, grid: list[list[str]], scan_rows: int, scan_cols: int) -> int:
    count = 0
    ranges = getattr(worksheet, "merged_cells", None)
    if ranges is None:
        return 0
    for merged in ranges.ranges:
        min_col, min_row, max_col, max_row = merged.bounds
        if min_row > scan_rows or min_col > scan_cols:
            continue
        count += 1
        value = grid[min_row - 1][min_col - 1]
        for row_index in range(min_row - 1, min(max_row, scan_rows)):
            for col_index in range(min_col - 1, min(max_col, scan_cols)):
                if not grid[row_index][col_index]:
                    grid[row_index][col_index] = value
    return count


def _trim_trailing_blank_rows(grid: list[list[str]]) -> list[list[str]]:
    end = len(grid)
    while end > 0 and not any(cell.strip() for cell in grid[end - 1]):
        end -= 1
    return grid[:end]


def _short_read_error(exc: Exception) -> str:
    detail = str(exc).strip().splitlines()[0] if str(exc).strip() else exc.__class__.__name__
    if len(detail) > 180:
        detail = detail[:180] + "..."
    return f"Could not read workbook ({exc.__class__.__name__}: {detail})."
