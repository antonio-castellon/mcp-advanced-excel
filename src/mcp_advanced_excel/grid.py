"""Load xlsx/xlsm/xls/csv into a small window of cells.

Workbooks are streamed. The reader does not allocate a rectangle the size of
``dimension``. xlsx cells come from ``read_only`` worksheets. Cached formula
values are preferred; the formula text is recorded separately and is not
treated as the cell value. Macros are not loaded (``keep_vba=False``) and are
never executed.

Merged cells are filled from the top-left cell when the sheet part is small
enough to scan. Very large sheets skip that scan and say so.
"""

from __future__ import annotations

import csv
import io
import math
import re
import zipfile
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from xml.etree.ElementTree import Element

import defusedxml.ElementTree as defused_et
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format
from openpyxl.utils import range_boundaries
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import from_excel

from .errors import ToolFailure

_SUPPORTED = {".xlsx": "xlsx", ".xlsm": "xlsm", ".xls": "xls", ".csv": "csv", ".tsv": "tsv"}
_WHITESPACE = re.compile(r"[\t\r\n\f\v]+")

# A 6 MB compressed lab export can inflate to about 100 MB and still be a real
# sheet. Anything past these limits is refused before openpyxl parses it.
MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024
MAX_COMPRESSION_RATIO = 100
# Merge records sit after the row data. Skip that scan on huge sheet parts.
_MAX_MERGE_SCAN_BYTES = 8 * 1024 * 1024
DETECT_PREFIX_ROWS = 2_000

_CACHE: OrderedDict[tuple, object] = OrderedDict()
_CACHE_LIMIT = 8


@dataclass
class SheetInfo:
    name: str
    index: int
    state: str = "visible"
    target: str = ""


@dataclass
class SheetGrid:
    name: str
    index: int
    rows: list[list[str]]
    start_row: int = 1
    encoding: str | None = None
    state: str = "visible"
    formula_cells: int = 0
    uncached_formulas: int = 0
    # (sheet_row, column) 1-based -> formula text. Not copied into cell values.
    formulas: dict[tuple[int, int], str] = field(default_factory=dict)
    merged_ranges: int = 0
    truncated_rows: bool = False
    truncated_cols: bool = False
    declared_rows: int | None = None
    declared_cols: int | None = None
    hidden_columns: list[int] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def clear_read_cache() -> None:
    _CACHE.clear()


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


def enforce_zip_limits(compressed: int, uncompressed: int) -> None:
    """Refuse a zip whose declared sizes look like a compression bomb."""
    if uncompressed < 0 or compressed < 0:
        raise ToolFailure("read_failed", "Zip metadata is invalid.")
    if uncompressed > MAX_UNCOMPRESSED_BYTES:
        raise ToolFailure(
            "zip_bomb",
            f"Uncompressed workbook is {uncompressed} bytes; "
            f"the limit is {MAX_UNCOMPRESSED_BYTES} bytes.",
        )
    baseline = max(compressed, 1)
    if uncompressed / baseline > MAX_COMPRESSION_RATIO:
        raise ToolFailure(
            "zip_bomb",
            "Workbook compression ratio exceeds the limit.",
        )


def guard_office_zip(path: Path) -> None:
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            uncompressed = sum(info.file_size for info in infos)
            compressed = sum(info.compress_size for info in infos)
    except zipfile.BadZipFile as exc:
        raise ToolFailure("read_failed", "File is not a valid xlsx zip archive.") from exc
    except OSError as exc:
        raise ToolFailure("read_failed", "File could not be read.") from exc
    enforce_zip_limits(compressed, uncompressed)


def list_sheet_infos(path: Path) -> list[SheetInfo]:
    kind = file_format(path)
    if kind in {"csv", "tsv"}:
        return [SheetInfo(name=path.stem or "csv", index=0, state="visible")]
    if kind == "xls":
        names = _xls_sheet_names(path)
        return [
            SheetInfo(name=name, index=index, state="visible") for index, name in enumerate(names)
        ]
    guard_office_zip(path)
    return _xlsx_sheet_infos(path)


def list_sheet_names(path: Path) -> list[str]:
    return [info.name for info in list_sheet_infos(path)]


def read_window(
    path: Path,
    *,
    sheet: str | None,
    sheet_index: int | None,
    max_columns: int,
    start_row: int = 1,
    row_limit: int = DETECT_PREFIX_ROWS,
) -> SheetGrid:
    """Stream ``row_limit`` sheet rows starting at ``start_row`` (1-based)."""
    if start_row < 1:
        raise ToolFailure("bad_argument", "start_row is a 1-based sheet row.")
    if row_limit < 1:
        raise ToolFailure("bad_argument", "row_limit must be >= 1.")
    kind = file_format(path)
    if kind in {"csv", "tsv"}:
        return _read_csv(
            path,
            kind=kind,
            max_columns=max_columns,
            start_row=start_row,
            row_limit=row_limit,
        )
    if kind == "xls":
        return _read_xls(
            path,
            sheet=sheet,
            sheet_index=sheet_index,
            max_columns=max_columns,
            start_row=start_row,
            row_limit=row_limit,
        )
    return _read_xlsx(
        path,
        sheet=sheet,
        sheet_index=sheet_index,
        max_columns=max_columns,
        start_row=start_row,
        row_limit=row_limit,
    )


def read_sheet(
    path: Path,
    *,
    sheet: str | None,
    sheet_index: int | None,
    max_columns: int,
    max_scan_rows: int,
) -> SheetGrid:
    """Compatibility wrapper: the detection prefix, not the whole dimension."""
    return read_window(
        path,
        sheet=sheet,
        sheet_index=sheet_index,
        max_columns=max_columns,
        start_row=1,
        row_limit=min(max_scan_rows, DETECT_PREFIX_ROWS),
    )


def tidy_cell(value: object) -> str:
    """Display text for a cached value. Formula strings are not produced here."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0 and value.microsecond == 0:
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float):
        if not math.isfinite(value):
            return ""
        if value.is_integer() and abs(value) < 10**15:
            return str(int(value))
        return format(value, ".12g")
    if isinstance(value, int):
        return str(value)
    return _WHITESPACE.sub(" ", str(value)).strip()


def _cache_get(key: tuple) -> object | None:
    value = _CACHE.get(key)
    if value is None:
        return None
    _CACHE.move_to_end(key)
    return value


def _cache_put(key: tuple, value: object) -> None:
    _CACHE[key] = value
    _CACHE.move_to_end(key)
    while len(_CACHE) > _CACHE_LIMIT:
        _CACHE.popitem(last=False)


def _file_token(path: Path) -> tuple:
    st = path.stat()
    return (str(path), st.st_mtime_ns, st.st_size)


def _xlsx_sheet_infos(path: Path) -> list[SheetInfo]:
    key = (*_file_token(path), "sheets")
    cached = _cache_get(key)
    if isinstance(cached, list):
        return cached
    try:
        with zipfile.ZipFile(path) as archive:
            workbook = _xml_root(archive, "xl/workbook.xml")
            rels = _xml_root(archive, "xl/_rels/workbook.xml.rels")
    except (KeyError, zipfile.BadZipFile, OSError) as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc
    targets: dict[str, str] = {}
    for rel in rels:
        if _local(rel.tag) != "Relationship":
            continue
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target") or ""
        if rel_id:
            targets[rel_id] = _normalize_sheet_target(target)
    infos: list[SheetInfo] = []
    sheets_parent = None
    for element in workbook.iter():
        if _local(element.tag) == "sheets":
            sheets_parent = element
            break
    if sheets_parent is None:
        raise ToolFailure("read_failed", "Workbook has no sheets.")
    for element in sheets_parent:
        if _local(element.tag) != "sheet":
            continue
        name = element.attrib.get("name") or f"Sheet{len(infos) + 1}"
        state = element.attrib.get("state") or "visible"
        rel_id = element.attrib.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
            "",
        )
        infos.append(
            SheetInfo(name=name, index=len(infos), state=state, target=targets.get(rel_id, ""))
        )
    _cache_put(key, infos)
    return infos


def _xml_root(archive: zipfile.ZipFile, name: str) -> Element:
    with archive.open(name) as handle:
        return defused_et.fromstring(handle.read())


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _normalize_sheet_target(target: str) -> str:
    text = target.lstrip("/")
    if text.startswith("xl/"):
        return text
    return "xl/" + text


def _peek_sheet(path: Path, target: str) -> tuple[int | None, int | None, list[int]]:
    """Read dimension and hidden columns, then stop at the first row."""
    declared_rows = None
    declared_cols = None
    hidden: list[int] = []
    if not target:
        return None, None, hidden
    try:
        with zipfile.ZipFile(path) as archive, archive.open(target) as handle:
            for event, elem in defused_et.iterparse(handle, events=("start", "end")):
                tag = _local(elem.tag)
                if event == "end" and tag == "dimension":
                    declared_rows, declared_cols = _dimension(elem.attrib.get("ref"))
                elif event == "end" and tag == "col":
                    if elem.attrib.get("hidden") in {"1", "true"}:
                        hidden.extend(_col_span(elem.attrib.get("min"), elem.attrib.get("max")))
                elif event == "start" and tag == "sheetData":
                    break
                if event == "end":
                    elem.clear()
    except (KeyError, zipfile.BadZipFile, OSError) as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc
    return declared_rows, declared_cols, hidden


def _dimension(ref: str | None) -> tuple[int | None, int | None]:
    if not ref or ":" not in ref:
        if not ref:
            return None, None
        try:
            min_col, min_row, max_col, max_row = range_boundaries(ref)
        except ValueError:
            return None, None
        return max_row, max_col
    try:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
    except ValueError:
        return None, None
    return max_row, max_col


def _col_span(min_raw: str | None, max_raw: str | None) -> list[int]:
    try:
        start = int(min_raw or "0")
        end = int(max_raw or min_raw or "0")
    except ValueError:
        return []
    if start < 1 or end < start or end - start > 512:
        return []
    return list(range(start, end + 1))


def _merge_ranges(path: Path, target: str) -> tuple[list[tuple[int, int, int, int]], bool]:
    """Return merge bounds and whether the scan was skipped."""
    if not target:
        return [], False
    try:
        with zipfile.ZipFile(path) as archive:
            info = archive.getinfo(target)
            if info.file_size > _MAX_MERGE_SCAN_BYTES:
                return [], True
            merges: list[tuple[int, int, int, int]] = []
            with archive.open(target) as handle:
                for _event, elem in defused_et.iterparse(handle, events=("end",)):
                    if _local(elem.tag) == "mergeCell":
                        ref = elem.attrib.get("ref")
                        if ref:
                            try:
                                merges.append(range_boundaries(ref))
                            except ValueError:
                                pass
                    elem.clear()
            return merges, False
    except KeyError:
        return [], False
    except (zipfile.BadZipFile, OSError) as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc


def _read_xlsx(
    path: Path,
    *,
    sheet: str | None,
    sheet_index: int | None,
    max_columns: int,
    start_row: int,
    row_limit: int,
) -> SheetGrid:
    guard_office_zip(path)
    infos = _xlsx_sheet_infos(path)
    index = _resolve_index([info.name for info in infos], sheet, sheet_index)
    info = infos[index]
    key = (*_file_token(path), "window", index, start_row, row_limit, max_columns)
    cached = _cache_get(key)
    if isinstance(cached, SheetGrid):
        return cached
    declared_rows, declared_cols, hidden = _peek_sheet(path, info.target)
    grid, formulas, uncached, saw_extra_column = _stream_sheet_rows(
        path,
        target=info.target,
        start_row=start_row,
        row_limit=row_limit,
        max_columns=max_columns,
    )
    _pad(grid)
    merges, skipped = _merge_ranges(path, info.target)
    merged_count = _apply_merges(grid, start_row, merges, max_columns)
    warnings: list[str] = []
    if skipped:
        warnings.append("Merge cells were not expanded because the sheet is large.")
    trimmed = _trim_trailing_blank_rows(grid)
    # A full window of non-blank rows means later sheet rows were not stored.
    # A huge dimension alone does not: empty rows past the data are not loaded.
    truncated_rows = len(trimmed) >= row_limit
    truncated_cols = saw_extra_column or bool(declared_cols and declared_cols > max_columns)
    result = SheetGrid(
        name=info.name,
        index=index,
        rows=trimmed,
        start_row=start_row,
        state=info.state,
        formula_cells=len(formulas),
        uncached_formulas=uncached,
        formulas=formulas,
        merged_ranges=merged_count,
        truncated_rows=truncated_rows,
        truncated_cols=truncated_cols,
        declared_rows=declared_rows,
        declared_cols=declared_cols,
        hidden_columns=hidden,
        warnings=warnings,
    )
    _cache_put(key, result)
    return result


def _stream_sheet_rows(
    path: Path,
    *,
    target: str,
    start_row: int,
    row_limit: int,
    max_columns: int,
) -> tuple[list[list[str]], dict[tuple[int, int], str], int, bool]:
    """Stream one sheet and stop once the window is past. One pass, both values and formulas."""
    end_row = start_row + row_limit - 1
    strings, date_styles = _xlsx_support(path)
    grid: list[list[str]] = []
    formulas: dict[tuple[int, int], str] = {}
    uncached = 0
    extra_column = False
    try:
        with zipfile.ZipFile(path) as archive, archive.open(target) as handle:
            for _event, elem in defused_et.iterparse(handle, events=("end",)):
                if _local(elem.tag) != "row":
                    continue
                sheet_row = int(elem.attrib["r"]) if elem.attrib.get("r") else start_row + len(grid)
                if sheet_row < start_row:
                    elem.clear()
                    continue
                if sheet_row > end_row:
                    elem.clear()
                    break
                while start_row + len(grid) < sheet_row and len(grid) < row_limit:
                    grid.append([])
                values, row_formulas, row_extra, row_uncached = _parse_row(
                    elem,
                    sheet_row=sheet_row,
                    max_columns=max_columns,
                    strings=strings,
                    date_styles=date_styles,
                )
                if len(grid) < row_limit:
                    grid.append(values)
                    formulas.update(row_formulas)
                    uncached += row_uncached
                    extra_column = extra_column or row_extra
                elem.clear()
    except KeyError as exc:
        raise ToolFailure("read_failed", "Worksheet XML is missing.") from exc
    except (zipfile.BadZipFile, OSError, ValueError) as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc
    return grid, formulas, uncached, extra_column


def _parse_row(
    elem: Element,
    *,
    sheet_row: int,
    max_columns: int,
    strings: list[str],
    date_styles: set[int],
) -> tuple[list[str], dict[tuple[int, int], str], bool, int]:
    by_col: dict[int, str] = {}
    formulas: dict[tuple[int, int], str] = {}
    extra = False
    uncached = 0
    next_col = 1
    for child in elem:
        if _local(child.tag) != "c":
            continue
        column, value, formula = _parse_cell(child, strings, date_styles, fallback_col=next_col)
        next_col = column + 1
        if column > max_columns:
            extra = True
            continue
        text = tidy_cell(value)
        if formula:
            if not formula.startswith("="):
                formula = "=" + formula
            formulas[(sheet_row, column)] = formula
            if not text:
                uncached += 1
        by_col[column] = text
    if not by_col:
        return [], formulas, extra, uncached
    width = max(by_col)
    row = [by_col.get(col, "") for col in range(1, width + 1)]
    return row, formulas, extra, uncached


def _parse_cell(
    elem: Element,
    strings: list[str],
    date_styles: set[int],
    *,
    fallback_col: int,
) -> tuple[int, object, str | None]:
    ref = elem.attrib.get("r")
    if ref:
        _row, column = coordinate_to_tuple(ref)
    else:
        column = fallback_col
    kind = elem.attrib.get("t")
    formula = None
    raw: str | None = None
    for child in elem:
        tag = _local(child.tag)
        if tag == "f" and child.text:
            formula = child.text
        elif tag == "v" and child.text is not None:
            raw = child.text
        elif tag == "is":
            raw = "".join(node.text or "" for node in child.iter() if _local(node.tag) == "t")
            kind = "inlineStr"
    value: object
    if kind == "inlineStr":
        value = raw or ""
    elif kind == "s":
        try:
            value = strings[int(raw or "-1")]
        except (ValueError, IndexError):
            value = ""
    elif kind == "b":
        value = raw == "1"
    elif kind == "e":
        value = ""
    elif kind == "str":
        value = raw or ""
    elif raw is None:
        value = None
    else:
        try:
            number = float(raw)
        except ValueError:
            value = raw
        else:
            style = elem.attrib.get("s")
            if style is not None and style.isdigit() and int(style) in date_styles:
                try:
                    value = from_excel(number)
                except (OverflowError, ValueError):
                    value = number
            elif number.is_integer() and abs(number) < 10**15:
                value = int(number)
            else:
                value = number
    return column, value, formula


def _xlsx_support(path: Path) -> tuple[list[str], set[int]]:
    key = (*_file_token(path), "support")
    cached = _cache_get(key)
    if isinstance(cached, tuple):
        return cached
    try:
        with zipfile.ZipFile(path) as archive:
            strings = _shared_strings(archive)
            date_styles = _date_style_indexes(archive)
    except (zipfile.BadZipFile, OSError) as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc
    _cache_put(key, (strings, date_styles))
    return strings, date_styles


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    strings: list[str] = []
    with archive.open("xl/sharedStrings.xml") as handle:
        for _event, elem in defused_et.iterparse(handle, events=("end",)):
            if _local(elem.tag) != "si":
                continue
            strings.append(
                "".join(node.text or "" for node in elem.iter() if _local(node.tag) == "t")
            )
            elem.clear()
    return strings


def _date_style_indexes(archive: zipfile.ZipFile) -> set[int]:
    if "xl/styles.xml" not in archive.namelist():
        return set()
    root = _xml_root(archive, "xl/styles.xml")
    date_ids = {int(fmt_id) for fmt_id, code in BUILTIN_FORMATS.items() if is_date_format(code)}
    for elem in root.iter():
        if _local(elem.tag) == "numFmt":
            fmt_id = elem.attrib.get("numFmtId")
            code = elem.attrib.get("formatCode") or ""
            if fmt_id and fmt_id.isdigit() and is_date_format(code):
                date_ids.add(int(fmt_id))
    indexes: set[int] = set()
    for elem in root.iter():
        if _local(elem.tag) != "cellXfs":
            continue
        for index, xf in enumerate(child for child in elem if _local(child.tag) == "xf"):
            fmt_id = xf.attrib.get("numFmtId")
            if fmt_id and fmt_id.isdigit() and int(fmt_id) in date_ids:
                indexes.add(index)
    return indexes


def _apply_merges(
    grid: list[list[str]],
    start_row: int,
    merges: list[tuple[int, int, int, int]],
    max_columns: int,
) -> int:
    if not grid or not merges:
        return 0
    end_row = start_row + len(grid) - 1
    count = 0
    for min_col, min_row, max_col, max_row in merges:
        if max_row < start_row or min_row > end_row or min_col > max_columns:
            continue
        count += 1
        anchor_r = min_row - start_row
        anchor_c = min_col - 1
        if not (0 <= anchor_r < len(grid)):
            continue
        row = grid[anchor_r]
        if anchor_c >= len(row):
            continue
        value = row[anchor_c]
        last_col = min(max_col, max_columns)
        for sheet_row in range(max(min_row, start_row), min(max_row, end_row) + 1):
            rr = sheet_row - start_row
            for col in range(min_col, last_col + 1):
                cc = col - 1
                if rr == anchor_r and cc == anchor_c:
                    continue
                while len(grid[rr]) <= cc:
                    grid[rr].append("")
                if not grid[rr][cc]:
                    grid[rr][cc] = value
    return count


def _xls_sheet_names(path: Path) -> list[str]:
    book = _open_xls(path, formatting_info=False)
    try:
        return list(book.sheet_names())
    finally:
        book.release_resources()


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
    start_row: int,
    row_limit: int,
) -> SheetGrid:
    from xlrd import XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_ERROR, XL_CELL_NUMBER, xldate

    key = (*_file_token(path), "xls", sheet, sheet_index, start_row, row_limit, max_columns)
    cached = _cache_get(key)
    if isinstance(cached, SheetGrid):
        return cached
    try:
        book = _open_xls(path, formatting_info=True)
    except ToolFailure:
        book = _open_xls(path, formatting_info=False)
    try:
        worksheet, index = _select_xlrd(book, sheet, sheet_index)
        declared_rows = int(worksheet.nrows)
        declared_cols = int(worksheet.ncols)
        first = start_row - 1
        last = min(declared_rows, first + row_limit)
        grid: list[list[str]] = []
        scan_cols = min(declared_cols, max_columns)
        for row_index in range(first, last):
            row: list[str] = []
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
                row.append(tidy_cell(value))
            grid.append(row)
        merged = 0
        for rlo, rhi, clo, chi in getattr(worksheet, "merged_cells", []) or []:
            if rhi <= first or rlo >= last or clo >= scan_cols:
                continue
            anchor_r = rlo - first
            if anchor_r < 0 or anchor_r >= len(grid) or clo >= len(grid[anchor_r]):
                continue
            value = grid[anchor_r][clo]
            merged += 1
            for row_index in range(max(rlo, first), min(rhi, last)):
                for col_index in range(clo, min(chi, scan_cols)):
                    rr = row_index - first
                    if not grid[rr][col_index]:
                        grid[rr][col_index] = value
        result = SheetGrid(
            name=worksheet.name,
            index=index,
            rows=_trim_trailing_blank_rows(grid),
            start_row=start_row,
            merged_ranges=merged,
            truncated_rows=declared_rows > last,
            truncated_cols=declared_cols > scan_cols,
            declared_rows=declared_rows,
            declared_cols=declared_cols,
        )
        _cache_put(key, result)
        return result
    finally:
        book.release_resources()


def _read_csv(
    path: Path,
    *,
    kind: str,
    max_columns: int,
    start_row: int,
    row_limit: int,
) -> SheetGrid:
    key = (*_file_token(path), "csv", kind, start_row, row_limit, max_columns)
    cached = _cache_get(key)
    if isinstance(cached, SheetGrid):
        return cached
    try:
        raw = path.read_bytes()
    except OSError as exc:
        raise ToolFailure("read_failed", "File could not be read.") from exc
    text, encoding = _decode(raw)
    delimiter = "\t" if kind == "tsv" else _sniff_delimiter(text)
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows: list[list[str]] = []
    truncated_rows = False
    truncated_cols = False
    end_row = start_row + row_limit - 1
    try:
        for row_number, row in enumerate(reader, start=1):
            if row_number < start_row:
                continue
            if row_number > end_row:
                truncated_rows = True
                break
            cleaned: list[str] = []
            for index, cell in enumerate(row):
                if index >= max_columns:
                    truncated_cols = True
                    break
                cleaned.append(tidy_cell(cell))
            rows.append(cleaned)
    except csv.Error as exc:
        raise ToolFailure("read_failed", _short_read_error(exc)) from exc
    _pad(rows)
    result = SheetGrid(
        name=path.stem or "csv",
        index=0,
        rows=_trim_trailing_blank_rows(rows),
        start_row=start_row,
        encoding=encoding,
        truncated_rows=truncated_rows,
        truncated_cols=truncated_cols,
    )
    _cache_put(key, result)
    return result


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


def _select_xlrd(book, sheet: str | None, sheet_index: int | None):
    names = list(book.sheet_names())
    index = _resolve_index(names, sheet, sheet_index)
    return book.sheet_by_index(index), index


def _resolve_index(names: list[str], sheet: str | None, sheet_index: int | None) -> int:
    has_sheet = sheet is not None and str(sheet).strip() != ""
    has_index = sheet_index is not None
    if has_sheet and has_index:
        raise ToolFailure("bad_argument", "Pass sheet or sheet_index, not both.")
    if has_sheet:
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


def _pad(grid: list[list[str]]) -> None:
    width = max((len(row) for row in grid), default=0)
    for row in grid:
        if len(row) < width:
            row.extend([""] * (width - len(row)))


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
