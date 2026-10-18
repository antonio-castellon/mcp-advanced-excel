"""Find header rows and data blocks in a rectangular string grid.

The block split follows AdvancedExcel ``get_all_tables``: a blank row ends a
block. Inside a block, the header is the text row (plus one parent row when it
is a unit row or a merged group label) sitting immediately above the longest
run of measurement rows. Key/value metadata above that header stays metadata.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .numbers import looks_numeric
from .units import is_pure_unit

_FOOTER = re.compile(
    r"(?i)\b(notes?|end of (report|file|data)|disclaimer|confidential|"
    r"printed|page \d+|do not distribute|comment)\b"
)


@dataclass
class DetectedTable:
    header_rows: list[int]
    data_rows: list[int]
    preamble_rows: list[int] = field(default_factory=list)
    note_rows: list[int] = field(default_factory=list)


def detect_layout(grid: list[list[str]]) -> tuple[list[DetectedTable], list[int]]:
    """Return tables in sheet order and metadata row indexes."""
    if not grid:
        return [], []
    blocks = _blocks(grid)
    tables: list[DetectedTable] = []
    metadata: list[int] = []
    for start, end in blocks:
        detected = _classify_block(grid, start, end)
        if detected is None:
            if tables:
                tables[-1].note_rows.extend(range(start, end))
            else:
                metadata.extend(range(start, end))
            continue
        # Sparse rows between this block and the previous table were already
        # stored. Rows inside the block before the header are preamble.
        tables.append(detected)
    return tables, metadata


def default_table_index(tables: list[DetectedTable]) -> int | None:
    """Prefer the table with the most data rows, then the widest, then the first."""
    if not tables:
        return None
    return max(
        range(len(tables)),
        key=lambda index: (len(tables[index].data_rows), _width_hint(tables[index]), -index),
    )


def _width_hint(table: DetectedTable) -> int:
    return len(table.header_rows) + len(table.data_rows)


def _blocks(grid: list[list[str]]) -> list[tuple[int, int]]:
    blocks: list[tuple[int, int]] = []
    start: int | None = None
    for index, row in enumerate(grid):
        if _nonempty(row):
            if start is None:
                start = index
        elif start is not None:
            blocks.append((start, index))
            start = None
    if start is not None:
        blocks.append((start, len(grid)))
    return blocks


def _classify_block(grid: list[list[str]], start: int, end: int) -> DetectedTable | None:
    rows = grid[start:end]
    flags = [_is_data_row(row) for row in rows]
    run = _longest_true_run(flags)
    if run is not None:
        data_start, data_end = run
        if data_start == 0 and (data_end - data_start) < 2 and _header_rows_for(rows, 0) == []:
            # A lone key/value row is metadata, not a one-cell table.
            if _looks_like_kv(rows[0]):
                return None
        header_rows = _header_rows_for(rows, data_start)
        global_header = [start + offset for offset in header_rows]
        preamble_end = header_rows[0] if header_rows else data_start
        return DetectedTable(
            header_rows=global_header,
            data_rows=list(range(start + data_start, start + data_end)),
            preamble_rows=list(range(start, start + preamble_end)),
            note_rows=list(range(start + data_end, end)),
        )
    if _is_text_table(rows):
        header_local = _header_rows_for(rows, 1) or [0]
        # _header_rows_for assumes a data row at index 1; force the text header.
        if 0 not in header_local:
            header_local = [0]
        return DetectedTable(
            header_rows=[start + offset for offset in header_local],
            data_rows=list(range(start + max(header_local) + 1, end)),
            preamble_rows=list(range(start, start + min(header_local))),
        )
    return None


def _header_rows_for(rows: list[list[str]], data_start: int) -> list[int]:
    if data_start <= 0:
        return []
    primary = data_start - 1
    if not _is_headerish(rows[primary]):
        return []
    chosen = [primary]
    above = primary - 1
    if above >= 0 and _include_parent(rows[above], rows[primary]):
        chosen.insert(0, above)
    return chosen


def _include_parent(upper: list[str], lower: list[str]) -> bool:
    if not _is_headerish(upper) or _looks_like_kv(upper):
        return False
    if _pure_unit_fraction(lower) >= 0.5:
        return True
    if _has_adjacent_duplicate(upper):
        return True
    if _nonempty_count(upper) + 1 < _nonempty_count(lower):
        return True
    return False


def _is_text_table(rows: list[list[str]]) -> bool:
    if len(rows) < 3:
        return False
    width = _nonempty_count(rows[0])
    if width < 2 or any(looks_numeric(cell) for cell in rows[0] if cell.strip()):
        return False
    # "Key: Value" stacks are metadata, not a header plus records.
    if all(":" in " ".join(_cells(row)) for row in rows):
        return False
    matches = sum(1 for row in rows[1:] if abs(_nonempty_count(row) - width) <= 1)
    return matches >= 2


def _is_data_row(row: list[str]) -> bool:
    cells = _cells(row)
    if not cells:
        return False
    numeric = sum(1 for cell in cells if looks_numeric(cell))
    if numeric < 1:
        return False
    if len(cells) == 1 and not looks_numeric(cells[0]):
        return False
    return True


def _is_headerish(row: list[str]) -> bool:
    cells = _cells(row)
    if len(cells) < 1:
        return False
    numeric = sum(1 for cell in cells if looks_numeric(cell))
    return numeric < len(cells)


def _looks_like_kv(row: list[str]) -> bool:
    cells = _cells(row)
    if len(cells) == 1 and ":" in cells[0] and not looks_numeric(cells[0]):
        return True
    if len(cells) == 2 and not looks_numeric(cells[0]) and len(cells[0]) <= 40:
        return True
    return False


def _pure_unit_fraction(row: list[str]) -> float:
    cells = _cells(row)
    if not cells:
        return 0.0
    pure = sum(1 for cell in cells if is_pure_unit(cell))
    return pure / len(cells)


def _has_adjacent_duplicate(row: list[str]) -> bool:
    for index in range(len(row) - 1):
        left = row[index].strip()
        right = row[index + 1].strip()
        if left and left == right:
            return True
    return False


def _cells(row: list[str]) -> list[str]:
    return [cell.strip() for cell in row if cell and cell.strip()]


def _nonempty(row: list[str]) -> bool:
    return any(cell and cell.strip() for cell in row)


def _nonempty_count(row: list[str]) -> int:
    return len(_cells(row))


def _longest_true_run(flags: list[bool]) -> tuple[int, int] | None:
    best: tuple[int, int, int] | None = None
    start: int | None = None
    for index, flag in enumerate([*flags, False]):
        if flag and start is None:
            start = index
        elif not flag and start is not None:
            length = index - start
            if best is None or length > best[0]:
                best = (length, start, index)
            start = None
    if best is None:
        return None
    return best[1], best[2]


def is_footer_text(text: str) -> bool:
    return _FOOTER.search(text or "") is not None
