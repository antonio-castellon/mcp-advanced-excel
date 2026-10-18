"""Environment-backed configuration. Nothing is read until a tool is called."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from .paths import parse_roots, resolve_read_file
from .units import BUILTIN, UnitSpec, load_units_file

DEFAULT_MAX_FILE_BYTES = 25 * 1024 * 1024
DEFAULT_MAX_ROWS = 5_000
DEFAULT_MAX_COLUMNS = 128
# Furthest sheet row a single call will stream. Skipped rows are not stored.
# The detection prefix is much smaller (see grid.DETECT_PREFIX_ROWS).
DEFAULT_MAX_SCAN_ROWS = 100_000

_CEILING_FILE = 100 * 1024 * 1024
_CEILING_ROWS = 100_000
_CEILING_COLUMNS = 512
_CEILING_SCAN = 100_000


def _parse_bool(raw: str | None, *, default: bool = False) -> bool:
    text = (raw or "").strip().lower()
    if not text:
        return default
    return text in {"1", "true", "yes", "on"}


def _parse_int(raw: str | None, *, default: int, ceiling: int) -> int:
    text = (raw or "").strip()
    if not text:
        return default
    try:
        value = int(text)
    except ValueError:
        return default
    if value < 1:
        return default
    return min(value, ceiling)


@dataclass(frozen=True)
class Settings:
    """Runtime limits and allowlists. Write stays off unless explicitly enabled."""

    allowed_roots: tuple[Path, ...] = ()
    output_roots: tuple[Path, ...] = ()
    allow_write: bool = False
    max_file_bytes: int = DEFAULT_MAX_FILE_BYTES
    max_rows: int = DEFAULT_MAX_ROWS
    max_columns: int = DEFAULT_MAX_COLUMNS
    max_scan_rows: int = DEFAULT_MAX_SCAN_ROWS
    redact: str = "off"
    units: dict[str, UnitSpec] = field(default_factory=lambda: dict(BUILTIN))

    @classmethod
    def from_env(cls) -> Settings:
        units = dict(BUILTIN)
        allowed = parse_roots(os.getenv("EXCEL_ALLOWED_ROOTS"))
        units_file = (os.getenv("EXCEL_UNITS_FILE") or "").strip()
        if units_file:
            # The overlay itself has to live inside the read allowlist.
            resolved = resolve_read_file(units_file, allowed)
            try:
                units = load_units_file(resolved, units)
            except ValueError as exc:
                raise RuntimeError(str(exc)) from exc
        max_rows = _parse_int(
            os.getenv("EXCEL_MAX_ROWS"), default=DEFAULT_MAX_ROWS, ceiling=_CEILING_ROWS
        )
        scan_default = max(DEFAULT_MAX_SCAN_ROWS, max_rows)
        max_scan = _parse_int(
            os.getenv("EXCEL_MAX_SCAN_ROWS"),
            default=scan_default,
            ceiling=_CEILING_SCAN,
        )
        return cls(
            allowed_roots=allowed,
            output_roots=parse_roots(os.getenv("EXCEL_OUTPUT_ROOTS")),
            allow_write=_parse_bool(os.getenv("EXCEL_ALLOW_WRITE")),
            max_file_bytes=_parse_int(
                os.getenv("EXCEL_MAX_FILE_BYTES"),
                default=DEFAULT_MAX_FILE_BYTES,
                ceiling=_CEILING_FILE,
            ),
            max_rows=max_rows,
            max_columns=_parse_int(
                os.getenv("EXCEL_MAX_COLUMNS"),
                default=DEFAULT_MAX_COLUMNS,
                ceiling=_CEILING_COLUMNS,
            ),
            max_scan_rows=max(max_scan, max_rows),
            redact=_parse_redact(os.getenv("EXCEL_REDACT")),
            units=units,
        )


def _parse_redact(raw: str | None) -> str:
    text = (raw or "").strip().lower()
    if text in {"secrets", "on", "1", "true", "yes"}:
        return "secrets"
    return "off"


def load_dotenv_if_present(path: Path | None = None) -> None:
    """Load a dotenv file only when the caller passes a path or sets EXCEL_DOTENV.

    The process does not read ``.env`` from the working directory.
    """
    if path is None:
        raw = (os.getenv("EXCEL_DOTENV") or "").strip()
        if not raw:
            return
        candidate = Path(raw).expanduser()
    else:
        candidate = path
    if not candidate.is_file():
        return
    try:
        for line in candidate.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in stripped:
                continue
            key, _, value = stripped.partition("=")
            key = key.strip()
            value = value.strip().strip("'").strip('"')
            if key and key not in os.environ:
                os.environ[key] = value
    except OSError:
        return
