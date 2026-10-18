"""Allowlisted filesystem access.

Read roots and write roots are separate. A path is accepted only after
``Path.resolve`` follows symlinks, so a link inside a root that points outside
is rejected. The opened path is the resolved one.
"""

from __future__ import annotations

import os
from pathlib import Path

from .errors import ToolFailure

_URL_PREFIXES = ("http://", "https://", "ftp://", "file://")


def parse_roots(raw: str | None) -> tuple[Path, ...]:
    """Split a path-separator list. Missing directories are dropped."""
    text = (raw or "").strip()
    if not text:
        return ()
    roots: list[Path] = []
    for part in text.split(os.pathsep):
        part = part.strip()
        if not part:
            continue
        try:
            resolved = Path(part).expanduser().resolve(strict=False)
        except OSError:
            continue
        if resolved.is_dir():
            roots.append(resolved)
    return tuple(roots)


def resolve_read_file(path: str, roots: tuple[Path, ...]) -> Path:
    """Resolve a file that must already exist inside a read root."""
    if not roots:
        raise ToolFailure(
            "roots_unset",
            "EXCEL_ALLOWED_ROOTS is empty or unset. Refusing to read any file.",
        )
    candidate = _candidate(path)
    try:
        resolved = candidate.resolve(strict=True)
    except OSError as exc:
        raise ToolFailure(
            "invalid_path",
            "Path does not exist or cannot be resolved inside the allowlist.",
        ) from exc
    if not _within(resolved, roots):
        raise ToolFailure(
            "outside_allowlist",
            "Path is outside EXCEL_ALLOWED_ROOTS.",
        )
    if not resolved.is_file():
        raise ToolFailure("not_a_file", "Path is not a regular file.")
    return resolved


def resolve_write_file(path: str, roots: tuple[Path, ...], *, overwrite: bool) -> Path:
    """Resolve a new ``.csv`` path that must land inside a write root."""
    if not roots:
        raise ToolFailure(
            "write_disabled",
            "EXCEL_OUTPUT_ROOTS is empty. Refusing to write.",
        )
    candidate = _candidate(path)
    if candidate.suffix.lower() != ".csv":
        raise ToolFailure("bad_extension", "Cleaned output must use a .csv filename.")
    if candidate.name in {"", ".", ".."}:
        raise ToolFailure("invalid_path", "Output filename is not valid.")
    try:
        parent = candidate.parent.resolve(strict=True)
    except OSError as exc:
        raise ToolFailure(
            "output_dir_missing",
            "Output directory does not exist inside EXCEL_OUTPUT_ROOTS.",
        ) from exc
    if not parent.is_dir() or not _within(parent, roots):
        raise ToolFailure(
            "outside_output_root",
            "Output path is outside EXCEL_OUTPUT_ROOTS.",
        )
    final = parent / candidate.name
    if final.exists() or final.is_symlink():
        try:
            resolved = final.resolve(strict=True)
        except OSError as exc:
            raise ToolFailure(
                "invalid_path",
                "Output path cannot be resolved inside EXCEL_OUTPUT_ROOTS.",
            ) from exc
        if not _within(resolved, roots):
            raise ToolFailure(
                "outside_output_root",
                "Output path is outside EXCEL_OUTPUT_ROOTS.",
            )
        if not overwrite:
            raise ToolFailure(
                "output_exists",
                "Output file already exists. Pass overwrite=true to replace it.",
            )
        if not resolved.is_file():
            raise ToolFailure("not_a_file", "Output path is not a regular file.")
        return resolved
    return final


def _candidate(path: str) -> Path:
    raw = (path or "").strip()
    if not raw or "\x00" in raw:
        raise ToolFailure("invalid_path", "Path is empty or invalid.")
    lowered = raw.lower()
    if lowered.startswith(_URL_PREFIXES):
        raise ToolFailure("invalid_path", "URLs are not allowed. Pass a local file path.")
    candidate = Path(raw).expanduser()
    if not candidate.is_absolute():
        candidate = Path.cwd() / candidate
    return candidate


def _within(path: Path, roots: tuple[Path, ...]) -> bool:
    for root in roots:
        try:
            path.relative_to(root)
        except ValueError:
            continue
        return True
    return False
