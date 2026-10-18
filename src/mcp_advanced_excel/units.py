"""Small extensible unit table.

Canonical names are ASCII on purpose (``degC``, ``ug/mL``) so tool output stays
stable across encodings. Add aliases with ``register_unit`` or an
``EXCEL_UNITS_FILE`` JSON document (see docs/UNITS.md).

Mass-per-volume factors are exact SI conversions. ``ppm`` is the usual dilute
aqueous approximation (1 ppm = 1 mg/L) and is tagged with a note.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

_MICRO = ("\u00b5", "\u03bc")
_DEGREE = "\u00b0"
_PAREN = re.compile(r"^(?P<name>.*?)(?:\s*[\(\[]\s*(?P<unit>[^)\]]+?)\s*[\)\]])\s*$")
_UNIT_CHARSET = re.compile(r"^[A-Za-z0-9/^.\-]+$")
_NOT_A_UNIT = {
    "id",
    "sample",
    "name",
    "date",
    "time",
    "total",
    "mean",
    "avg",
    "average",
    "replicate",
    "raw",
    "blank",
    "control",
    "positive",
    "negative",
    "std",
    "standard",
    "count",
    "flag",
    "note",
    "notes",
    "type",
    "batch",
    "lot",
}


@dataclass(frozen=True)
class UnitSpec:
    """Convert ``canonical = value * factor + offset``."""

    canonical: str
    dimension: str
    factor: float = 1.0
    offset: float = 0.0
    note: str | None = None

    def convert(self, value: float) -> float:
        converted = (value * self.factor) + self.offset
        rounded = round(converted, 10)
        if rounded == 0:
            return 0.0
        return rounded


def unit_key(raw: str) -> str:
    """Fold an alias to the registry key."""
    text = (raw or "").strip()
    for mark in _MICRO:
        text = text.replace(mark, "u")
    text = text.replace(_DEGREE, "deg")
    text = text.lower()
    text = text.replace(" ", "").replace("_", "")
    text = text.replace("per", "/")
    # mg.L-1 and mgL-1, common on chromatography reports.
    text = re.sub(r"^([a-z]+)(?:\.)?([a-z]+)-1$", r"\1/\2", text)
    return text


def _spec(
    canonical: str,
    dimension: str,
    factor: float = 1.0,
    offset: float = 0.0,
    note: str | None = None,
) -> UnitSpec:
    return UnitSpec(canonical, dimension, factor, offset, note)


def _build_builtin() -> dict[str, UnitSpec]:
    table: dict[str, UnitSpec] = {}

    def add(spec: UnitSpec, *aliases: str) -> None:
        for alias in aliases:
            table[unit_key(alias)] = spec

    # Concentration, canonical mg/L. 1 ug/mL == 1 mg/L.
    conc = _spec("mg/L", "concentration", 1.0)
    add(conc, "mg/L", "mg/l", "mg per L", "mg.L-1", "ppm")
    table[unit_key("ppm")] = _spec(
        "mg/L",
        "concentration",
        1.0,
        note="ppm is treated as mg/L (dilute aqueous approximation)",
    )
    add(_spec("mg/L", "concentration", 1.0), "ug/mL", "ug/ml", "mcg/mL", "mcg/ml")
    add(_spec("mg/L", "concentration", 1000.0), "g/L", "g/l")
    add(_spec("mg/L", "concentration", 0.001), "ug/L", "ug/l", "mcg/L")
    add(_spec("mg/L", "concentration", 0.001), "ng/mL", "ng/ml")
    add(_spec("mg/L", "concentration", 1e-6), "ng/L", "ng/l")
    add(_spec("mg/L", "concentration", 1.0), "g/m3", "g/m^3")
    add(_spec("mg/L", "concentration", 1000.0), "kg/m3", "kg/m^3")
    add(_spec("mg/L", "concentration", 1000.0), "mg/mL", "mg/ml")

    # Temperature, canonical degC. Affine, so offset is required.
    add(_spec("degC", "temperature", 1.0, 0.0), "degC", "celsius", "centigrade")
    add(_spec("degC", "temperature", 1.0, -273.15), "K", "kelvin")
    add(
        _spec("degC", "temperature", 5.0 / 9.0, -32.0 * 5.0 / 9.0),
        "degF",
        "fahrenheit",
    )

    # Mass, canonical g.
    add(_spec("g", "mass", 1.0), "g", "gram", "grams")
    add(_spec("g", "mass", 0.001), "mg", "milligram", "milligrams")
    add(_spec("g", "mass", 1e-6), "ug", "mcg", "microgram", "micrograms")
    add(_spec("g", "mass", 1e-9), "ng", "nanogram", "nanograms")
    add(_spec("g", "mass", 1000.0), "kg", "kilogram", "kilograms")

    # Volume, canonical mL.
    add(_spec("mL", "volume", 1.0), "mL", "ml", "milliliter", "milliliters")
    add(_spec("mL", "volume", 1000.0), "L", "l", "liter", "liters", "litre", "litres")
    add(_spec("mL", "volume", 0.001), "uL", "ul", "microliter", "microliters")
    add(_spec("mL", "volume", 100.0), "dL", "dl")

    # Pressure, canonical kPa.
    add(_spec("kPa", "pressure", 1.0), "kPa", "kpa")
    add(_spec("kPa", "pressure", 0.001), "Pa", "pa")
    add(_spec("kPa", "pressure", 1000.0), "MPa", "mpa")
    add(_spec("kPa", "pressure", 100.0), "bar")
    add(_spec("kPa", "pressure", 0.1), "mbar")
    add(_spec("kPa", "pressure", 101.325), "atm")
    add(_spec("kPa", "pressure", 6.8947572932), "psi")
    add(_spec("kPa", "pressure", 0.1333223874), "mmHg", "torr")

    # Length, canonical m. Bare "m" is metre, not minute (minute is "min").
    add(_spec("m", "length", 1.0), "m", "meter", "metre", "meters", "metres")
    add(_spec("m", "length", 0.01), "cm")
    add(_spec("m", "length", 0.001), "mm")
    add(_spec("m", "length", 1e-6), "um", "micron", "microns")
    add(_spec("m", "length", 1e-9), "nm")
    add(_spec("m", "length", 1000.0), "km")

    # Time, canonical s.
    add(_spec("s", "time", 1.0), "s", "sec", "second", "seconds")
    add(_spec("s", "time", 0.001), "ms", "millisecond", "milliseconds")
    add(_spec("s", "time", 60.0), "min", "minute", "minutes")
    add(_spec("s", "time", 3600.0), "h", "hr", "hour", "hours")

    add(_spec("%", "fraction", 1.0), "%", "percent", "pct")
    return table


BUILTIN: dict[str, UnitSpec] = _build_builtin()


def lookup_unit(raw: str, table: dict[str, UnitSpec] | None = None) -> UnitSpec | None:
    key = unit_key(raw)
    if not key:
        return None
    return (table if table is not None else BUILTIN).get(key)


def register_unit(
    alias: str,
    *,
    canonical: str,
    dimension: str,
    factor: float = 1.0,
    offset: float = 0.0,
    note: str | None = None,
    table: dict[str, UnitSpec] | None = None,
) -> None:
    """Add or replace one alias. Pass ``table`` to keep the builtin map intact."""
    target = BUILTIN if table is None else table
    target[unit_key(alias)] = UnitSpec(canonical, dimension, factor, offset, note)


def split_header_cell(
    cell: str, table: dict[str, UnitSpec] | None = None
) -> tuple[str, str | None]:
    """Split ``Conc. (mg/L)`` into a name and a unit. Pure unit cells have an empty name."""
    text = (cell or "").strip()
    if not text:
        return "", None
    if _is_pure_unit(text, table):
        return "", text
    matched = _PAREN.match(text)
    if matched:
        unit = matched.group("unit").strip()
        name = matched.group("name").strip(" -_/")
        if name and _accept_unit_token(unit, table):
            return name, unit
    return text, None


def is_pure_unit(cell: str, table: dict[str, UnitSpec] | None = None) -> bool:
    return _is_pure_unit(cell, table)


def _is_pure_unit(text: str, table: dict[str, UnitSpec] | None) -> bool:
    """True when the whole cell is a unit, not a short English header.

    Unknown words such as ``Value`` or ``Flag`` stay names. Unknown tokens
    still count when they contain ``/``, ``%``, ``^``, or a digit (``mg/L``,
    ``m2``). Short labels like ``RFU`` are accepted only inside parentheses.
    """
    if not text or any(ch.isspace() for ch in text):
        return False
    if lookup_unit(text, table) is not None:
        return True
    if not _UNIT_CHARSET.fullmatch(text) or len(text) > 24:
        return False
    return any(mark in text for mark in ("/", "%", "^")) or any(ch.isdigit() for ch in text)


def _accept_unit_token(token: str, table: dict[str, UnitSpec] | None) -> bool:
    text = token.strip()
    if not text or len(text) > 24 or " " in text:
        return False
    if lookup_unit(text, table) is not None:
        return True
    key = unit_key(text)
    if key in _NOT_A_UNIT:
        return False
    if not _UNIT_CHARSET.fullmatch(text):
        return False
    # Unknown short tokens with a slash, digit, or percent are still unit labels
    # (RFU is all letters and is handled below; "replicate" is denied).
    if any(mark in text for mark in ("/", "%", "^")) or any(ch.isdigit() for ch in text):
        return True
    # Short all-letter labels (RFU, NTU, AU) are kept as unconverted units.
    return text.isalpha() and 1 < len(text) <= 5 and key not in _NOT_A_UNIT


def load_units_file(path: Path, base: dict[str, UnitSpec] | None = None) -> dict[str, UnitSpec]:
    """Copy ``base`` (builtin by default) and overlay a JSON unit file."""
    table = dict(BUILTIN if base is None else base)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read unit table: {exc}") from exc
    entries = payload.get("units") if isinstance(payload, dict) else None
    if not isinstance(entries, list):
        raise ValueError("Unit table JSON must be an object with a 'units' list")
    for entry in entries:
        if not isinstance(entry, dict):
            raise ValueError("Each unit entry must be an object")
        aliases = entry.get("aliases")
        canonical = entry.get("canonical")
        dimension = entry.get("dimension")
        if not isinstance(aliases, list) or not aliases:
            raise ValueError("Unit entry needs a non-empty aliases list")
        if not isinstance(canonical, str) or not canonical.strip():
            raise ValueError("Unit entry needs a canonical name")
        if not isinstance(dimension, str) or not dimension.strip():
            raise ValueError("Unit entry needs a dimension")
        factor = float(entry.get("factor", 1.0))
        offset = float(entry.get("offset", 0.0))
        note = entry.get("note")
        if note is not None and not isinstance(note, str):
            raise ValueError("Unit note must be a string")
        spec = UnitSpec(canonical.strip(), dimension.strip(), factor, offset, note)
        for alias in aliases:
            if not isinstance(alias, str) or not alias.strip():
                raise ValueError("Unit aliases must be strings")
            table[unit_key(alias)] = spec
    return table
