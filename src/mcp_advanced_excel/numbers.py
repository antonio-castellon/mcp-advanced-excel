"""Decimal-comma / thousands-separator parsing for instrument exports.

AdvancedExcel converts columns with ``pandas.to_numeric`` after the sheet is
already loaded. Exports that store European decimals as text never become
numbers that way. This module sniffs the separator from the text cells and
parses them directly.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

DecimalStyle = Literal["comma", "dot"]

_NULL = re.compile(r"(?i)^(na|n/?a|n\.?d\.?|nan|null|none|nil|na_na|tbd|---+)$")
_QUALIFIER = re.compile(r"^(<=|>=|<|>)\s*(.+)$")
# LOD/LOQ tokens are measurements that were not quantified. They are not numbers.
_CENSOR_WORD = re.compile(r"(?i)^(lod|loq|dl|nd|bdl|bql|not[\s_-]*detected|below[\s_-]*detection)$")
# One decimal separator, optional exponent: 12,5 / 12.5 / 1e-3.
_SIMPLE_NUMBER = re.compile(r"^[+-]?(?:\d+(?:[.,]\d+)?|\d*[.,]\d+)(?:[eE][+-]?\d+)?$")
# Thousands groups plus an optional decimal tail: 1.250,5 / 1,234.56 / 1.234.567.
_GROUPED_NUMBER = re.compile(r"^[+-]?\d{1,3}(?:[.,]\d{3})+(?:[.,]\d+)?(?:[eE][+-]?\d+)?$")
_THOUSANDS_SPACE = re.compile(r"(?<=\d)[ \u00a0\u202f](?=\d{3}(?:\D|$))")


@dataclass(frozen=True)
class ParsedValue:
    """One cell after numeric normalization."""

    kind: Literal["empty", "null", "number", "censored", "text"]
    number: int | float | None = None
    qualifier: str | None = None
    text: str | None = None
    percent: bool = False


def _is_censor_token(text: str) -> bool:
    raw = (text or "").strip()
    if not raw:
        return False
    matched = _QUALIFIER.match(raw)
    body = matched.group(2).strip() if matched else raw
    return _CENSOR_WORD.fullmatch(body) is not None


def looks_numeric(text: str) -> bool:
    """True when the cell is a number, a censored limit, or a numeric qualifier."""
    if _is_censor_token(text):
        return True
    body, _qualifier, _percent = _peel(text)
    if body is None:
        return False
    compact = _THOUSANDS_SPACE.sub("", body).replace("'", "").replace("\u2019", "")
    return (
        _SIMPLE_NUMBER.fullmatch(compact) is not None
        or _GROUPED_NUMBER.fullmatch(compact) is not None
    )


def sniff_decimal(cells: list[str]) -> DecimalStyle:
    """Vote comma vs dot from text cells. Ambiguous thousands-only groups abstain."""
    comma = 0
    dot = 0
    for raw in cells:
        body, _qualifier, _percent = _peel(raw)
        if not body:
            continue
        compact = _THOUSANDS_SPACE.sub("", body).replace("'", "").replace("\u2019", "")
        if not re.search(r"\d", compact):
            continue
        # Drop a scientific exponent so the mantissa decides the vote.
        mantissa = re.sub(r"[eE][+-]?\d+$", "", compact)
        has_comma = "," in mantissa
        has_dot = "." in mantissa
        if has_comma and has_dot:
            if mantissa.rfind(",") > mantissa.rfind("."):
                comma += 1
            else:
                dot += 1
            continue
        if has_comma or has_dot:
            sep = "," if has_comma else "."
            other: DecimalStyle = "dot" if sep == "," else "comma"
            parts = mantissa.split(sep)
            if len(parts) > 2 and all(len(part) == 3 for part in parts[1:]):
                # 1.234.567 or 1,234,567 — separator is thousands.
                if other == "comma":
                    comma += 1
                else:
                    dot += 1
                continue
            frac = parts[-1]
            if len(parts) == 2 and len(frac) == 3 and len(parts[0]) <= 3:
                # Single group of three (1,234 or 1.234) is ambiguous.
                continue
            if sep == ",":
                comma += 1
            else:
                dot += 1
    if comma > dot:
        return "comma"
    return "dot"


def parse_number(text: str, decimal: DecimalStyle) -> ParsedValue:
    """Parse one cell. Non-numbers come back as text so identifiers survive."""
    raw = (text or "").strip()
    if not raw:
        return ParsedValue("empty")
    if _NULL.fullmatch(raw):
        return ParsedValue("null")
    body, qualifier, percent = _peel(raw)
    if _is_censor_token(raw):
        token = qualifier or (body or raw)
        return ParsedValue("censored", number=None, qualifier=token, text=raw)
    if body is None:
        return ParsedValue("text", text=raw)
    number = _to_number(body, decimal)
    if number is None:
        return ParsedValue("text", text=raw)
    if qualifier:
        # Keep the raw text. Do not turn "<0.5" into the number 0.5.
        return ParsedValue("censored", number=None, qualifier=qualifier, text=raw, percent=percent)
    return ParsedValue("number", number=number, percent=percent)


def _peel(text: str) -> tuple[str | None, str | None, bool]:
    raw = (text or "").strip()
    if not raw or _NULL.fullmatch(raw):
        return None, None, False
    qualifier = None
    matched = _QUALIFIER.match(raw)
    if matched:
        qualifier = matched.group(1)
        raw = matched.group(2).strip()
    percent = False
    if raw.endswith("%"):
        percent = True
        raw = raw[:-1].strip()
    if not raw:
        return None, qualifier, percent
    return raw, qualifier, percent


def _to_number(body: str, decimal: DecimalStyle) -> int | float | None:
    compact = _THOUSANDS_SPACE.sub("", body)
    compact = compact.replace("'", "").replace("\u2019", "")
    had_decimal_sep = ("," in compact) or ("." in compact) or ("e" in compact.lower())
    if decimal == "comma":
        # Dots are thousands; the comma is the decimal mark.
        if compact.count(",") > 1:
            return None
        normalized = compact.replace(".", "").replace(",", ".")
    else:
        if compact.count(".") > 1:
            return None
        normalized = compact.replace(",", "")
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", normalized):
        return None
    try:
        value = float(normalized)
    except ValueError:
        return None
    if not _finite(value):
        return None
    # Keep instrument codes like "007" as text instead of the integer 7.
    if not had_decimal_sep and _leading_zero_int(compact):
        return None
    if not had_decimal_sep and value.is_integer():
        return int(value)
    return value


def _leading_zero_int(compact: str) -> bool:
    token = compact[1:] if compact[:1] in "+-" else compact
    return len(token) > 1 and token.startswith("0")


def _finite(value: float) -> bool:
    return value == value and value not in (float("inf"), float("-inf"))
