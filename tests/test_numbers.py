"""Decimal commas, thousands separators, and null tokens."""

from __future__ import annotations

from mcp_advanced_excel.numbers import parse_number, sniff_decimal


def test_grouped_thousands_still_look_numeric() -> None:
    from mcp_advanced_excel.numbers import looks_numeric

    assert looks_numeric("1.250,5")
    assert looks_numeric("1,234.56")
    assert looks_numeric("1 234,56")
    assert not looks_numeric("S-001")


def test_sniff_prefers_decimal_comma_over_dot_thousands() -> None:
    style = sniff_decimal(["12,5", "1.234,56", "0,125"])
    assert style == "comma"


def test_parse_decimal_comma_and_thousands() -> None:
    assert parse_number("12,5", "comma").number == 12.5
    assert parse_number("1.234,56", "comma").number == 1234.56
    assert parse_number("1 234,56", "comma").number == 1234.56
    assert parse_number("1'234,56", "comma").number == 1234.56
    assert parse_number("1,25e-1", "comma").number == 0.125


def test_parse_dot_thousands_and_qualifiers() -> None:
    parsed = parse_number("1,234.50", "dot")
    assert parsed.number == 1234.5
    censored = parse_number("<0,5", "comma")
    assert censored.kind == "censored"
    assert censored.qualifier == "<"
    assert censored.number == 0.5


def test_identifiers_and_nulls_stay_non_numeric() -> None:
    assert parse_number("S-001", "comma").kind == "text"
    assert parse_number("007", "dot").kind == "text"
    assert parse_number("NA", "dot").kind == "null"
    assert parse_number("n.d.", "comma").kind == "null"
    assert parse_number("", "dot").kind == "empty"


def test_ambiguous_thousands_group_does_not_force_comma() -> None:
    # A lone 1,234 abstains; the US-style value decides the vote.
    assert sniff_decimal(["1,234", "2.5"]) == "dot"
    assert parse_number("1,234", "dot").number == 1234
