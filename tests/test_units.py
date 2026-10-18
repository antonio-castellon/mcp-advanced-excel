"""Unit aliases and conversion into canonical ASCII units."""

from __future__ import annotations

import json

import pytest

from mcp_advanced_excel.units import load_units_file, lookup_unit, register_unit, split_header_cell


def test_concentration_aliases_share_canonical_mg_per_l() -> None:
    ug = lookup_unit("ug/mL")
    mg = lookup_unit("mg/L")
    micro = lookup_unit("\u00b5g/mL")
    assert ug is not None and mg is not None and micro is not None
    assert ug.canonical == mg.canonical == "mg/L"
    assert ug.convert(2.5) == pytest.approx(2.5)
    assert micro.convert(2.5) == pytest.approx(2.5)
    assert lookup_unit("g/L").convert(1) == pytest.approx(1000)
    assert lookup_unit("ng/mL").convert(5) == pytest.approx(0.005)


def test_temperature_kelvin_and_fahrenheit_to_degc() -> None:
    kelvin = lookup_unit("K")
    fahrenheit = lookup_unit("degF")
    assert kelvin is not None and fahrenheit is not None
    assert kelvin.canonical == "degC"
    assert kelvin.convert(298.15) == pytest.approx(25.0)
    assert fahrenheit.convert(77) == pytest.approx(25.0)
    assert lookup_unit("\u00b0C").canonical == "degC"


def test_split_header_extracts_unit_and_keeps_words() -> None:
    assert split_header_cell("Conc. (ug/mL)") == ("Conc.", "ug/mL")
    assert split_header_cell("Temp (K)") == ("Temp", "K")
    assert split_header_cell("mg/L") == ("", "mg/L")
    assert split_header_cell("Sample (replicate)") == ("Sample (replicate)", None)


def test_register_and_json_overlay(tmp_path) -> None:
    table = {}
    register_unit(
        "RFU",
        canonical="RFU",
        dimension="signal",
        factor=1.0,
        table=table,
    )
    path = tmp_path / "units.json"
    path.write_text(
        json.dumps(
            {
                "units": [
                    {
                        "aliases": ["cups"],
                        "canonical": "mL",
                        "dimension": "volume",
                        "factor": 236.588,
                    }
                ]
            }
        ),
        encoding="utf-8",
    )
    loaded = load_units_file(path)
    assert lookup_unit("cups", loaded).convert(1) == pytest.approx(236.588)
    assert lookup_unit("mg/L", loaded) is not None
    assert lookup_unit("RFU", table).dimension == "signal"
