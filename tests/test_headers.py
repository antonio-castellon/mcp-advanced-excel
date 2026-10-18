"""Header detection on synthetic instrument exports."""

from __future__ import annotations

from pathlib import Path

from mcp_advanced_excel.config import Settings
from mcp_advanced_excel.normalize import normalize_file

FIXTURES = Path(__file__).resolve().parent / "fixtures"


def test_skips_metadata_blank_rows_and_footer(fixture_settings: Settings) -> None:
    table = normalize_file(str(FIXTURES / "instrument.xlsx"), fixture_settings)
    payload = table.as_dict(include_rows=True)
    assert payload["sheet"] == "Results"
    assert payload["header_rows"] == [4]
    assert payload["data_start_row"] == 5
    assert payload["data_end_row"] == 6
    assert payload["metadata"][0]["text"] == "Instrument: Agilent 1260"
    assert payload["metadata"][1]["text"] == "Method: HPLC-UV"
    assert payload["notes"][0]["row"] == 8
    assert "End of report" in payload["notes"][0]["text"]
    assert [column["name"] for column in payload["columns"]] == [
        "sample_id",
        "compound",
        "conc",
        "temp",
    ]
    by_name = {column["name"]: column for column in payload["columns"]}
    assert payload["decimal"] == "comma"
    assert by_name["conc"]["source_unit"] == "ug/mL"
    assert by_name["conc"]["unit"] == "mg/L"
    assert by_name["conc"]["unit_converted"] is True
    assert by_name["temp"]["source_unit"] == "K"
    assert by_name["temp"]["unit"] == "degC"
    assert payload["rows"][0]["sample_id"] == "S-001"
    assert payload["rows"][0]["conc"] == 12.5
    assert payload["rows"][1]["conc"] == 1234.5
    assert payload["rows"][0]["temp"] == 25
    assert payload["rows"][1]["temp"] == 26.85


def test_multi_row_header_and_merged_cells(fixture_settings: Settings) -> None:
    table = normalize_file(str(FIXTURES / "merged_header.xlsx"), fixture_settings)
    payload = table.as_dict(include_rows=True)
    assert payload["merged_ranges"] >= 1
    assert payload["header_rows"] == [3, 4]
    assert payload["data_start_row"] == 5
    names = [column["name"] for column in payload["columns"]]
    assert names == ["sample_id", "concentration", "concentration_2", "temperature"]
    units = {column["name"]: column["unit"] for column in payload["columns"]}
    assert units["concentration"] == "mg/L"
    assert units["concentration_2"] == "mg/L"
    assert units["temperature"] == "degC"
    assert payload["rows"] == [
        {
            "sample_id": "S-1",
            "concentration": 10,
            "concentration_2": 10,
            "temperature": 25,
        }
    ]


def test_header_row_override(tmp_path: Path) -> None:
    # A two-column block with no numeric cells is metadata unless forced.
    target = tmp_path / "qual.csv"
    target.write_text("Sample,Flag\nS-1,ok\nS-2,ok\n", encoding="utf-8")
    settings = Settings(allowed_roots=(tmp_path,))
    forced = normalize_file(str(target), settings, header_row=1)
    assert forced.header_rows == [1]
    assert forced.rows == [
        {"sample": "S-1", "flag": "ok"},
        {"sample": "S-2", "flag": "ok"},
    ]


def test_largest_table_is_the_default(tmp_path: Path) -> None:
    target = tmp_path / "blocks.csv"
    target.write_text(
        "Sample,Value\nonly,1\n\nSample,Value\na,2\nb,3\nc,4\n",
        encoding="utf-8",
    )
    settings = Settings(allowed_roots=(tmp_path,))
    table = normalize_file(str(target), settings)
    assert table.table_count == 2
    assert table.table_index == 1
    assert len(table.rows) == 3
    first = normalize_file(str(target), settings, table_index=0)
    assert len(first.rows) == 1
    assert first.rows[0]["sample"] == "only"
