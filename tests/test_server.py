"""Tool wiring, CSV quirks, scrubbing, formulas, and the write gate."""

from __future__ import annotations

from pathlib import Path

from mcp_advanced_excel.config import Settings
from mcp_advanced_excel.server import create_server, list_tool_names

FIXTURES = Path(__file__).resolve().parent / "fixtures"

EXPECTED_TOOLS = {
    "excel_list_sheets",
    "excel_overview",
    "excel_inspect_layout",
    "excel_read_table",
    "excel_read_range",
    "excel_write_csv",
}


def _tools(settings: Settings):
    app = create_server(settings)
    return {tool.name: tool.fn for tool in app._tool_manager.list_tools()}


def test_tool_list_is_stable() -> None:
    assert set(list_tool_names(create_server(Settings()))) == EXPECTED_TOOLS


def test_list_and_read_decimal_comma_csv(fixture_settings: Settings) -> None:
    tools = _tools(fixture_settings)
    listed = tools["excel_list_sheets"](path=str(FIXTURES / "decimal_comma.csv"))
    assert listed["ok"] is True
    assert listed["format"] == "csv"
    assert listed["sheets"] == [{"name": "decimal_comma", "index": 0, "state": "visible"}]
    assert listed["content_trust"] == "untrusted"

    read = tools["excel_read_table"](path=str(FIXTURES / "decimal_comma.csv"))
    assert read["ok"] is True
    assert read["decimal"] == "comma"
    assert read["header_rows"] == [3]
    assert read["rows"] == [
        {"sample": "A-1", "result": 0.125, "flag": "ok"},
        {"sample": "A-2", "result": 1250.5, "flag": "ok"},
    ]
    result = next(column for column in read["columns"] if column["name"] == "result")
    assert result["unit"] == "mg/L"
    assert read["notes"][0]["text"].startswith("Notes:")


def test_xls_mass_units(fixture_settings: Settings) -> None:
    tools = _tools(fixture_settings)
    read = tools["excel_read_table"](path=str(FIXTURES / "simple.xls"), normalize_units=True)
    assert read["ok"] is True
    assert read["format"] == "xls"
    assert read["rows"][0]["name"] == "alpha"
    assert read["rows"][0]["mass"] == 0.0015
    assert read["rows"][1]["mass"] == 0.002
    mass = next(column for column in read["columns"] if column["name"] == "mass")
    assert mass["source_unit"] == "mg"
    assert mass["unit"] == "g"


def test_xlsx_open_does_not_load_macros_or_links(fixture_settings: Settings) -> None:
    """xlsx is parsed with defusedxml. openpyxl's workbook loader, which can
    keep VBA and external links, is not on this path."""
    from mcp_advanced_excel import grid

    assert not hasattr(grid, "load_workbook")
    tools = _tools(fixture_settings)
    result = tools["excel_read_table"](path=str(FIXTURES / "instrument.xlsx"))
    assert result["ok"] is True
    assert result["rows"][0]["sample_id"] == "S-001"


def test_formula_is_not_evaluated(fixture_settings: Settings) -> None:
    tools = _tools(fixture_settings)
    read = tools["excel_read_table"](path=str(FIXTURES / "formula.xlsx"))
    assert read["ok"] is True
    assert read["formula_cells"] >= 1
    assert read["rows"][0]["sum"] is None
    assert "formulas" not in read
    assert any("no cached value" in warning for warning in read["warnings"])
    shown = tools["excel_read_table"](path=str(FIXTURES / "formula.xlsx"), include_formulas=True)
    assert shown["formulas"][0]["formula"] == "=1+1"


def test_cp1252_degree_symbol(tmp_path: Path) -> None:
    target = tmp_path / "latin.csv"
    target.write_bytes("Sample;Temp (\xb0C)\nA-1;25,5\n".encode("latin-1"))
    tools = _tools(Settings(allowed_roots=(tmp_path,)))
    read = tools["excel_read_table"](path=str(target))
    assert read["ok"] is True
    assert read["encoding"] == "cp1252"
    temp = next(column for column in read["columns"] if column["name"] == "temp")
    assert temp["unit"] == "degC"
    assert read["rows"][0]["temp"] == 25.5


def test_sensitive_token_is_scrubbed(tmp_path: Path) -> None:
    target = tmp_path / "notes.csv"
    target.write_text(
        "Sample,Note,Value\nS-1,token sk-abcdefghijklmnopqrst,1\nS-2,password: hunter2,2\n",
        encoding="utf-8",
    )
    plain = _tools(Settings(allowed_roots=(tmp_path,)))
    visible = plain["excel_read_table"](path=str(target))
    assert "sk-abcdefghijklmnopqrst" in str(visible["rows"])
    assert visible["redactions"] == 0
    tools = _tools(Settings(allowed_roots=(tmp_path,), redact="secrets"))
    read = tools["excel_read_table"](path=str(target))
    blob = str(read["rows"])
    assert "sk-abcdefghijklmnopqrst" not in blob
    assert "hunter2" not in blob
    assert "[REDACTED]" in blob
    assert read["redactions"] >= 2
    assert read["redaction_events"]


def test_file_size_limit(tmp_path: Path) -> None:
    target = tmp_path / "big.csv"
    target.write_text("Sample,Value\n" + "A,1\n" * 50, encoding="utf-8")
    tools = _tools(Settings(allowed_roots=(tmp_path,), max_file_bytes=32))
    read = tools["excel_read_table"](path=str(target))
    assert read["ok"] is False
    assert read["error"] == "file_too_large"


def test_write_csv_is_off_by_default_and_honours_output_root(tmp_path: Path) -> None:
    source = tmp_path / "in" / "plate.csv"
    source.parent.mkdir()
    source.write_text(
        "Sample;Conc. (ug/mL)\nS-1;1,5\nS-2;2,5\n",
        encoding="utf-8",
    )
    out = tmp_path / "out"
    out.mkdir()
    locked = Settings(allowed_roots=(source.parent.resolve(),), output_roots=(out.resolve(),))
    tools = _tools(locked)
    refused = tools["excel_write_csv"](path=str(source), output_path=str(out / "clean.csv"))
    assert refused["ok"] is False
    assert refused["error"] == "write_disabled"
    assert not (out / "clean.csv").exists()

    enabled = Settings(
        allowed_roots=(source.parent.resolve(),),
        output_roots=(out.resolve(),),
        allow_write=True,
    )
    tools = _tools(enabled)
    escaped = tools["excel_write_csv"](
        path=str(source),
        output_path=str(tmp_path / "escaped.csv"),
    )
    assert escaped["error"] == "outside_output_root"

    written = tools["excel_write_csv"](
        path=str(source),
        output_path=str(out / "clean.csv"),
        normalize_units=True,
    )
    assert written["ok"] is True
    body = (out / "clean.csv").read_text(encoding="utf-8")
    assert body.splitlines()[0] == "sample,conc [mg/L]"
    assert "1.5" in body
    assert "2.5" in body

    again = tools["excel_write_csv"](path=str(source), output_path=str(out / "clean.csv"))
    assert again["error"] == "output_exists"

    same = tools["excel_write_csv"](
        path=str(source),
        output_path=str(source),
        overwrite=True,
    )
    # Source lives in the read root, which is not the output root.
    assert same["ok"] is False


def test_refuses_to_overwrite_source_inside_output_root(tmp_path: Path) -> None:
    root = tmp_path / "both"
    root.mkdir()
    source = root / "plate.csv"
    source.write_text("Sample,Value\nS-1,1\n", encoding="utf-8")
    tools = _tools(
        Settings(
            allowed_roots=(root.resolve(),),
            output_roots=(root.resolve(),),
            allow_write=True,
        )
    )
    result = tools["excel_write_csv"](
        path=str(source),
        output_path=str(source),
        overwrite=True,
    )
    assert result["ok"] is False
    assert result["error"] == "invalid_path"
    assert "S-1,1" in source.read_text(encoding="utf-8")


def test_csv_formula_injection_is_prefixed(tmp_path: Path) -> None:
    root = tmp_path / "in"
    root.mkdir()
    out = tmp_path / "out"
    out.mkdir()
    source = root / "flags.csv"
    source.write_text("Sample,Value,Flag\nS-1,1,=1+1\n", encoding="utf-8")
    tools = _tools(
        Settings(
            allowed_roots=(root.resolve(),),
            output_roots=(out.resolve(),),
            allow_write=True,
        )
    )
    result = tools["excel_write_csv"](path=str(source), output_path=str(out / "safe.csv"))
    assert result["ok"] is True
    body = (out / "safe.csv").read_text(encoding="utf-8")
    assert "'=1+1" in body
