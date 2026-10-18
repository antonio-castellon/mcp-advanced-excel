"""Allowlist, symlink escape, and path traversal."""

from __future__ import annotations

from pathlib import Path

import pytest

from mcp_advanced_excel.config import Settings
from mcp_advanced_excel.errors import ToolFailure
from mcp_advanced_excel.paths import resolve_read_file, resolve_write_file
from mcp_advanced_excel.server import create_server


def _server(settings: Settings):
    app = create_server(settings)
    return {tool.name: tool.fn for tool in app._tool_manager.list_tools()}


def test_empty_allowlist_refuses_every_path(tmp_path: Path) -> None:
    target = tmp_path / "data.csv"
    target.write_text("a\n1\n", encoding="utf-8")
    with pytest.raises(ToolFailure) as caught:
        resolve_read_file(str(target), ())
    assert caught.value.code == "roots_unset"


def test_path_outside_root_is_rejected(tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    outside = tmp_path / "secret.csv"
    outside.write_text("hidden\n", encoding="utf-8")
    tools = _server(Settings(allowed_roots=(root.resolve(),)))
    result = tools["excel_list_sheets"](path=str(outside))
    assert result["ok"] is False
    assert result["error"] == "outside_allowlist"


def test_parent_traversal_is_rejected(tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    outside = tmp_path / "secret.csv"
    outside.write_text("Sample,Value\nA,1\n", encoding="utf-8")
    traversed = root / ".." / "secret.csv"
    tools = _server(Settings(allowed_roots=(root.resolve(),)))
    result = tools["excel_read_table"](path=str(traversed))
    assert result["ok"] is False
    assert result["error"] == "outside_allowlist"
    assert "hidden" not in result["message"]
    assert "Sample" not in json_dump(result)


def test_symlink_escape_is_rejected_and_internal_symlink_is_allowed(tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    outside = tmp_path / "secret.csv"
    outside.write_text("secret\n", encoding="utf-8")
    escape = root / "escape.csv"
    escape.symlink_to(outside)
    inside = root / "inside.csv"
    inside.write_text("Sample,Value\nA,1\n", encoding="utf-8")
    linked = root / "linked.csv"
    linked.symlink_to(inside)

    tools = _server(Settings(allowed_roots=(root.resolve(),)))
    denied = tools["excel_list_sheets"](path=str(escape))
    assert denied["ok"] is False
    assert denied["error"] == "outside_allowlist"

    allowed = tools["excel_read_table"](path=str(linked))
    assert allowed["ok"] is True
    assert allowed["rows"] == [{"sample": "A", "value": 1}]


def test_url_and_null_byte_are_rejected(tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    settings = Settings(allowed_roots=(root.resolve(),))
    tools = _server(settings)
    url = tools["excel_list_sheets"](path="https://example.com/sheet.xlsx")
    assert url["error"] == "invalid_path"
    nul = tools["excel_list_sheets"](path=str(root / "a.csv") + "\x00.xlsx")
    assert nul["ok"] is False


def test_write_target_cannot_escape_output_root(tmp_path: Path) -> None:
    out = tmp_path / "out"
    out.mkdir()
    escaped = out / ".." / "outside.csv"
    with pytest.raises(ToolFailure) as caught:
        resolve_write_file(str(escaped), (out.resolve(),), overwrite=False)
    assert caught.value.code == "outside_output_root"


def json_dump(payload: dict) -> str:
    return str(payload)
