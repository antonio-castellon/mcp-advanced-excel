"""Packaging markers the registry and a future PyPI publish depend on."""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_readme_is_ascii_and_carries_mcp_name() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    readme.encode("ascii")
    assert "<!-- mcp-name: io.github.antonio-castellon/mcp-advanced-excel -->" in readme
    for doc in (ROOT / "docs").glob("*.md"):
        doc.read_text(encoding="utf-8").encode("ascii")


def test_server_json_matches_registry_shape() -> None:
    payload = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    assert payload["name"] == "io.github.antonio-castellon/mcp-advanced-excel"
    assert payload["version"] == "0.1.0"
    assert len(payload["description"]) <= 100
    package = payload["packages"][0]
    assert package["registryType"] == "pypi"
    assert package["identifier"] == "mcp-advanced-excel"
    assert package["version"] == "0.1.0"
    assert package["transport"]["type"] == "stdio"


def test_pyproject_pins_mcp_above_the_streamable_http_advisory() -> None:
    text = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    assert 'name = "mcp-advanced-excel"' in text
    assert 'version = "0.1.0"' in text
    assert "mcp>=1.30.0,<2" in text
    assert "mcp-advanced-excel = " in text
