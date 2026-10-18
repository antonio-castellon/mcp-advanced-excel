"""Packaging markers the registry and a future PyPI publish depend on."""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_readme_is_ascii_and_carries_mcp_name() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    readme.encode("ascii")
    prose = " ".join(readme.split())
    assert "<!-- mcp-name: io.github.antonio-castellon/mcp-advanced-excel -->" in readme
    assert "Available on PyPI" not in readme
    assert "mailto:antonio@castellon.ch" in readme
    assert "https://antonio-castellon.github.io" not in readme
    assert "https://github.com/antonio-castellon/mcp-glp-audit-proxy" in readme
    assert "mcp-glp-audit-proxy-pro" not in readme.lower()
    assert "designed to support GLP / 21 CFR Part 11 / EU GMP Annex 11 workflows" in prose
    assert "certified" not in readme.lower()
    assert "compliant" not in readme.lower()
    assert "Agensi" not in readme
    for doc in (ROOT / "docs").glob("*.md"):
        text = doc.read_text(encoding="utf-8")
        text.encode("ascii")
        assert "https://antonio-castellon.github.io" not in text
        assert "certified" not in text.lower()
        assert "compliant" not in text.lower()
        assert "mcp-glp-audit-proxy-pro" not in text.lower()
        if doc.name == "PLUGINS.md":
            plugins = " ".join(text.split())
            assert "mailto:antonio@castellon.ch" in text
            assert "https://github.com/antonio-castellon/mcp-glp-audit-proxy" in text
            assert "designed to support GLP / 21 CFR Part 11 / EU GMP Annex 11 workflows" in plugins


def test_server_json_matches_registry_shape() -> None:
    payload = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    assert payload["name"] == "io.github.antonio-castellon/mcp-advanced-excel"
    assert payload["version"] == "0.2.0"
    assert len(payload["description"]) <= 100
    package = payload["packages"][0]
    assert package["registryType"] == "pypi"
    assert package["identifier"] == "mcp-advanced-excel"
    assert package["version"] == "0.2.0"
    assert package["transport"]["type"] == "stdio"


def test_pyproject_pins_mcp_above_the_streamable_http_advisory() -> None:
    text = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    assert 'name = "mcp-advanced-excel"' in text
    assert 'version = "0.2.0"' in text
    assert "defusedxml" in text
    assert "mcp>=1.30.0,<2" in text
    assert "mcp-advanced-excel = " in text
