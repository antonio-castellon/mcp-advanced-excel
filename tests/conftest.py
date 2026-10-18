"""Shared paths for the synthetic fixtures."""

from __future__ import annotations

from pathlib import Path

import pytest

from mcp_advanced_excel.config import Settings

FIXTURES = Path(__file__).resolve().parent / "fixtures"


@pytest.fixture
def fixture_settings() -> Settings:
    return Settings(allowed_roots=(FIXTURES,))
