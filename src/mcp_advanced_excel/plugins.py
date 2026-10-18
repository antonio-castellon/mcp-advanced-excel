"""Extension points for a separate Pro package.

This module is the community side of the plugin boundary. It defines the
interfaces and the no-op defaults. It does not implement provenance, instrument
templates, batch processing, signing, or licence checks, and it does not
open a network connection.

A private package can register implementations with these entry-point groups
without forking this repository:

- ``mcp_advanced_excel.features`` — offline feature gate
- ``mcp_advanced_excel.provenance`` — source and cell provenance
- ``mcp_advanced_excel.templates`` — instrument template matching
- ``mcp_advanced_excel.batch`` — folder batch processing
- ``mcp_advanced_excel.signing`` — signed output

The community package registers none of those groups. A feature runs only when
the gate's ``enabled`` returns true and a provider is installed. The default
gate returns false and does not read a licence file or contact a server.
"""

from __future__ import annotations

import importlib.metadata
from dataclasses import dataclass, field
from typing import Any, Protocol

GROUP_FEATURES = "mcp_advanced_excel.features"
GROUP_PROVENANCE = "mcp_advanced_excel.provenance"
GROUP_TEMPLATES = "mcp_advanced_excel.templates"
GROUP_BATCH = "mcp_advanced_excel.batch"
GROUP_SIGNING = "mcp_advanced_excel.signing"

FEATURES = ("provenance", "templates", "batch", "signing")


@dataclass(frozen=True)
class SheetView:
    """The slice of a read that a plugin is allowed to see.

    Cell text in ``preview_rows`` is untrusted workbook content.
    """

    path: str
    sheet: str
    sheet_index: int
    header_rows: tuple[int, ...] = ()
    column_names: tuple[str, ...] = ()
    preview_rows: tuple[dict[str, Any], ...] = ()
    extra: dict[str, Any] = field(default_factory=dict)


class FeatureGate(Protocol):
    """Offline allow/deny for Pro features.

    ``enabled`` must not open a socket, fetch a URL, or phone home.
    A replacement may read a local licence file. This package does not.
    """

    def enabled(self, feature: str) -> bool:
        """Return whether ``feature`` may run in this installation."""


class ProvenanceProvider(Protocol):
    def annotate(self, view: SheetView) -> dict[str, Any]:
        """Return JSON fields to attach, or an empty dict."""


class TemplateProvider(Protocol):
    def match(self, view: SheetView) -> dict[str, Any] | None:
        """Return template hints such as ``header_row``, or None."""


class BatchProvider(Protocol):
    def run(self, folder: str, options: dict[str, Any]) -> dict[str, Any]:
        """Process a folder of workbooks. Not called unless the feature is active."""


class SigningProvider(Protocol):
    def sign(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Return ``payload`` plus any signature fields."""


class NullFeatureGate:
    """Community default. Every Pro feature stays off. No I/O."""

    def enabled(self, feature: str) -> bool:
        return False


class NullProvenance:
    def annotate(self, view: SheetView) -> dict[str, Any]:
        return {}


class NullTemplates:
    def match(self, view: SheetView) -> dict[str, Any] | None:
        return None


class NullBatch:
    def run(self, folder: str, options: dict[str, Any]) -> dict[str, Any]:
        return {
            "ok": False,
            "error": "not_available",
            "message": "Batch processing is not part of the community edition.",
        }


class NullSigner:
    def sign(self, payload: dict[str, Any]) -> dict[str, Any]:
        return payload


def _entry_points(group: str) -> list[importlib.metadata.EntryPoint]:
    points = importlib.metadata.entry_points()
    selected = points.select(group=group)
    return list(selected)


def _load_first(group: str, default: Any) -> Any:
    found = _entry_points(group)
    if not found:
        return default
    loaded = found[0].load()
    if isinstance(loaded, type):
        return loaded()
    return loaded


class PluginHost:
    """Resolved providers. Defaults are no-ops and the gate is closed."""

    def __init__(self) -> None:
        self.gate: FeatureGate = _load_first(GROUP_FEATURES, NullFeatureGate())
        self.provenance: ProvenanceProvider = _load_first(GROUP_PROVENANCE, NullProvenance())
        self.templates: TemplateProvider = _load_first(GROUP_TEMPLATES, NullTemplates())
        self.batch: BatchProvider = _load_first(GROUP_BATCH, NullBatch())
        self.signer: SigningProvider = _load_first(GROUP_SIGNING, NullSigner())

    def using_defaults(self) -> bool:
        return (
            isinstance(self.gate, NullFeatureGate)
            and isinstance(self.provenance, NullProvenance)
            and isinstance(self.templates, NullTemplates)
            and isinstance(self.batch, NullBatch)
            and isinstance(self.signer, NullSigner)
        )

    def feature_active(self, feature: str) -> bool:
        """True only when the offline gate allows it and a provider is installed."""
        if feature not in FEATURES:
            return False
        if not self.gate.enabled(feature):
            return False
        provider = {
            "provenance": self.provenance,
            "templates": self.templates,
            "batch": self.batch,
            "signing": self.signer,
        }[feature]
        return not isinstance(
            provider, (NullProvenance, NullTemplates, NullBatch, NullSigner, NullFeatureGate)
        )

    def template_header_row(self, view: SheetView) -> int | None:
        if not self.feature_active("templates"):
            return None
        hint = self.templates.match(view)
        if not isinstance(hint, dict):
            return None
        header = hint.get("header_row")
        if isinstance(header, int) and header >= 1:
            return header
        return None

    def enrich(self, payload: dict[str, Any], view: SheetView) -> dict[str, Any]:
        """Apply active Pro hooks. The community default returns ``payload`` unchanged."""
        if self.feature_active("templates"):
            hint = self.templates.match(view)
            if isinstance(hint, dict) and hint:
                payload = {**payload, "template": hint}
        if self.feature_active("provenance"):
            extra = self.provenance.annotate(view)
            if isinstance(extra, dict) and extra:
                payload = {**payload, "provenance": extra}
        if self.feature_active("signing"):
            signed = self.signer.sign(payload)
            if isinstance(signed, dict):
                return signed
        return payload


_HOST: PluginHost | None = None


def get_plugin_host() -> PluginHost:
    global _HOST
    if _HOST is None:
        _HOST = PluginHost()
    return _HOST


def reset_plugin_host() -> PluginHost:
    """Drop the cached host. Tests use this after changing entry points."""
    global _HOST
    _HOST = PluginHost()
    return _HOST
