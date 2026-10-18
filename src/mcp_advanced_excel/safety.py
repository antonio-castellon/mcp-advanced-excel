"""Scrub obviously sensitive tokens from cell text before it leaves the process."""

from __future__ import annotations

import re

# Token shapes copied in spirit from mcp-telegram-bridge safety.scrub_outbound.
# Lab exports occasionally paste a credential into a note or operator field.
_TOKENISH = re.compile(
    r"""(?x)
    (?:
        gh[pousr]_[A-Za-z0-9_]{20,}
        | xox[baprs]-[A-Za-z0-9-]{10,}
        | whsec_[A-Za-z0-9+/=_-]{16,}
        | crsr_[A-Za-z0-9_]{16,}
        | sk-[A-Za-z0-9]{20,}
        | AKIA[0-9A-Z]{16}
        | \d{8,12}:[A-Za-z0-9_-]{30,}
    )
    """
)

_SECRET_ASSIGN = re.compile(
    r"""(?ix)
    (?<![A-Za-z0-9_])
    ((?:export\s+)?[A-Z][A-Z0-9_]*(?:TOKEN|KEY|SECRET|PASSWORD|PASSWD|CREDENTIAL)[A-Z0-9_]*)
    \s*=\s*([\"']?)([^\s\"']+)\2
    """
)

_LABELED_SECRET = re.compile(
    r"""(?ix)
    \b(password|passwd|secret|api[_\s-]?key|access[_\s-]?token|bot\s*token)
    \b\s*[:=]\s*\S+
    """
)

_SSN = re.compile(r"\b\d{3}-\d{2}-\d{4}\b")


def scrub_text(text: str) -> tuple[str, int]:
    """Return scrubbed text and how many replacements were made."""
    if not text:
        return text, 0
    count = 0

    def _sub(pattern: re.Pattern[str], repl: str, value: str) -> str:
        nonlocal count

        def _one(_match: re.Match[str]) -> str:
            nonlocal count
            count += 1
            return repl

        return pattern.sub(_one, value)

    scrubbed = _sub(_TOKENISH, "[REDACTED]", text)
    scrubbed = _sub(_SECRET_ASSIGN, r"\1=[REDACTED]", scrubbed)
    scrubbed = _sub(_LABELED_SECRET, "[REDACTED]", scrubbed)
    scrubbed = _sub(_SSN, "[REDACTED-ID]", scrubbed)
    return scrubbed, count
