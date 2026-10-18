"""Optional scrubbing of token-shaped secrets in cell text.

Redaction is off unless the caller asks for it. Lab identifiers such as
``123-45-6789`` are not treated as secrets.
"""

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
    \b(\s*[:=]\s*)(\S+)
    """
)


def scrub_text(text: str) -> tuple[str, list[str]]:
    """Return scrubbed text and the kinds of replacement that were made.

    Kinds are ``token``, ``assignment``, and ``labeled``. The secret itself
    is not returned.
    """
    if not text:
        return text, []
    kinds: list[str] = []

    def apply(pattern: re.Pattern[str], kind: str, repl) -> None:
        nonlocal text

        def one(match: re.Match[str]) -> str:
            replaced = repl(match) if callable(repl) else repl
            if replaced != match.group(0):
                kinds.append(kind)
            return replaced

        text = pattern.sub(one, text)

    apply(_TOKENISH, "token", "[REDACTED]")
    apply(_SECRET_ASSIGN, "assignment", lambda match: f"{match.group(1)}=[REDACTED]")
    apply(
        _LABELED_SECRET,
        "labeled",
        lambda match: (
            match.group(0)
            if match.group(3) == "[REDACTED]"
            else f"{match.group(1)}{match.group(2)}[REDACTED]"
        ),
    )
    return text, kinds
