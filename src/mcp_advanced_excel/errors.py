"""Expected failures returned to the MCP client as structured errors."""

from __future__ import annotations


class ToolFailure(Exception):
    """A request the server refuses or cannot complete.

    The message is safe to show to the host. It must not include file contents
    or paths that sit outside the allowlist.
    """

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code
