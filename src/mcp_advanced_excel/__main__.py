"""python -m mcp_advanced_excel — start the stdio MCP server."""

from __future__ import annotations


def main() -> None:
    from .server import main as run

    run()


if __name__ == "__main__":
    main()
