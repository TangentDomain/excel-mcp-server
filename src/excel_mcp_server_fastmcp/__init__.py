"""Excel MCP Server package - game dev focused Excel configuration table MCP server."""

__version__ = "1.11.0"

# Import main function for CLI entry point
from .server import main

__all__ = ["main", "__version__"]
