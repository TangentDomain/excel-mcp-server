"""Excel MCP Server package - game dev focused Excel configuration table MCP server."""

__version__ = "1.17.0"

# main 通过 __getattr__ 延迟 import，避免 import 包时触发 server.py 的重依赖链

__all__ = ["main", "__version__"]

def __getattr__(name):
    if name == "main":
        from .server import main
        return main
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
