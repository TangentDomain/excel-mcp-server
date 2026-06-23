#!/usr/bin/env python3
"""ExcelMCP CLI — 开发模式入口。pip install 后使用 excel-cli 命令。"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
from excel_mcp_server_fastmcp.cli import main
import sys
sys.exit(main())
