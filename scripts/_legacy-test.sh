#!/bin/bash
echo "Testing MCP server..."
echo '{"jsonrpc":"2.0","id":1,"method":"ping"}' | uv run excel-mcp-server-fastmcp --stdio 2>&1 | head -5
