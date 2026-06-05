#!/bin/bash
echo "🔍 MCP服务器连接诊断"
echo "=========================="
echo ""
echo "1️⃣ 检查uv安装..."
which uv && uv --version || echo "❌ uv未安装"
echo ""
echo "2️⃣ 检查虚拟环境..."
ls -la .venv/Scripts/excel-mcp-server-fastmcp.exe && echo "✅ exe存在" || echo "❌ exe不存在"
echo ""
echo "3️⃣ 测试MCP服务器..."
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}' | uv run --directory "D:\excel-mcp-server" excel-mcp-server-fastmcp --stdio 2>&1 | head -1
echo ""
echo "✅ 诊断完成！"
