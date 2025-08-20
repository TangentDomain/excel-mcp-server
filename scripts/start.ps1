# Excel MCP Server 快速启动脚本
# 用于快速启动 Excel MCP 服务器

Write-Host "🚀 启动 Excel MCP Server (FastMCP)..." -ForegroundColor Green

# 检查虚拟环境
if (-not (Test-Path "venv")) {
    Write-Host "❌ 虚拟环境不存在，请先运行 setup.ps1 -Deploy" -ForegroundColor Red
    exit 1
}

# 激活虚拟环境并启动服务器
Write-Host "🔧 激活虚拟环境..." -ForegroundColor Yellow
& ".\venv\Scripts\activate.ps1"

Write-Host "📊 启动 Excel MCP 服务器..." -ForegroundColor Yellow
Write-Host "💡 服务器将在 stdio 模式下运行，等待 MCP 客户端连接" -ForegroundColor Cyan
Write-Host "🔗 在 Claude Desktop 中配置 excel-mcp-config.json 即可使用" -ForegroundColor Cyan
Write-Host ""

python server.py
