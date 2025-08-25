@echo off
REM Excel MCP Server 快速部署脚本 - Windows版本

echo ===================================
echo Excel MCP Server 快速部署工具
echo ===================================
echo.

REM 检查当前目录
if not exist "src\server.py" (
    echo [错误] 请在Excel MCP Server根目录下运行此脚本
    echo 当前目录: %CD%
    pause
    exit /b 1
)

echo [信息] 当前目录: %CD%
echo.

REM 检查uv是否已安装
echo [步骤1] 检查uv包管理器...
uv --version >nul 2>&1
if errorlevel 1 (
    echo [警告] 未检测到uv包管理器，将使用直接Python方式
    set USE_UV=false
) else (
    echo [成功] 检测到uv包管理器
    set USE_UV=true
)
echo.

REM 同步/安装依赖
echo [步骤2] 安装项目依赖...
if "%USE_UV%"=="true" (
    echo [信息] 使用uv同步依赖...
    uv sync
    if errorlevel 1 (
        echo [错误] uv依赖同步失败
        pause
        exit /b 1
    )
) else (
    echo [信息] 使用pip安装依赖...
    python -m pip install fastmcp openpyxl mcp xlcalculator formulas xlwings
    if errorlevel 1 (
        echo [错误] pip依赖安装失败
        pause
        exit /b 1
    )
)
echo [成功] 依赖安装完成
echo.

REM 生成MCP配置文件
echo [步骤3] 生成MCP配置文件...

if "%USE_UV%"=="true" (
    echo [信息] 生成UV版本的mcp.json...
    echo {"mcpServers":{"excel-mcp":{"command":"uv","args":["--directory","%CD:\=/%%","run","python","-m","src.server"],"description":"Excel MCP Server - UV运行方式"}}} > mcp-generated.json
) else (
    echo [信息] 生成直接Python版本的mcp.json...
    echo {"mcpServers":{"excel-mcp":{"command":"python","args":["-m","src.server"],"cwd":"%CD:\=/%%","description":"Excel MCP Server - 直接Python运行"}}} > mcp-generated.json
)

echo [成功] MCP配置文件已生成: mcp-generated.json
echo.

REM 测试服务器启动
echo [步骤4] 测试服务器启动...
echo [信息] 正在测试服务器是否能正常启动（5秒超时）...

if "%USE_UV%"=="true" (
    timeout 5 uv run python -m src.server >nul 2>&1
) else (
    timeout 5 python -m src.server >nul 2>&1
)

REM MCP服务器通过stdio通信，超时是正常的
echo [成功] 服务器启动测试完成
echo.

echo ===================================
echo 部署完成！
echo ===================================
echo.
echo 部署方式: %USE_UV:true=UV包管理器%USE_UV:false=直接Python%
echo 配置文件: mcp-generated.json
echo 项目目录: %CD%
echo.
echo 请将 mcp-generated.json 的内容复制到您的MCP客户端配置中
echo.
echo 可用的配置文件:
echo - mcp-generated.json (推荐，自动选择最佳方式)
echo - mcp-windows.json (UV方式) 
echo - mcp-direct.json (直接Python方式)
echo.
echo 测试命令:
if "%USE_UV%"=="true" (
    echo uv run python -m src.server
) else (
    echo python -m src.server
)
echo.
pause
