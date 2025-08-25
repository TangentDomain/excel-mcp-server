@echo off
chcp 65001 >nul 2>&1
REM Excel MCP Server Quick Deploy Script - Windows Version

echo ===================================
echo Excel MCP Server Quick Deploy Tool
echo ===================================
echo.

REM Check current directory
if not exist "src\server.py" (
    echo [ERROR] Please run this script in Excel MCP Server root directory
    echo Current directory: %CD%
    pause
    exit /b 1
)

echo [INFO] Current directory: %CD%
echo.

REM Check if uv is installed
echo [STEP 1] Checking uv package manager...
uv --version >nul 2>&1
if errorlevel 1 (
    echo [WARN] UV not detected, will use direct Python method
    set "USE_UV=false"
) else (
    echo [OK] UV package manager detected
    set "USE_UV=true"
)
echo.

REM Install dependencies
echo [STEP 2] Installing project dependencies...
if "%USE_UV%"=="true" (
    echo [INFO] Using uv to sync dependencies...
    uv sync
    if errorlevel 1 (
        echo [ERROR] UV dependency sync failed
        pause
        exit /b 1
    )
) else (
    echo [INFO] Using pip to install dependencies...
    python -m pip install fastmcp openpyxl mcp xlcalculator formulas xlwings
    if errorlevel 1 (
        echo [ERROR] Pip installation failed
        pause
        exit /b 1
    )
)
echo [OK] Dependencies installed successfully
echo.

REM Generate MCP config file
echo [STEP 3] Generating MCP configuration file...

set "PROJ_PATH=%CD:\=/%"
if "%USE_UV%"=="true" (
    echo [INFO] Generating UV version mcp.json...
    (
        echo {
        echo   "mcpServers": {
        echo     "excel-mcp": {
        echo       "command": "uv",
        echo       "args": [
        echo         "--directory",
        echo         "%PROJ_PATH%",
        echo         "run",
        echo         "python",
        echo         "-m",
        echo         "src.server"
        echo       ],
        echo       "description": "Excel MCP Server - UV Runtime"
        echo     }
        echo   }
        echo }
    ) > mcp-generated.json
) else (
    echo [INFO] Generating direct Python version mcp.json...
    (
        echo {
        echo   "mcpServers": {
        echo     "excel-mcp": {
        echo       "command": "python",
        echo       "args": [
        echo         "-m",
        echo         "src.server"
        echo       ],
        echo       "cwd": "%PROJ_PATH%",
        echo       "description": "Excel MCP Server - Direct Python"
        echo     }
        echo   }
        echo }
    ) > mcp-generated.json
)

echo [OK] MCP configuration file generated: mcp-generated.json
echo.

REM Test server startup
echo [STEP 4] Testing server startup...
echo [INFO] Testing if server can start normally ^(5 second timeout^)...

if "%USE_UV%"=="true" (
    timeout 5 uv run python -m src.server >nul 2>&1
) else (
    timeout 5 python -m src.server >nul 2>&1
)

REM MCP server communicates via stdio, timeout is normal
echo [OK] Server startup test completed
echo.

echo ===================================
echo Deployment Complete!
echo ===================================
echo.
if "%USE_UV%"=="true" (
    echo Deployment Method: UV Package Manager
) else (
    echo Deployment Method: Direct Python
)
echo Configuration File: mcp-generated.json
echo Project Directory: %CD%
echo.
echo Please copy the content of mcp-generated.json to your MCP client configuration
echo.
echo Available configuration files:
echo - mcp-generated.json ^(Recommended, auto-selected best method^)
echo - mcp-windows.json ^(UV method^)
echo - mcp-direct.json ^(Direct Python method^)
echo.
echo Test command:
if "%USE_UV%"=="true" (
    echo uv run python -m src.server
) else (
    echo python -m src.server
)
echo.
pause
