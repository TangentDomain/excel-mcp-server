@echo off
REM Excel MCP Server - 监控脚本启动器
REM
REM 使用方法:
REM run-monitor.bat          - 运行快速监控
REM run-monitor.bat full     - 运行完整监控
REM run-monitor.bat help     - 显示帮助

setlocal enabledelayedexpansion

echo.
echo ==========================================
echo  Excel MCP Server 监控工具
echo ==========================================
echo.

REM 检查Python环境
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未找到Python环境
    echo 请确保Python已安装并添加到PATH中
    pause
    exit /b 1
)

REM 获取脚本目录
set SCRIPT_DIR=%~dp0
set PROJECT_DIR=%SCRIPT_DIR%..

REM 切换到项目目录
cd /d "%PROJECT_DIR%"

REM 检查是否在正确的目录
if not exist "src\server.py" (
    echo ❌ 错误: 请在Excel MCP Server项目根目录下运行此脚本
    pause
    exit /b 1
)

REM 创建必要的目录
if not exist "reports" mkdir reports
if not exist "logs" mkdir logs

REM 处理命令行参数
set MODE=%1
if "%MODE%"=="" set MODE=quick

if "%MODE%"=="help" (
    echo.
    echo 使用方法:
    echo   run-monitor.bat          - 运行快速监控 (默认)
    echo   run-monitor.bat full     - 运行完整监控
    echo   run-monitor.bat quick    - 运行快速监控
    echo   run-monitor.bat help     - 显示此帮助信息
    echo.
    echo 监控选项:
    echo   quick  - 快速监控 (约1-2分钟)
    echo   full   - 完整监控 (约5-10分钟)
    echo.
    pause
    exit /b 0
)

if "%MODE%"=="quick" (
    echo 🚀 启动快速监控...
    echo.
    python "%SCRIPT_DIR%quick-monitor.py"
    if errorlevel 1 (
        echo.
        echo ❌ 快速监控失败
        pause
        exit /b 1
    )
) else if "%MODE%"=="full" (
    echo 🔍 启动完整监控...
    echo 注意: 完整监控需要较长时间 (5-10分钟)
    echo.
    python "%SCRIPT_DIR%monitor-and-maintain.py"
    if errorlevel 1 (
        echo.
        echo ❌ 完整监控失败
        pause
        exit /b 1
    )
) else (
    echo ❌ 未知模式: %MODE%
    echo 运行 'run-monitor.bat help' 查看帮助
    pause
    exit /b 1
)

echo.
echo ✅ 监控完成！
echo.
echo 查看报告:
echo   - HTML报告: reports\monitoring-report-*.html
echo   - JSON数据: reports\quick-monitor-*.json
echo   - 日志文件: logs\monitor.log
echo.

pause