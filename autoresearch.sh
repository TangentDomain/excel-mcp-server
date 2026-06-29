#!/usr/bin/env bash
# autoresearch.sh — ExcelMCP SQL 准确率差分测试基准入口
#
# 主指标: accuracy (ExcelMCP SQL 引擎 vs SQLite oracle 对齐率)
# 确定性: 固定 fixture、固定 SQL 集，无网络/时间依赖。
#
set -euo pipefail

cd "$(dirname "$0")"

# 选择 Python 解释器（优先级: uv run > .venv > python3 > python）
if command -v uv >/dev/null 2>&1; then
    PY="uv run python"
elif [ -x ".venv/Scripts/python.exe" ]; then
    PY=".venv/Scripts/python.exe"
elif command -v python3 >/dev/null 2>&1; then
    PY="python3"
else
    PY="python"
fi

# 运行差分测试基准
$PY tools/sql-accuracy-benchmark.py "$@"
