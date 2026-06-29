#!/usr/bin/env bash
# autoresearch.sh — Excel MCP Server 性能优化实验入口
#
# 被 init_experiment/run_experiment 调用。执行 benchmark.py 并产出
# autoresearch 约定格式的指标行 (METRIC / ASI), 供 harness 自动解析。
#
# 退出码: 0=成功, 非0=失败
set -euo pipefail

# autoresearch 框架已将工作目录设为项目根; 不需要 cd
# (避免 $0 为相对/纯文件名时 cd 到错误位置)

# 项目依赖 (numpy/pandas/openpyxl/calamine) 装在 Windows Python.
# autoresearch 可能在 WSL/git-bash 运行 (路径呈 /mnt/d 风格), 其原生 python 无依赖.
# 因此优先用 cmd.exe interop 调用 Windows Python; benchmark.py 自身会在 import 时
# 验证依赖, 缺失则报错退出.
if [ -n "${PYTHON:-}" ]; then
    # 用户显式指定: 直接用 (可能是 "python" 或完整路径)
    echo "[autoresearch] running perf benchmark (PYTHON=$PYTHON)..." >&2
    "$PYTHON" tools/perf-benchmark/benchmark.py
elif command -v cmd.exe >/dev/null 2>&1; then
    # Windows interop 通道 (WSL 或 git-bash): 用 cmd.exe 调 Windows python
    echo "[autoresearch] running perf benchmark via cmd.exe python..." >&2
    cmd.exe /c "python tools/perf-benchmark/benchmark.py"
else
    # 纯 Linux/WSL 环境: 用原生 python (需自行安装依赖)
    echo "[autoresearch] running perf benchmark via python3..." >&2
    python3 tools/perf-benchmark/benchmark.py
fi
