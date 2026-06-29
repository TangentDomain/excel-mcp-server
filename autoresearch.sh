#!/usr/bin/env bash
# autoresearch.sh — Excel MCP Server 准确率/正确性优化实验入口
#
# 被 init_experiment/run_experiment 调用。执行 accuracy-benchmark/benchmark.py
# 并产出 autoresearch 约定格式的指标行 (METRIC / ASI), 供 harness 自动解析。
#
# 退出码: 0=成功, 非0=失败
set -eu

# autoresearch 框架已将工作目录设为项目根; 不需要 cd

# 项目依赖装在 Windows Python. cmd.exe interop 是可靠通道.
if [ -n "${PYTHON:-}" ]; then
    echo "[autoresearch] running accuracy benchmark (PYTHON=$PYTHON)..." >&2
    "$PYTHON" tools/accuracy-benchmark/benchmark.py
elif command -v cmd.exe >/dev/null 2>&1; then
    echo "[autoresearch] running accuracy benchmark via cmd.exe python..." >&2
    cmd.exe /c "python tools/accuracy-benchmark/benchmark.py"
else
    echo "[autoresearch] running accuracy benchmark via python3..." >&2
    python3 tools/accuracy-benchmark/benchmark.py
fi
