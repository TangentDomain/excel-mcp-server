#!/bin/bash
# scripts/pre_commit_check.sh
# Pre-commit validation checks for the excel-mcp-server project

set -e

echo "执行pre-commit检查..."


# 检查README语言
if ! head -n 3 README.md | tail -n 1 | grep -q "^# 🎮 ExcelMCP"; then
    echo "ERROR: README.md 首行语言检查失败"
    exit 1
fi

if ! head -n 3 README.en.md | tail -n 1 | grep -q "^# 🎮 ExcelMCP"; then
    echo "ERROR: README.en.md 首行语言检查失败"
    exit 1
fi

# 检查冲突标记（排除归档目录和脚本目录和运行记录）
if find . -name "*.md" -not -path "./docs/archive/*" -not -path "./scripts/*" -not -path "./runs/*" | xargs grep -l "<<<<<<<" 2>/dev/null; then
    echo "ERROR: 发现未解决的合并冲突"
    exit 1
fi

# 检查敏感信息（排除二进制、测试数据、脚本、归档、.venv、.github）
if find . -name "*.py" -o -name "*.md" -o -name "*.yml" -o -name "*.yaml" -o -name "*.json" -o -name "*.toml" | \
   grep -v "__pycache__" | \
   grep -v "test_data" | \
   grep -v "dist" | \
   grep -v ".venv" | \
   grep -v ".github" | \
   grep -v "docs/archive" | \
   grep -v "scripts/pre_commit_check.sh" | \
   xargs grep -l "AK.*=\|LTAI.*=\|ft522.*=" 2>/dev/null; then
    echo "ERROR: 发现可能的敏感信息"
    exit 1
fi

# 检查docstring契约（REQ-029）
echo "检查docstring契约..."
if ! python3 scripts/lint_docstring_contract.py --quiet; then
    echo "ERROR: Docstring契约验证失败，请运行 'python3 scripts/lint_docstring_contract.py' 查看详情"
    exit 1
fi
# Ruff 格式检查
echo "检查 ruff format..."
if command -v ruff &>/dev/null; then
    ruff format --check src/ tests/ || { echo "ERROR: ruff format 检查失败，请运行 'ruff format src/ tests/'"; exit 1; }
    ruff check src/ tests/ || { echo "ERROR: ruff check 检查失败，请运行 'ruff check src/ tests/ --fix'"; exit 1; }
fi

# 不变量测试
echo "运行不变量测试..."
uv run python -m pytest tests/invariants/ -q --tb=short --timeout=30 || { echo "ERROR: 不变量测试失败"; exit 1; }

echo "Pre-commit检查通过"
exit 0
