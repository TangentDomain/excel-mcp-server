#!/bin/bash
# scripts/pre_commit_check.sh

set -e

echo "执行pre-commit检查..."

# 检查docs/RULES.md存在且非空
if [ ! -f "docs/RULES.md" ] || [ ! -s "docs/RULES.md" ]; then
    echo "ERROR: docs/RULES.md 不存在或为空"
    exit 1
fi

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

# 检查敏感信息（排除二进制文件、测试数据、脚本和归档）
if find . -name "*.py" -o -name "*.md" -o -name "*.yml" -o -name "*.yaml" -o -name "*.json" -o -name "*.toml" | \
   grep -v "__pycache__" | \
   grep -v "test_data" | \
   grep -v "dist" | \
   grep -v "docs/archive" | \
   grep -v "scripts/pre_commit_check.sh" | \
   xargs grep -l "AK.*=\|LTAI.*=\|ft522.*=\|admin.*=\|secret.*=" 2>/dev/null; then
    echo "ERROR: 发现可能的敏感信息"
    exit 1
fi

echo "Pre-commit检查通过"
exit 0