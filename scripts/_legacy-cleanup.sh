#!/bin/bash
# 递归自底向上删除空目录
# 用法: ./scripts/cleanup_empty_dirs.sh [目录路径]

set -e

# 默认清理当前目录，可指定路径
TARGET_DIR="${1:-.}"

# 排除的目录（不检查和删除）
EXCLUDE_DIRS=".git .venv node_modules __pycache__ .pytest_cache .mypy_cache htmlcov"

echo "🔍 开始扫描空目录: $TARGET_DIR"
echo "🚫 排除目录: $EXCLUDE_DIRS"
echo ""

# 构建排除参数
exclude_args=""
for dir in $EXCLUDE_DIRS; do
    exclude_args="$exclude_args -name $dir -prune -o"
done

# 自底向上查找并删除空目录
# depth 0-15 控制搜索深度
# sort -r 反转排序实现自底向上
deleted_count=0

while IFS= read -r dir; do
    if [ -d "$dir" ] && [ -z "$(ls -A "$dir" 2>/dev/null)" ]; then
        echo "🗑️  删除空目录: $dir"
        rmdir "$dir" 2>/dev/null && ((deleted_count++))
    fi
done < <(find "$TARGET_DIR" -mindepth 1 -maxdepth 15 \( $exclude_args -type d -print \) | sort -rz)

echo ""
echo "✅ 清理完成！共删除 $deleted_count 个空目录"
