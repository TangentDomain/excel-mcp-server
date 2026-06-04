#!/bin/bash
# scripts/setup-hooks.sh
# 将 pre-commit 检查脚本链接到 .git/hooks/pre-commit

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
HOOK_DIR="$(git rev-parse --git-common-dir)/hooks"
HOOK_FILE="$HOOK_DIR/pre-commit"

if [ -L "$HOOK_FILE" ]; then
    echo "pre-commit hook 已存在（符号链接），跳过"
    exit 0
fi

if [ -f "$HOOK_FILE" ]; then
    echo "警告: $HOOK_FILE 已存在且非符号链接，备份为 .bak"
    mv "$HOOK_FILE" "$HOOK_FILE.bak"
fi

mkdir -p "$HOOK_DIR"
ln -sf "$SCRIPT_DIR/pre_commit_check.sh" "$HOOK_FILE"
chmod +x "$HOOK_FILE"
echo "pre-commit hook 已安装: $HOOK_FILE → $SCRIPT_DIR/pre_commit_check.sh"
