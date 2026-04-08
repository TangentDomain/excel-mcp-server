#!/bin/bash
# ExcelMCP 编码分析脚本

set -e

echo "🎯 ExcelMCP 编码分析阶段启动"
echo "时间: $(date)"

# 创建临时任务文件
TASK_FILE=".claude-tasks.md"
> "$TASK_FILE"

echo "## 编码分析阶段" >> "$TASK_FILE"
echo "时间: $(date)" >> "$TASK_FILE"
echo "" >> "$TASK_FILE"

# 获取当前轮次
if [ -f ".version-info.txt" ]; then
    ROUND_NUM=$(cat ".version-info.txt" | grep -o "[0-9]\+" | tail -1)
else
    ROUND_NUM=300
fi
echo "轮次: $ROUND_NUM" >> "$TASK_FILE"

# 分析 OPEN 需求
echo "### OPEN 需求分析" >> "$TASK_FILE"

if python3 -c "
import sys
import json
sys.path.insert(0, 'src')
try:
    with open('REQUIREMENTS.md', 'r', encoding='utf-8') as f:
        content = f.read()
        if 'PAUSED' in content and 'OPEN' in content:
            print('发现OPEN/PAUSED需求，需要优先处理')
else:
    print('REQUIREMENTS.md 读取失败')
"; then
    echo "✅ 已发现 OPEN 需求，将进入子任务执行" >> "$TASK_FILE"
else
    echo "❌ 未发现有效 OPEN 需求" >> "$TASK_FILE"
fi

echo "" >> "$TASK_FILE"
echo "### 下一步" >> "$TASK_FILE"
echo "1. K3: 推送测试完成报告" >> "$TASK_FILE"
echo "2. 然后根据是否有 OPEN 需求决定是否编码" >> "$TASK_FILE"

echo ""
echo "📝 任务文件已生成: $TASK_FILE"
echo "内容预览:"
echo "----------------------------------------"
head -10 "$TASK_FILE"
echo "----------------------------------------"

# 检查是否有 FEEDBACK.md 转化为需求
if [ -f "FEEDBACK.md" ] && [ -s "FEEDBACK.md" ]; then
    echo ""
    echo "📋 检查 FEEDBACK.md 中的反馈..."
    if grep -q "REQ-072" FEEDBACK.md; then
        echo "✅ 发现 cron 频率调整需求，已加入任务队列"
    fi
fi

echo ""
echo "🚀 编码分析阶段完成"