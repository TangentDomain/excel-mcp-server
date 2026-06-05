#!/bin/bash
#
# 启动边缘测试轮次脚本
# 功能:
# 1. 调用边缘案例发现脚本获取最近发现的高优先级边缘案例
# 2. 生成当前轮次测试脚本
# 3. 记录测试开始时间
#

set -e  # 遇到错误立即退出

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
LOG_DIR="${PROJECT_ROOT}/.test-logs"
TIMESTAMP_FILE="${LOG_DIR}/round-start-time.txt"

# 创建日志目录
mkdir -p "${LOG_DIR}"

# 函数: 调用边缘案例发现脚本
discover_edge_cases() {
    """
    调用边缘案例发现脚本获取高优先级边缘案例

    Returns:
        0: 成功
        1: 失败
    """
    echo "=================================================="
    echo "步骤1: 获取最近发现的高优先级边缘案例"
    echo "=================================================="

    cd "${PROJECT_ROOT}"

    if ! python3 "${SCRIPT_DIR}/edge_case_discovery.py" --recent --limit 10; then
        echo "错误: 边缘案例发现失败"
        return 1
    fi

    echo "✓ 边缘案例发现完成"
    return 0
}

# 函数: 生成当前轮次测试脚本
generate_test_script() {
    """
    生成当前轮次测试脚本

    Returns:
        0: 成功
        1: 失败
    """
    echo "=================================================="
    echo "步骤2: 生成当前轮次测试脚本"
    echo "=================================================="

    cd "${PROJECT_ROOT}"

    if ! python3 "${SCRIPT_DIR}/edge_case_automation.py" --load --generate; then
        echo "错误: 测试脚本生成失败"
        return 1
    fi

    echo "✓ 测试脚本生成完成"
    return 0
}

# 函数: 记录测试开始时间
record_start_time() {
    """
    记录测试开始时间到文件

    Returns:
        0: 成功
        1: 失败
    """
    echo "=================================================="
    echo "步骤3: 记录测试开始时间"
    echo "=================================================="

    local start_time
    start_time=$(date -u +"%Y-%m-%d %H:%M:%S UTC")

    # 写入时间戳文件
    {
        echo "轮次开始时间: ${start_time}"
        echo "Unix时间戳: $(date +%s)"
        echo "工作目录: ${PROJECT_ROOT}"
        echo "Git分支: $(git branch --show-current 2>/dev/null || echo 'unknown')"
        echo "Git提交: $(git rev-parse --short HEAD 2>/dev/null || echo 'unknown')"
    } > "${TIMESTAMP_FILE}"

    echo "✓ 测试开始时间已记录到: ${TIMESTAMP_FILE}"
    echo "  时间: ${start_time}"

    return 0
}

# 主函数
main() {
    """
    主函数：按顺序执行所有步骤
    """
    echo "=================================================="
    echo "边缘测试轮次启动脚本"
    echo "=================================================="
    echo "项目根目录: ${PROJECT_ROOT}"
    echo "工作目录: $(pwd)"
    echo "=================================================="

    # 步骤1: 获取边缘案例
    if ! discover_edge_cases; then
        echo "❌ 边缘案例发现失败，停止执行"
        exit 1
    fi

    echo ""

    # 步骤2: 生成测试脚本
    if ! generate_test_script; then
        echo "❌ 测试脚本生成失败，停止执行"
        exit 1
    fi

    echo ""

    # 步骤3: 记录开始时间
    if ! record_start_time; then
        echo "❌ 记录开始时间失败，停止执行"
        exit 1
    fi

    echo ""
    echo "=================================================="
    echo "✓ 轮次启动成功"
    echo "=================================================="
    echo "下一步: 执行测试脚本和 end-round.sh"
    echo "=================================================="

    return 0
}

# 执行主函数
main "$@"
