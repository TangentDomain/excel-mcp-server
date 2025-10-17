#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel MCP Server - 快速监控脚本

简化版的监控脚本，专注于核心功能和易用性。
"""

import subprocess
import sys
import time
import json
import os
from datetime import datetime
from pathlib import Path

# 设置输出编码
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def run_test_monitoring():
    """运行快速测试监控"""
    print("="*60)
    print("Excel MCP Server 快速监控")
    print(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)

    base_dir = Path(__file__).parent.parent
    results = {}

    try:
        # 1. 运行测试
        print("[INFO] 运行测试套件...")
        start_time = time.time()

        test_cmd = [
            sys.executable, "-m", "pytest",
            "tests/",
            "--cov=src",
            "--cov-report=json",
            "--cov-report=term-missing",
            "--tb=short",
            "-v"
        ]

        result = subprocess.run(test_cmd, cwd=base_dir, capture_output=True, text=True, timeout=600)

        execution_time = time.time() - start_time

        # 2. 解析结果
        print(f"[RESULT] 执行时间: {execution_time:.1f}秒")

        # 解析测试输出
        output_lines = result.stdout.split('\n')
        test_count = 0
        passed_count = 0
        failed_count = 0

        for line in output_lines:
            if "passed" in line and ("failed" in line or "error" in line):
                import re
                # 提取测试数量
                passed_match = re.search(r'(\d+)\s+passed', line)
                failed_match = re.search(r'(\d+)\s+failed', line)
                error_match = re.search(r'(\d+)\s+error', line)

                if passed_match:
                    passed_count = int(passed_match.group(1))
                if failed_match:
                    failed_count = int(failed_match.group(1))
                if error_match:
                    failed_count += int(error_match.group(1))

                test_count = passed_count + failed_count

        # 3. 读取覆盖率数据
        coverage_file = base_dir / "coverage.json"
        total_coverage = 0
        file_count = 0

        if coverage_file.exists():
            with open(coverage_file, 'r', encoding='utf-8') as f:
                coverage_data = json.load(f)

            total_coverage = coverage_data['totals']['percent_covered']
            file_count = len(coverage_data['files'])

            print(f"[RESULT] 测试覆盖率: {total_coverage:.1f}%")
            print(f"[RESULT] 覆盖文件数: {file_count}")

        # 4. 显示测试结果
        if test_count > 0:
            success_rate = (passed_count / test_count) * 100
            print(f"[RESULT] 通过测试: {passed_count}/{test_count} ({success_rate:.1f}%)")

            if failed_count > 0:
                print(f"[WARNING] 失败测试: {failed_count}")
        else:
            print("[WARNING] 无法解析测试结果")

        # 5. 性能评估
        if execution_time < 120:
            performance_status = "优秀"
        elif execution_time < 300:
            performance_status = "良好"
        elif execution_time < 600:
            performance_status = "一般"
        else:
            performance_status = "需要优化"

        print(f"[RESULT] 性能状态: {performance_status}")

        # 6. 生成建议
        print("\n[RECOMMENDATIONS] 建议:")
        recommendations = []

        if total_coverage < 85:
            recommendations.append(f"提高测试覆盖率 (当前: {total_coverage:.1f}%, 目标: 85%)")

        if failed_count > 0:
            recommendations.append(f"修复 {failed_count} 个失败的测试")

        if execution_time > 300:
            recommendations.append("优化测试执行速度")

        if not recommendations:
            recommendations.append("所有指标良好，继续保持！")

        for i, rec in enumerate(recommendations, 1):
            print(f"  {i}. {rec}")

        # 7. 保存结果
        results = {
            'timestamp': datetime.now().isoformat(),
            'test_count': test_count,
            'passed_count': passed_count,
            'failed_count': failed_count,
            'success_rate': success_rate if test_count > 0 else 0,
            'coverage': total_coverage,
            'file_count': file_count,
            'execution_time': execution_time,
            'performance_status': performance_status,
            'recommendations': recommendations
        }

        # 保存到文件
        results_file = base_dir / "reports" / f"quick-monitor-{datetime.now().strftime('%Y%m%d-%H%M%S')}.json"
        results_file.parent.mkdir(exist_ok=True)

        with open(results_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        print(f"\n[INFO] 结果已保存: {results_file}")

    except subprocess.TimeoutExpired:
        print("[ERROR] 测试超时 (超过10分钟)")
    except Exception as e:
        print(f"[ERROR] 监控失败: {e}")

    print("="*60)
    print("[COMPLETE] 监控完成")

    return results

if __name__ == "__main__":
    run_test_monitoring()