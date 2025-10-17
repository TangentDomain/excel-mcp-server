#!/usr/bin/env python3
"""
Excel MCP 服务器安全测试运行脚本

运行所有安全相关的测试，验证安全功能的正确性和有效性。
"""

import os
import sys
import subprocess
import time
import tempfile
import shutil
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src"))

def run_test_module(test_module):
    """运行单个测试模块"""
    print(f"\n{'='*60}")
    print(f"运行测试模块: {test_module}")
    print(f"{'='*60}")

    test_file = project_root / "tests" / test_module

    if not test_file.exists():
        print(f"❌ 测试文件不存在: {test_file}")
        return False

    try:
        # 使用pytest运行测试
        result = subprocess.run([
            sys.executable, "-m", "pytest",
            str(test_file),
            "-v",
            "--tb=short",
            "--maxfail=10"
        ],
        cwd=project_root,
        capture_output=True,
        text=True,
        timeout=300  # 5分钟超时
        )

        print(result.stdout)
        if result.stderr:
            print("错误输出:")
            print(result.stderr)

        if result.returncode == 0:
            print(f"✅ {test_module} 测试通过")
            return True
        else:
            print(f"❌ {test_module} 测试失败 (退出码: {result.returncode})")
            return False

    except subprocess.TimeoutExpired:
        print(f"❌ {test_module} 测试超时")
        return False
    except Exception as e:
        print(f"❌ {test_module} 测试执行出错: {str(e)}")
        return False

def run_security_tests():
    """运行所有安全测试"""
    print("🛡️ Excel MCP 服务器安全测试")
    print("=" * 60)

    # 确保在项目根目录
    os.chdir(project_root)

    # 创建临时目录用于测试
    temp_dir = tempfile.mkdtemp(prefix="excel_security_test_")
    print(f"📁 使用临时目录: {temp_dir}")

    try:
        # 设置环境变量，确保测试使用临时目录
        os.environ['EXCEL_TEST_TEMP_DIR'] = temp_dir

        security_tests = [
            "test_safety_features.py",
            "test_backup_recovery.py",
            "test_user_confirmation.py",
            "test_security_penetration.py"
        ]

        results = {}
        total_start_time = time.time()

        for test_module in security_tests:
            start_time = time.time()
            success = run_test_module(test_module)
            end_time = time.time()

            results[test_module] = {
                'success': success,
                'duration': end_time - start_time
            }

        total_duration = time.time() - total_start_time

        # 生成测试报告
        print("\n" + "="*60)
        print("🏁 安全测试总结")
        print("="*60)

        passed_count = sum(1 for r in results.values() if r['success'])
        total_count = len(results)

        print(f"总测试模块: {total_count}")
        print(f"通过模块: {passed_count}")
        print(f"失败模块: {total_count - passed_count}")
        print(f"总耗时: {total_duration:.2f}秒")

        print("\n详细结果:")
        for test_module, result in results.items():
            status = "✅ 通过" if result['success'] else "❌ 失败"
            duration = result['duration']
            print(f"  {test_module:<30} {status} ({duration:.2f}s)")

        if passed_count == total_count:
            print("\n🎉 所有安全测试都通过了！")
            print("✅ Excel MCP 服务器安全功能验证成功")
            return True
        else:
            print(f"\n⚠️  {total_count - passed_count} 个测试模块失败")
            print("❌ 需要修复失败的测试用例")
            return False

    finally:
        # 清理临时目录
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            print(f"\n🧹 已清理临时目录: {temp_dir}")

def run_coverage_analysis():
    """运行测试覆盖率分析"""
    print("\n" + "="*60)
    print("📊 生成测试覆盖率报告")
    print("="*60)

    try:
        # 运行覆盖率测试
        result = subprocess.run([
            sys.executable, "-m", "pytest",
            "tests/test_safety_features.py",
            "tests/test_backup_recovery.py",
            "tests/test_user_confirmation.py",
            "tests/test_security_penetration.py",
            "--cov=src/api",
            "--cov=src/utils",
            "--cov-report=html",
            "--cov-report=term",
            "--cov-report=xml",
            "-v"
        ],
        cwd=project_root,
        capture_output=True,
        text=True,
        timeout=600  # 10分钟超时
        )

        print(result.stdout)
        if result.stderr:
            print("覆盖率错误输出:")
            print(result.stderr)

        if result.returncode == 0:
            print("✅ 覆盖率报告生成成功")
            print("📄 HTML报告: htmlcov/index.html")
            return True
        else:
            print("❌ 覆盖率报告生成失败")
            return False

    except subprocess.TimeoutExpired:
        print("❌ 覆盖率分析超时")
        return False
    except Exception as e:
        print(f"❌ 覆盖率分析出错: {str(e)}")
        return False

def generate_security_report():
    """生成安全测试报告"""
    print("\n" + "="*60)
    print("📋 生成安全测试报告")
    print("="*60)

    report = {
        "test_date": time.strftime("%Y-%m-%d %H:%M:%S"),
        "python_version": sys.version,
        "platform": sys.platform,
        "project_root": str(project_root),
        "security_features": {
            "data_impact_assessment": "✅ 实现",
            "dangerous_operation_warnings": "✅ 实现",
            "file_status_checks": "✅ 实现",
            "operation_confirmation": "✅ 实现",
            "automatic_backup": "✅ 实现",
            "operation_cancellation": "✅ 实现",
            "safety_guidance": "✅ 实现",
            "security_documentation": "✅ 实现"
        },
        "test_categories": [
            "安全功能测试",
            "备份恢复测试",
            "用户确认测试",
            "渗透测试"
        ]
    }

    report_file = project_root / "security_test_report.json"

    try:
        import json
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

        print(f"✅ 安全测试报告已生成: {report_file}")
        return True
    except Exception as e:
        print(f"❌ 报告生成失败: {str(e)}")
        return False

def main():
    """主函数"""
    print("🚀 开始Excel MCP服务器安全测试流程")

    # 检查依赖
    try:
        import pytest
        import openpyxl
        print("✅ 测试依赖检查通过")
    except ImportError as e:
        print(f"❌ 缺少测试依赖: {e}")
        print("请运行: pip install pytest openpyxl")
        return False

    # 运行安全测试
    test_success = run_security_tests()

    if test_success:
        # 生成覆盖率报告
        coverage_success = run_coverage_analysis()

        # 生成安全报告
        report_success = generate_security_report()

        if coverage_success and report_success:
            print("\n🎯 所有安全测试流程完成！")
            print("🛡️ Excel MCP服务器已通过全面安全验证")
            return True
        else:
            print("\n⚠️ 部分后续流程失败，但核心安全测试通过")
            return True
    else:
        print("\n❌ 安全测试失败，请修复后重试")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)