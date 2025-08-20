#!/usr/bin/env python3
"""
Excel MCP Server - 测试运行器

全面运行所有单元测试，生成详细的测试报告和覆盖率统计
"""

import subprocess
import sys
import os
from pathlib import Path
import time

def run_tests():
    """运行所有测试并生成报告"""

    print("🧪 Excel MCP Server - 全面单元测试")
    print("=" * 60)

    # 检查测试环境
    test_files = [
        'test_validators.py',      # 验证器测试（现有）
        'test_parsers.py',         # 解析器测试（现有）
        'test_excel_reader.py',    # ExcelReader测试（新增）
        'test_excel_writer.py',    # ExcelWriter测试（新增）
        'test_mcp_tools.py',       # MCP工具测试（新增）
        'test_edge_cases.py',      # 边界测试（新增）
    ]

    tests_dir = Path('tests')
    missing_files = []

    print("📋 检查测试文件...")
    for test_file in test_files:
        test_path = tests_dir / test_file
        if test_path.exists():
            print(f"  ✅ {test_file}")
        else:
            print(f"  ❌ {test_file} (缺失)")
            missing_files.append(test_file)

    if missing_files:
        print(f"\n⚠️  警告: 缺失 {len(missing_files)} 个测试文件")
        print("     测试将继续运行现有文件")

    print(f"\n📊 测试统计:")
    print(f"  - 总测试文件: {len(test_files)}")
    print(f"  - 可用文件: {len(test_files) - len(missing_files)}")
    print(f"  - 缺失文件: {len(missing_files)}")

    # 运行测试
    print("\n🚀 开始运行测试...")
    print("-" * 60)

    start_time = time.time()

    # pytest命令选项
    pytest_args = [
        'python', '-m', 'pytest',
        'tests/',
        '-v',                    # 详细输出
        '--tb=short',           # 简短错误回溯
        '--durations=10',       # 显示最慢的10个测试
        '--strict-markers',     # 严格标记模式
        '--disable-warnings',   # 禁用警告显示
    ]

    # 尝试添加覆盖率报告
    try:
        subprocess.run(['python', '-m', 'pytest_cov', '--version'],
                      capture_output=True, check=True)
        pytest_args.extend([
            '--cov=server',
            '--cov=excel_mcp',
            '--cov-report=term-missing',
            '--cov-report=html:htmlcov',
        ])
        print("📈 启用覆盖率分析")
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("⚠️  pytest-cov 未安装，跳过覆盖率分析")

    try:
        # 运行pytest
        result = subprocess.run(pytest_args, capture_output=False)

        end_time = time.time()
        execution_time = end_time - start_time

        print("\n" + "=" * 60)
        print(f"⏱️  测试执行时间: {execution_time:.2f} 秒")

        if result.returncode == 0:
            print("🎉 所有测试通过!")
            return True
        else:
            print("❌ 部分测试失败")
            return False

    except FileNotFoundError:
        print("❌ pytest 未安装")
        print("请运行: pip install pytest pytest-cov")
        return False
    except Exception as e:
        print(f"❌ 测试运行出错: {e}")
        return False

def run_specific_test(test_name):
    """运行指定的测试文件"""
    print(f"🧪 运行特定测试: {test_name}")

    pytest_args = [
        'python', '-m', 'pytest',
        f'tests/{test_name}',
        '-v',
        '--tb=short',
    ]

    try:
        result = subprocess.run(pytest_args)
        return result.returncode == 0
    except FileNotFoundError:
        print("❌ pytest 未安装")
        return False

def show_test_info():
    """显示测试套件信息"""
    print("📚 Excel MCP Server 测试套件信息")
    print("=" * 60)

    test_categories = {
        "核心模块测试": [
            "test_excel_reader.py - ExcelReader模块全面测试",
            "test_excel_writer.py - ExcelWriter模块全面测试",
            "test_validators.py - 数据验证器测试",
            "test_parsers.py - 数据解析器测试",
        ],
        "MCP工具测试": [
            "test_mcp_tools.py - 所有15个MCP工具功能测试",
            "  ├─ excel_list_sheets",
            "  ├─ excel_regex_search",
            "  ├─ excel_get_range",
            "  ├─ excel_update_range",
            "  ├─ excel_insert_rows/columns",
            "  ├─ excel_delete_rows/columns",
            "  ├─ excel_create_file/sheet",
            "  ├─ excel_delete_sheet",
            "  ├─ excel_rename_sheet",
            "  ├─ excel_set_formula",
            "  └─ excel_format_cells",
        ],
        "边界和压力测试": [
            "test_edge_cases.py - 边界条件和错误处理",
            "  ├─ 边界值测试（最大行列、字符串长度等）",
            "  ├─ 错误处理测试（文件未找到、权限错误等）",
            "  ├─ 内存和性能测试（大文件、大数据集）",
            "  ├─ 恢复和稳定性测试（错误恢复、并发处理）",
            "  └─ 特殊字符测试（Unicode、控制字符等）",
        ]
    }

    for category, tests in test_categories.items():
        print(f"\n🔍 {category}:")
        for test in tests:
            print(f"    {test}")

    print(f"\n📊 测试覆盖范围:")
    print(f"    ✅ 正常功能场景")
    print(f"    ✅ 边界条件测试")
    print(f"    ✅ 异常错误处理")
    print(f"    ✅ 性能和内存测试")
    print(f"    ✅ 并发和稳定性")
    print(f"    ✅ Unicode和特殊字符")
    print(f"    ✅ 恢复能力测试")

def main():
    """主函数"""
    if len(sys.argv) > 1:
        command = sys.argv[1]

        if command == 'info':
            show_test_info()
        elif command == 'all':
            success = run_tests()
            sys.exit(0 if success else 1)
        elif command.startswith('test_'):
            success = run_specific_test(command)
            sys.exit(0 if success else 1)
        else:
            print("用法:")
            print("  python run_tests.py all       # 运行所有测试")
            print("  python run_tests.py info      # 显示测试信息")
            print("  python run_tests.py test_xxx.py  # 运行特定测试")
    else:
        # 默认运行所有测试
        success = run_tests()
        sys.exit(0 if success else 1)

if __name__ == '__main__':
    main()
