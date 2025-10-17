#!/usr/bin/env python3
"""
Excel MCP Server - 测试文件清理脚本
清理测试过程中生成的Excel文件
"""

import os
import shutil
import tempfile
import time
from pathlib import Path

def cleanup_test_excel_files():
    """清理测试Excel文件"""

    # 获取系统temp目录
    system_temp = tempfile.gettempdir()
    excel_test_dir = os.path.join(system_temp, "excel_mcp_server_test_files")

    # 创建专用目录
    os.makedirs(excel_test_dir, exist_ok=True)

    current_dir = Path('.')

    # 需要清理的Excel文件模式
    test_excel_patterns = [
        '*test*.xlsx',
        '*test*.xlsm',
        'cell_test*',
        'large_test*',
        'edge_case_test*',
        'structure_test*',
        'structured_test*',
        'memory_test*',
        'performance_test*',
        'temp*',
        'temp.xlsx',
        'temp_test.xlsx'
    ]

    moved_count = 0
    total_size = 0

    print(f"Moving test Excel files to: {excel_test_dir}")

    for pattern in test_excel_patterns:
        for file_path in current_dir.glob(pattern):
            if file_path.is_file() and file_path.suffix.lower() in ['.xlsx', '.xlsm']:
                try:
                    # 生成带时间戳的文件名避免冲突
                    timestamp = int(time.time())
                    new_name = f"{timestamp}_{file_path.name}"
                    new_path = os.path.join(excel_test_dir, new_name)

                    # 获取文件大小
                    file_size = file_path.stat().st_size

                    # 移动文件
                    shutil.move(str(file_path), new_path)

                    moved_count += 1
                    total_size += file_size

                    print(f"  Moved: {file_path.name} -> {new_name} ({file_size:,} bytes)")

                except Exception as e:
                    print(f"  ERROR moving {file_path.name}: {e}")

    # 创建清理报告
    report_file = os.path.join(excel_test_dir, "test_files_cleanup_report.txt")
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write(f"Excel MCP Server Test Files Cleanup Report\n")
        f.write(f"Cleanup time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Target directory: {excel_test_dir}\n")
        f.write(f"Moved files count: {moved_count}\n")
        f.write(f"Total size: {total_size:,} bytes\n")

    print(f"\nCleanup complete!")
    print(f"  Moved {moved_count} Excel files")
    print(f"  Total size: {total_size:,} bytes")
    print(f"  Report: {report_file}")

    return moved_count, total_size

def cleanup_temp_python_files():
    """清理临时Python文件"""

    current_dir = Path('.')
    temp_dir = current_dir / 'temp'

    # 需要清理的Python文件模式
    temp_python_patterns = [
        '*temp*.py',
        'test_*template*.py',
        'comprehensive_verification.py',
        'test_*enhanced*.py'
    ]

    cleaned_count = 0

    for pattern in temp_python_patterns:
        for file_path in current_dir.glob(pattern):
            if file_path.is_file():
                try:
                    dst = temp_dir / file_path.name
                    temp_dir.mkdir(exist_ok=True)
                    shutil.move(str(file_path), str(dst))
                    print(f"  Moved Python temp: {file_path.name} -> temp/")
                    cleaned_count += 1
                except Exception as e:
                    print(f"  ERROR moving {file_path.name}: {e}")

    return cleaned_count

def main():
    """主函数"""
    print("Excel MCP Server - Test Files Cleanup")
    print("=" * 50)

    # 清理测试Excel文件
    print("\n1. Cleaning test Excel files...")
    moved_excel, excel_size = cleanup_test_excel_files()

    # 清理临时Python文件
    print("\n2. Cleaning temporary Python files...")
    moved_python = cleanup_temp_python_files()

    # 总结
    print("\n" + "=" * 50)
    print(f"CLEANUP COMPLETE!")
    print(f"Excel files moved: {moved_excel} ({excel_size:,} bytes)")
    print(f"Python files moved: {moved_python}")
    print(f"Total files cleaned: {moved_excel + moved_python}")

if __name__ == "__main__":
    main()