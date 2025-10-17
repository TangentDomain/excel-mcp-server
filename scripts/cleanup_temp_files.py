#!/usr/bin/env python3
"""
Excel MCP Server - 临时文件清理脚本

将测试过程中生成的临时文件移动到系统temp目录中统一管理
"""

import os
import tempfile
import shutil
import time
from pathlib import Path


def cleanup_test_temp_files():
    """清理测试过程中生成的临时文件"""

    # 获取系统temp目录
    system_temp = tempfile.gettempdir()
    excel_temp_dir = os.path.join(system_temp, "excel_mcp_server_tests")

    # 创建专用目录
    os.makedirs(excel_temp_dir, exist_ok=True)

    current_dir = Path(__file__).parent.parent

    # 需要清理的文件模式
    temp_file_patterns = [
        "*.xlsx",  # Excel临时文件
        "*.xlsm",  # Excel宏文件
        "*temp*",  # 包含temp的文件
        "test_*",  # 测试文件
        "*test*.xlsx",  # 测试Excel文件
        "large_test*",  # 大型测试文件
        "cell_test*",  # 单元格测试文件
        "performance_test*",  # 性能测试文件
        "edge_case_test*",  # 边界情况测试文件
        "structure_test*",  # 结构测试文件
        "structured_test*",  # 结构化测试文件
        "memory_test*",  # 内存测试文件
        "temp*",  # temp开头的文件
    ]

    moved_files = []
    cleaned_size = 0

    print(f"清理临时文件到: {excel_temp_dir}")

    # 遍历当前目录下的文件
    for pattern in temp_file_patterns:
        for file_path in current_dir.glob(pattern):
            if file_path.is_file():
                try:
                    # 跳过重要的项目文件
                    if file_path.name in [
                        'pyproject.toml', 'uv.lock', 'README.md',
                        'LICENSE', 'CLAUDE.md', 'CONTRIBUTING.md',
                        'deploy.bat', 'mcp-windows.json', 'mcp-direct.json',
                        'mcp-generated.json', '项目说明.md'
                    ]:
                        continue

                    # 生成新的文件名（添加时间戳避免冲突）
                    timestamp = int(time.time())
                    new_name = f"{timestamp}_{file_path.name}"
                    new_path = os.path.join(excel_temp_dir, new_name)

                    # 移动文件
                    shutil.move(str(file_path), new_path)

                    # 记录移动的文件
                    file_size = os.path.getsize(new_path)
                    moved_files.append(new_name)
                    cleaned_size += file_size

                    print(f"  移动: {file_path.name} -> {new_name} ({file_size:,} bytes)")

                except (PermissionError, OSError) as e:
                    print(f"  跳过文件 {file_path.name}: {e}")
                except Exception as e:
                    print(f"  错误处理文件 {file_path.name}: {e}")

    # 清理空的JSON文件（通常是测试元数据）
    for json_file in current_dir.glob("*.json"):
        if json_file.is_file():
            try:
                # 检查文件大小，很小的可能是临时元数据
                if json_file.stat().st_size < 1024:  # 小于1KB
                    timestamp = int(time.time())
                    new_name = f"{timestamp}_{json_file.name}"
                    new_path = os.path.join(excel_temp_dir, new_name)

                    shutil.move(str(json_file), new_path)
                    moved_files.append(new_name)
                    print(f"  移动JSON: {json_file.name}")

            except Exception as e:
                print(f"  跳过JSON文件 {json_file.name}: {e}")

    # 创建清理报告
    report_file = os.path.join(excel_temp_dir, "cleanup_report.txt")
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write(f"Excel MCP Server 临时文件清理报告\n")
        f.write(f"清理时间: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"目标目录: {excel_temp_dir}\n")
        f.write(f"移动文件数量: {len(moved_files)}\n")
        f.write(f"总大小: {cleaned_size:,} bytes\n")
        f.write(f"\n移动的文件列表:\n")
        for file_name in moved_files:
            f.write(f"  - {file_name}\n")

    print(f"\n清理完成!")
    print(f"  移动了 {len(moved_files)} 个文件")
    print(f"  释放了 {cleaned_size:,} 字节空间")
    print(f"  清理报告: {report_file}")

    return {
        'moved_files': len(moved_files),
        'cleaned_size': cleaned_size,
        'target_dir': excel_temp_dir
    }


def setup_temp_file_management():
    """设置临时文件管理"""

    # 确保系统temp目录存在专用子目录
    system_temp = tempfile.gettempdir()
    excel_temp_dir = os.path.join(system_temp, "excel_mcp_server_tests")

    os.makedirs(excel_temp_dir, exist_ok=True)

    # 创建.gitignore以避免意外提交
    gitignore_path = os.path.join(excel_temp_dir, ".gitignore")
    if not os.path.exists(gitignore_path):
        with open(gitignore_path, 'w', encoding='utf-8') as f:
            f.write("# Excel MCP Server 临时文件目录\n")
            f.write("# 此目录下的文件都是测试过程中生成的临时文件\n")
            f.write("# 可以安全删除\n")
            f.write("*\n")
            f.write("!.gitignore\n")

    print(f"临时文件管理目录: {excel_temp_dir}")
    return excel_temp_dir


if __name__ == "__main__":
    print("Excel MCP Server - 临时文件清理工具")
    print("=" * 50)

    # 设置临时文件管理
    setup_temp_file_management()

    # 清理临时文件
    result = cleanup_test_temp_files()

    print("\n清理完成。临时文件已移动到系统temp目录中统一管理。")