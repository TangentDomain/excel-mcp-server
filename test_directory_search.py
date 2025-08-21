#!/usr/bin/env python3
"""
测试目录搜索功能
"""

import os
import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# 导入模块
try:
    from src.core.excel_search import ExcelSearcher
    print("✅ 成功导入ExcelSearcher")
except ImportError as e:
    print(f"❌ 导入ExcelSearcher失败: {e}")
    sys.exit(1)

def test_directory_search():
    """Test directory search functionality"""

    # 检查是否有数据目录
    data_dir = project_root / "data"
    if not data_dir.exists():
        print(f"❌ 数据目录不存在: {data_dir}")
        return

    print(f"✅ 找到数据目录: {data_dir}")

    # 查找 Excel 文件
    excel_files = list(data_dir.rglob("*.xlsx")) + list(data_dir.rglob("*.xlsm"))
    print(f"✅ 找到 {len(excel_files)} 个 Excel 文件")

    if not excel_files:
        print("ℹ️ 没有找到 Excel 文件，无法测试")
        return

    # 创建搜索器实例(使用第一个文件作为初始化)
    searcher = ExcelSearcher(str(excel_files[0]))

    # 测试目录搜索
    print("\n正在测试目录搜索功能...")

    try:
        # 简单的数字搜索测试
        result = searcher.regex_search_directory(
            directory_path=str(data_dir),
            pattern=r'\d+',  # 搜索数字
            flags="",
            search_values=True,
            search_formulas=False,
            recursive=True,
            max_files=10
        )

        if result.success:
            print(f"✅ 目录搜索成功!")
            print(f"   - 总匹配数: {result.metadata['total_matches']}")
            print(f"   - 找到文件数: {result.metadata['total_files_found']}")
            print(f"   - 搜索成功文件: {len(result.metadata['searched_files'])}")
            print(f"   - 跳过文件: {len(result.metadata['skipped_files'])}")

            if result.data and len(result.data) > 0:
                print(f"   - 示例匹配: {result.data[0]}")
        else:
            print(f"❌ 目录搜索失败: {result.error}")

    except Exception as e:
        print(f"❌ 测试发生异常: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    print("🚀 开始测试目录搜索功能")
    test_directory_search()
    print("✨ 测试完成")
