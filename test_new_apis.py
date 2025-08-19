#!/usr/bin/env python3
"""
Excel MCP 新API功能测试脚本
测试新增的6个API功能的正确性
"""

import os
import sys
import tempfile
from pathlib import Path

# 添加当前目录到Python路径以导入server模块
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from server import (
    excel_create_file, excel_create_sheet, excel_delete_sheet,
    excel_rename_sheet, excel_delete_rows, excel_delete_columns,
    excel_list_sheets, excel_get_range, excel_update_range
)

def test_new_apis():
    """测试新增的6个API功能"""
    print("🧪 开始测试Excel MCP新API功能...")

    # 创建临时测试文件路径
    temp_dir = Path(tempfile.mkdtemp())
    test_file = temp_dir / "test_new_apis.xlsx"

    try:
        # 测试1: excel_create_file - 创建新文件
        print("\n📁 测试1: excel_create_file - 创建新Excel文件")
        result = excel_create_file(
            file_path=str(test_file),
            sheet_names=["主数据", "备份数据", "统计数据"]
        )

        if result['success']:
            print(f"  ✅ 成功创建文件: {result['message']}")
            print(f"     📊 工作表数量: {result['total_sheets']}")
            for sheet in result['sheets']:
                marker = "🎯" if sheet['is_active'] else "📄"
                print(f"     {marker} {sheet['index']+1}. {sheet['name']}")
        else:
            print(f"  ❌ 创建失败: {result['error']}")
            return False

        # 测试2: excel_create_sheet - 创建新工作表
        print("\n📋 测试2: excel_create_sheet - 添加新工作表")
        result = excel_create_sheet(
            file_path=str(test_file),
            sheet_name="临时工作表",
            index=1
        )

        if result['success']:
            print(f"  ✅ 成功创建工作表: {result['message']}")
            print(f"     📍 位置索引: {result['sheet_info']['index']}")
            print(f"     📚 总工作表数: {result['total_sheets']}")
        else:
            print(f"  ❌ 创建失败: {result['error']}")

        # 测试3: excel_rename_sheet - 重命名工作表
        print("\n✏️ 测试3: excel_rename_sheet - 重命名工作表")
        result = excel_rename_sheet(
            file_path=str(test_file),
            old_name="临时工作表",
            new_name="重命名工作表"
        )

        if result['success']:
            print(f"  ✅ 成功重命名: {result['message']}")
            print(f"     📝 新名称: {result['new_name']}")
        else:
            print(f"  ❌ 重命名失败: {result['error']}")

        # 添加一些测试数据
        print("\n📊 添加测试数据...")
        excel_update_range(
            file_path=str(test_file),
            range_expression="主数据!A1:C5",
            data=[
                ["姓名", "年龄", "城市"],
                ["张三", 25, "北京"],
                ["李四", 30, "上海"],
                ["王五", 28, "广州"],
                ["赵六", 32, "深圳"]
            ]
        )

        # 测试4: excel_delete_rows - 删除行
        print("\n🗑️ 测试4: excel_delete_rows - 删除行")
        result = excel_delete_rows(
            file_path=str(test_file),
            sheet_name="主数据",
            start_row=3,
            count=2
        )

        if result['success']:
            print(f"  ✅ 成功删除行: {result['message']}")
            print(f"     📊 删除数量: {result['actual_deleted_count']}")
            print(f"     📈 原行数: {result['original_max_row']} → 新行数: {result['new_max_row']}")
        else:
            print(f"  ❌ 删除失败: {result['error']}")

        # 测试5: excel_delete_columns - 删除列
        print("\n🗑️ 测试5: excel_delete_columns - 删除列")
        result = excel_delete_columns(
            file_path=str(test_file),
            sheet_name="主数据",
            start_column=3,
            count=1
        )

        if result['success']:
            print(f"  ✅ 成功删除列: {result['message']}")
            print(f"     📊 删除数量: {result['actual_deleted_count']}")
            print(f"     📈 原列数: {result['original_max_column']} → 新列数: {result['new_max_column']}")
        else:
            print(f"  ❌ 删除失败: {result['error']}")

        # 测试6: excel_delete_sheet - 删除工作表
        print("\n🗑️ 测试6: excel_delete_sheet - 删除工作表")
        result = excel_delete_sheet(
            file_path=str(test_file),
            sheet_name="重命名工作表"
        )

        if result['success']:
            print(f"  ✅ 成功删除工作表: {result['message']}")
            print(f"     🎯 新活动工作表: {result['new_active_sheet']}")
            print(f"     📚 剩余工作表: {result['remaining_sheets']}")
        else:
            print(f"  ❌ 删除失败: {result['error']}")

        # 验证最终状态
        print("\n🔍 验证最终文件状态...")
        result = excel_list_sheets(file_path=str(test_file))
        if result['success']:
            print(f"  📊 最终工作表数量: {result['total_sheets']}")
            print(f"  🎯 活动工作表: {result['active_sheet']}")
            for sheet in result['sheets']:
                marker = "🎯" if sheet['is_active'] else "📄"
                print(f"     {marker} {sheet['index']+1}. {sheet['name']} (数据范围: {sheet['max_column_letter']}{sheet['max_row']})")

        print("\n🎉 所有新API测试完成！")
        return True

    except Exception as e:
        print(f"\n❌ 测试过程中发生错误: {e}")
        return False

    finally:
        # 清理临时文件
        if test_file.exists():
            test_file.unlink()
            print(f"\n🧹 已清理临时文件: {test_file}")

def main():
    """主测试函数"""
    print("=" * 60)
    print("Excel MCP 新API功能测试")
    print("=" * 60)

    success = test_new_apis()

    if success:
        print("\n✅ 所有测试通过！新API功能正常。")
        return 0
    else:
        print("\n❌ 测试失败，请检查实现。")
        return 1

if __name__ == "__main__":
    exit(main())
