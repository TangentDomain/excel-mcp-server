#!/usr/bin/env python3
"""
测试双行表头获取功能
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook
import json
from src.api.excel_operations import ExcelOperations

def create_test_excel():
    """创建测试用的Excel文件"""
    wb = Workbook()

    # 创建技能配置表
    ws1 = wb.active
    ws1.title = "技能配置表"

    # 第1行：字段描述
    ws1['A1'] = "技能ID描述"
    ws1['B1'] = "技能名称描述"
    ws1['C1'] = "技能类型描述"
    ws1['D1'] = "技能等级描述"
    ws1['E1'] = "技能消耗描述"

    # 第2行：字段名
    ws1['A2'] = "skill_id"
    ws1['B2'] = "skill_name"
    ws1['C2'] = "skill_type"
    ws1['D2'] = "skill_level"
    ws1['E2'] = "skill_cost"

    # 第3行开始：实际数据
    ws1['A3'] = 10001
    ws1['B3'] = "火球术"
    ws1['C3'] = "攻击"
    ws1['D3'] = 1
    ws1['E3'] = 20

    ws1['A4'] = 10002
    ws1['B4'] = "治疗术"
    ws1['C4'] = "治疗"
    ws1['D4'] = 1
    ws1['E4'] = 15

    # 创建装备配置表
    ws2 = wb.create_sheet("装备配置表")

    # 第1行：字段描述
    ws2['A1'] = "装备ID描述"
    ws2['B1'] = "装备名称描述"
    ws2['C1'] = "装备品质描述"
    ws2['D1'] = "装备类型描述"

    # 第2行：字段名
    ws2['A2'] = "item_id"
    ws2['B2'] = "item_name"
    ws2['C2'] = "item_quality"
    ws2['D2'] = "item_type"

    # 第3行开始：实际数据
    ws2['A3'] = 20001
    ws2['B3'] = "炎之剑"
    ws2['C3'] = "史诗"
    ws2['D3'] = "武器"

    # 保存文件
    test_file = "test_dual_headers.xlsx"
    wb.save(test_file)
    print(f"✅ 创建测试文件: {test_file}")
    return test_file

def test_get_headers(file_path):
    """测试单个工作表的双行表头获取"""
    print("\n🔍 测试 excel_get_headers 功能:")

    # 测试技能配置表
    result = ExcelOperations.get_headers(file_path, "技能配置表")
    print(f"📋 技能配置表结果:")
    print(f"  success: {result.get('success')}")
    print(f"  descriptions: {result.get('descriptions', [])}")
    print(f"  field_names: {result.get('field_names', [])}")
    print(f"  headers (兼容): {result.get('headers', [])}")
    print(f"  header_count: {result.get('header_count', 0)}")
    print(f"  message: {result.get('message', '')}")

    # 测试装备配置表
    result2 = ExcelOperations.get_headers(file_path, "装备配置表")
    print(f"\n📦 装备配置表结果:")
    print(f"  success: {result2.get('success')}")
    print(f"  descriptions: {result2.get('descriptions', [])}")
    print(f"  field_names: {result2.get('field_names', [])}")
    print(f"  headers (兼容): {result2.get('headers', [])}")
    print(f"  header_count: {result2.get('header_count', 0)}")

def test_get_sheet_headers(file_path):
    """测试所有工作表的双行表头获取"""
    print("\n🔍 测试 excel_get_sheet_headers 功能:")

    result = ExcelOperations.get_sheet_headers(file_path)
    print(f"📊 所有工作表结果:")
    print(f"  success: {result.get('success')}")
    print(f"  total_sheets: {result.get('total_sheets', 0)}")

    sheets = result.get('sheets_with_headers', [])
    for i, sheet in enumerate(sheets, 1):
        print(f"\n  📋 工作表 {i}: {sheet.get('name')}")
        print(f"    descriptions: {sheet.get('descriptions', [])}")
        print(f"    field_names: {sheet.get('field_names', [])}")
        print(f"    headers (兼容): {sheet.get('headers', [])}")
        print(f"    header_count: {sheet.get('header_count', 0)}")

        if 'error' in sheet:
            print(f"    ❌ error: {sheet['error']}")

def test_max_columns(file_path):
    """测试max_columns参数"""
    print("\n🔍 测试 max_columns 参数:")

    # 只获取前3列
    result = ExcelOperations.get_headers(file_path, "技能配置表", max_columns=3)
    print(f"📋 限制前3列结果:")
    print(f"  descriptions: {result.get('descriptions', [])}")
    print(f"  field_names: {result.get('field_names', [])}")
    print(f"  header_count: {result.get('header_count', 0)}")

def main():
    """主测试函数"""
    print("🚀 开始测试双行表头获取功能")

    try:
        # 创建测试文件
        test_file = create_test_excel()

        # 运行测试
        test_get_headers(test_file)
        test_get_sheet_headers(test_file)
        test_max_columns(test_file)

        print("\n✅ 所有测试完成!")

    except Exception as e:
        print(f"❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()

    finally:
        # 清理测试文件
        if os.path.exists("test_dual_headers.xlsx"):
            os.remove("test_dual_headers.xlsx")
            print("🧹 清理测试文件")

if __name__ == "__main__":
    main()
