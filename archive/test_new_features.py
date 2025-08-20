#!/usr/bin/env python3
"""
测试新功能的简单脚本
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from server import excel_list_sheets, excel_get_range
import openpyxl

def test_new_features():
    """测试新添加的功能"""
    print("🧪 测试Excel MCP新功能")
    print("=" * 40)

    # 创建简单测试文件
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = '数据表'
    ws1['A1'] = '测试数据'
    ws1['A2'] = '第二行'
    ws1['B1'] = '列B数据'

    ws2 = wb.create_sheet('计算表')
    ws2['A1'] = '计算结果'

    test_file = 'temp_test.xlsx'
    wb.save(test_file)

    try:
        # 测试1: 工作表列表功能
        print("\n📋 测试工作表列表功能:")
        result = excel_list_sheets(test_file)
        if result['success']:
            print(f"  ✅ 共 {result['total_sheets']} 个工作表")
            for sheet in result['sheets']:
                active = "🎯" if sheet['is_active'] else "📄"
                print(f"    {active} {sheet['name']} (数据范围: {sheet['max_column_letter']}{sheet['max_row']})")
        else:
            print(f"  ❌ 失败: {result['error']}")

        # 测试2: 行访问功能
        print("\n🔢 测试行访问功能:")
        result = excel_get_range(test_file, '1:1')
        if result['success']:
            print(f"  ✅ 第1行访问成功，类型: {result['range_type']}")
            print(f"    📊 维度: {result['dimensions']['rows']}x{result['dimensions']['columns']}")
        else:
            print(f"  ❌ 失败: {result['error']}")

        # 测试3: 列访问功能
        print("\n📊 测试列访问功能:")
        result = excel_get_range(test_file, 'A:A')
        if result['success']:
            print(f"  ✅ A列访问成功，类型: {result['range_type']}")
            print(f"    📊 维度: {result['dimensions']['rows']}x{result['dimensions']['columns']}")
        else:
            print(f"  ❌ 失败: {result['error']}")

        print("\n🎉 新功能测试完成!")

    except Exception as e:
        print(f"\n❌ 测试错误: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # 清理测试文件
        import os
        if os.path.exists(test_file):
            os.unlink(test_file)

if __name__ == "__main__":
    test_new_features()
