#!/usr/bin/env python3
"""
测试新增的sheet_name参数功能
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from core.excel_search import ExcelSearcher
import tempfile
from openpyxl import Workbook

def test_sheet_name_parameter():
    """测试新增的sheet_name参数"""
    print("🔍 测试excel_regex_search的sheet_name参数功能")

    # 创建测试文件
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'Sheet1'
    ws1['A1'] = 'test123'
    ws1['B1'] = 'hello'
    ws1['A2'] = 'data999'

    ws2 = wb.create_sheet('Sheet2')
    ws2['A1'] = 'test456'
    ws2['B1'] = 'world'
    ws2['A2'] = 'info888'

    # 保存临时文件
    temp_file = tempfile.mktemp(suffix='.xlsx')
    wb.save(temp_file)

    try:
        searcher = ExcelSearcher(temp_file)

        print("\n1. 测试搜索所有工作表:")
        result_all = searcher.regex_search(r'test\d+')
        if result_all.success:
            print(f"   ✓ 找到 {len(result_all.data)} 个匹配")
            for match in result_all.data:
                print(f"     - {match.sheet}!{match.cell}: '{match.match}'")
        else:
            print(f"   ✗ 错误: {result_all.error}")

        print("\n2. 测试只搜索Sheet1:")
        result_sheet1 = searcher.regex_search(r'test\d+', sheet_name='Sheet1')
        if result_sheet1.success:
            print(f"   ✓ 找到 {len(result_sheet1.data)} 个匹配")
            for match in result_sheet1.data:
                print(f"     - {match.sheet}!{match.cell}: '{match.match}'")
        else:
            print(f"   ✗ 错误: {result_sheet1.error}")

        print("\n3. 测试只搜索Sheet2:")
        result_sheet2 = searcher.regex_search(r'test\d+', sheet_name='Sheet2')
        if result_sheet2.success:
            print(f"   ✓ 找到 {len(result_sheet2.data)} 个匹配")
            for match in result_sheet2.data:
                print(f"     - {match.sheet}!{match.cell}: '{match.match}'")
        else:
            print(f"   ✗ 错误: {result_sheet2.error}")

        print("\n4. 测试搜索不存在的工作表:")
        result_invalid = searcher.regex_search(r'test\d+', sheet_name='NonExistent')
        if not result_invalid.success:
            print(f"   ✓ 正确处理了不存在的工作表: {result_invalid.error}")
        else:
            print("   ✗ 应该返回错误，但没有")

        print("\n5. 测试搜索数字模式:")
        result_numbers = searcher.regex_search(r'\d{3}', sheet_name='Sheet1')
        if result_numbers.success:
            print(f"   ✓ 在Sheet1中找到 {len(result_numbers.data)} 个三位数")
            for match in result_numbers.data:
                print(f"     - {match.sheet}!{match.cell}: '{match.match}'")
        else:
            print(f"   ✗ 错误: {result_numbers.error}")

        print("\n🎉 sheet_name参数功能测试完成！")
        return True

    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False
    finally:
        # 清理临时文件
        if os.path.exists(temp_file):
            os.remove(temp_file)

if __name__ == "__main__":
    success = test_sheet_name_parameter()
    sys.exit(0 if success else 1)
