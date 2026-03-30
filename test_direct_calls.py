#!/usr/bin/env python3
"""直接调用server.py中的方法来复现API问题"""

import tempfile
import os
import sys
from pathlib import Path

# 添加项目路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src"))

from openpyxl import Workbook

def create_test_file():
    """创建测试Excel文件"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        test_file = f.name
    
    # 创建工作簿并写入测试数据
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    
    # 写入测试数据
    test_data = [
        ['Name', 'Age', 'Score'],
        ['Alice', 25, 85],
        ['Bob', 30, 90],
        ['Charlie', 28, 78]
    ]
    
    for i, row in enumerate(test_data):
        for j, cell_value in enumerate(row):
            ws.cell(row=i+1, column=j+1, value=cell_value)
    
    wb.save(test_file)
    print(f'✅ 创建测试文件成功: {test_file}')
    return test_file

def test_direct_server_calls():
    """直接调用server.py中的方法来复现问题"""
    
    test_file = create_test_file()
    
    try:
        # 直接导入server模块中的函数
        from excel_mcp_server_fastmcp.server import (
            excel_get_range,
            excel_format_cells,
            excel_set_formula,
            excel_search,
            excel_update_range
        )
        
        print("\n=== 直接调用server.py中的方法 ===")
        
        # 问题1: excel_get_range 参数顺序问题
        print("\n--- 问题1: excel_get_range 参数顺序测试 ---")
        
        # 测试正常调用
        result1 = excel_get_range(test_file, "Sheet1!A1:C3")
        print(f"正常调用结果: {result1['success']}")
        
        # 测试start_cell和end_cell参数
        result2 = excel_get_range(
            test_file,
            "Sheet1!A1:C3",
            start_cell="A1",
            end_cell="C3"
        )
        print(f"start_cell/end_cell调用结果: {result2['success']}")
        
        # 测试颠倒的顺序
        result3 = excel_get_range(
            test_file,
            "Sheet1!A1:C3",
            start_cell="C3",  # 颠倒
            end_cell="A1"
        )
        print(f"颠倒顺序调用结果: {result3['success']}")
        
        # 问题2: format_cells 缺少参数处理
        print("\n--- 问题2: format_cells 缺少参数测试 ---")
        
        # 测试正常调用
        result4 = excel_format_cells(
            test_file,
            "Sheet1",
            "A1:C3",
            formatting={"bold": True, "font_color": "FF0000"}
        )
        print(f"正常调用结果: {result4['success']}")
        
        # 测试缺少formatting参数
        result5 = excel_format_cells(test_file, "Sheet1", "A1:C3")
        print(f"缺少formatting结果: {result5['success']}")
        if not result5['success']:
            print(f"错误信息: {result5['message']}")
        
        # 测试空的formatting
        result6 = excel_format_cells(
            test_file,
            "Sheet1", 
            "A1:C3",
            formatting={}
        )
        print(f"空的formatting结果: {result6['success']}")
        
        # 问题3: excel_set_formula 缺少formula参数
        print("\n--- 问题3: excel_set_formula 缺少参数测试 ---")
        
        # 测试正常调用
        result7 = excel_set_formula(test_file, "Sheet1", "A4", "=SUM(B1:B3)")
        print(f"正常调用结果: {result7['success']}")
        
        # 测试空字符串formula
        result8 = excel_set_formula(test_file, "Sheet1", "A5", "")
        print(f"空字符串结果: {result8['success']}")
        if not result8['success']:
            print(f"错误信息: {result8['message']}")
        
        # 测试None formula
        result9 = excel_set_formula(test_file, "Sheet1", "A6", None)
        print(f"None formula结果: {result9['success']}")
        
        # 问题4: excel_search 参数混淆
        print("\n--- 问题4: excel_search 参数混淆测试 ---")
        
        # 测试range中指定sheet_name
        result10 = excel_search(test_file, "Sheet1!A1:C3", "Alice")
        print(f"range中指定sheet_name结果: {result10['success']}")
        
        # 测试通过sheet_name参数指定
        result11 = excel_search(
            test_file,
            "A1:C3",
            "Alice",
            sheet_name="Sheet1"
        )
        print(f"sheet_name参数指定结果: {result11['success']}")
        
        # 测试重复指定sheet_name
        result12 = excel_search(
            test_file,
            "Sheet1!A1:C3",  # range中已指定
            "Alice",
            sheet_name="Sheet1"  # 参数中又指定
        )
        print(f"重复指定sheet_name结果: {result12['success']}")
        
        # 问题5: excel_update_range 数据格式处理
        print("\n--- 问题5: excel_update_range 数据格式测试 ---")
        
        # 测试正常数据格式
        normal_data = [["New", "Data", "Here"], ["Test", 123, "OK"]]
        result13 = excel_update_range(
            test_file,
            "Sheet1!A5:C6",
            normal_data
        )
        print(f"正常数据结果: {result13['success']}")
        
        # 测试错误的数据格式
        wrong_data = ["not", "a", "list", "of", "lists"]
        result14 = excel_update_range(
            test_file,
            "Sheet1!A8:A8",
            wrong_data
        )
        print(f"错误数据格式结果: {result14['success']}")
        if not result14['success']:
            print(f"错误信息: {result14['message']}")
        
        # 测试数据维度不匹配
        mismatch_data = [["Only", "one", "row"]]
        result15 = excel_update_range(
            test_file,
            "Sheet1!A10:C12",  # 期望3行数据
            mismatch_data  # 只有1行数据
        )
        print(f"数据维度不匹配结果: {result15['success']}")
        
    except Exception as e:
        print(f"❌ 直接调用测试失败: {e}")
        import traceback
        traceback.print_exc()
    finally:
        os.unlink(test_file)

if __name__ == "__main__":
    print("开始直接调用server.py中的方法复现API问题...")
    test_direct_server_calls()