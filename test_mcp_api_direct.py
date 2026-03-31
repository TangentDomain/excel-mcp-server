#!/usr/bin/env python3
"""
使用MCP协议直接测试API兼容性问题
"""

import json
import sys
import tempfile
import os
from pathlib import Path

# 添加项目路径
sys.path.insert(0, 'src')

def create_test_excel():
    """创建测试Excel文件"""
    test_file = "/tmp/test_mcp_api.xlsx"
    
    test_data = [
        ["ID", "姓名", "等级", "经验值"],
        [1, "张三", 10, 1500],
        [2, "李四", 20, 3500],
        [3, "王五", 30, 6000],
        [4, "赵六", 40, 10000]
    ]
    
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        
        for row_idx, row_data in enumerate(test_data):
            for col_idx, cell_value in enumerate(row_data):
                ws.cell(row=row_idx+1, column=col_idx+1, value=cell_value)
        
        wb.save(test_file)
        return test_file
    except ImportError:
        return None

def test_mcp_api_directly():
    """直接测试MCP API调用"""
    
    print("🚀 开始直接MCP API测试")
    print("=" * 60)
    
    # 导入必要的模块
    try:
        from excel_mcp_server_fastmcp.server import FastMCP, _wrap, _fail, _validate_path
        from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations
        import openpyxl
    except ImportError as e:
        print(f"❌ 导入失败: {e}")
        return False
    
    test_file = create_test_excel()
    if not test_file:
        print("❌ 无法创建测试文件")
        return False
    
    # 创建MCP服务器实例（用于测试，不实际启动服务）
    mcp = FastMCP("test-server")
    
    results = {}
    
    # 测试1: excel_get_range 参数处理
    print("\n🔍 测试1: excel_get_range 参数处理")
    print("-" * 40)
    
    try:
        # 测试正常调用
        from excel_mcp_server_fastmcp.server import excel_get_range
        result1 = excel_get_range(test_file, "A1:C4")
        
        if result1.get('success'):
            print("✅ 正常参数调用成功")
            results['test1_normal'] = "PASS"
        else:
            print(f"❌ 正常参数调用失败: {result1.get('message')}")
            results['test1_normal'] = "FAIL"
            
        # 测试参数顺序问题 - 检查函数签名是否正确
        import inspect
        sig = inspect.signature(excel_get_range)
        params = list(sig.parameters.keys())
        
        # 期望的参数顺序应该是: file_path, range, include_formatting, sheet_name, start_cell, end_cell
        expected_order = ['file_path', 'range', 'include_formatting', 'sheet_name', 'start_cell', 'end_cell']
        if params == expected_order[:len(params)]:
            print("✅ 参数顺序正确")
            results['test1_order'] = "PASS"
        else:
            print(f"⚠️ 参数顺序可能有问题: {params}")
            results['test1_order'] = "WARNING"
            
    except Exception as e:
        print(f"❌ 测试异常: {e}")
        results['test1_normal'] = "ERROR"
        results['test1_order'] = "ERROR"
    
    # 测试2: excel_format_cells 参数处理
    print("\n🔍 测试2: excel_format_cells 参数处理")
    print("-" * 40)
    
    try:
        from excel_mcp_server_fastmcp.server import excel_format_cells
        
        # 测试正常调用
        formatting = {"bold": True, "font_color": "FF0000"}
        result2 = excel_format_cells(test_file, "Sheet1", "A1:C1", formatting=formatting)
        
        if result2.get('success'):
            print("✅ format_cells 正常调用成功")
            results['test2_normal'] = "PASS"
        else:
            print(f"❌ format_cells 正常调用失败: {result2.get('message')}")
            results['test2_normal'] = "FAIL"
            
        # 测试缺少bold参数时的处理
        result2_missing = excel_format_cells(test_file, "Sheet1", "A2:C2", formatting={})
        
        if result2_missing.get('success'):
            print("✅ 缺少bold参数处理成功")
            results['test2_missing_bold'] = "PASS"
        else:
            print(f"⚠️ 缺少bold参数报错: {result2_missing.get('message')}")
            results['test2_missing_bold'] = "WARNING"
            
    except Exception as e:
        print(f"❌ 测试异常: {e}")
        results['test2_normal'] = "ERROR"
        results['test2_missing_bold'] = "ERROR"
    
    # 测试3: excel_set_formula 参数验证
    print("\n🔍 测试3: excel_set_formula 参数验证")
    print("-" * 40)
    
    try:
        from excel_mcp_server_fastmcp.server import excel_set_formula
        
        # 测试正常调用
        result3_normal = excel_set_formula(test_file, "Sheet1", "D1", "SUM(A1:C1)")
        
        if result3_normal.get('success'):
            print("✅ set_formula 正常调用成功")
            results['test3_normal'] = "PASS"
        else:
            print(f"❌ set_formula 正常调用失败: {result3_normal.get('message')}")
            results['test3_normal'] = "FAIL"
        
        # 测试缺少formula参数 - 这里应该通过函数签名检查
        # 由于Python函数不能缺少必需参数，我们需要模拟这种情况
        try:
            # 这会抛出TypeError因为缺少必需参数
            result3_missing = excel_set_formula(test_file, "Sheet1", "D2")
            print("❌ 缺少formula参数应该报错但没有")
            results['test3_missing_formula'] = "FAIL"
        except TypeError as e:
            print("✅ 缺少formula参数正确报错")
            results['test3_missing_formula'] = "PASS"
        except Exception as e:
            print(f"⚠️ 缺少formula参数报错类型: {type(e).__name__}")
            results['test3_missing_formula'] = "PASS"
            
    except Exception as e:
        print(f"❌ 测试异常: {e}")
        results['test3_normal'] = "ERROR"
        results['test3_missing_formula'] = "ERROR"
    
    # 测试4: excel_update_range 数据格式处理
    print("\n🔍 测试4: excel_update_range 数据格式处理")
    print("-" * 40)
    
    try:
        from excel_mcp_server_fastmcp.server import excel_update_range
        
        # 测试正常数据格式
        test_data = [[50, 60, 70], [80, 90, 100]]
        result4_normal = excel_update_range(test_file, "Sheet1!A5:C6", test_data)
        
        if result4_normal.get('success'):
            print("✅ update_range 正常数据格式调用成功")
            results['test4_normal'] = "PASS"
        else:
            print(f"❌ update_range 正常数据格式调用失败: {result4_normal.get('message')}")
            results['test4_normal'] = "FAIL"
        
        # 测试空数据格式
        try:
            result4_empty = excel_update_range(test_file, "Sheet1!A8:C10", [])
            if result4_empty.get('success'):
                print("✅ 空数据格式处理成功")
                results['test4_empty_data'] = "PASS"
            else:
                print(f"❌ 空数据格式处理失败: {result4_empty.get('message')}")
                results['test4_empty_data'] = "FAIL"
        except Exception as e:
            print(f"⚠️ 空数据格式处理异常: {e}")
            results['test4_empty_data'] = "WARNING"
            
    except Exception as e:
        print(f"❌ 测试异常: {e}")
        results['test4_normal'] = "ERROR"
        results['test4_empty_data'] = "ERROR"
    
    # 测试5: excel_search 搜索逻辑
    print("\n🔍 测试5: excel_search 搜索逻辑")
    print("-" * 40)
    
    try:
        from excel_mcp_server_fastmcp.server import excel_search
        
        # 测试正常搜索
        result5_normal = excel_search(test_file, "张三")
        
        if result5_normal.get('success'):
            print("✅ search 正常调用成功")
            results['test5_normal'] = "PASS"
        else:
            print(f"❌ search 正常调用失败: {result5_normal.get('message')}")
            results['test5_normal'] = "FAIL"
        
        # 测试sheet_name参数处理
        result5_with_sheet = excel_search(test_file, "李四", sheet_name="Sheet1")
        
        if result5_with_sheet.get('success'):
            print("✅ search 带sheet_name参数调用成功")
            results['test5_with_sheet'] = "PASS"
        else:
            print(f"❌ search 带sheet_name参数调用失败: {result5_with_sheet.get('message')}")
            results['test5_with_sheet'] = "FAIL"
            
    except Exception as e:
        print(f"❌ 测试异常: {e}")
        results['test5_normal'] = "ERROR"
        results['test5_with_sheet'] = "ERROR"
    
    # 输出测试结果
    print("\n" + "=" * 60)
    print("📊 测试结果汇总")
    print("=" * 60)
    
    passed = 0
    total = 0
    
    for test_name, result in results.items():
        total += 1
        status = "✅ PASS" if result == "PASS" else "❌ FAIL" if result == "FAIL" else "⚠️ WARNING" if result == "WARNING" else "❌ ERROR"
        print(f"{status} {test_name}")
        if result == "PASS":
            passed += 1
    
    print(f"\n📈 通过率: {passed}/{total} ({passed/total*100:.1f}%)")
    
    if passed >= total * 0.8:  # 80%通过率
        print("🎉 大部分测试通过！")
        return True
    else:
        print("⚠️ 发现多个问题需要修复")
        return False

if __name__ == "__main__":
    success = test_mcp_api_directly()
    exit(0 if success else 1)