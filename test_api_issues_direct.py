#!/usr/bin/env python3
"""
Reproduce the 5 API issues using the correct function names
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

import tempfile
import pandas as pd
from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations

def create_test_excel():
    """创建测试用的Excel文件"""
    # 创建测试数据
    data = {
        'A': [1, 2, 3, 4, 5],
        'B': ['a', 'b', 'c', 'd', 'e'],
        'C': [10.5, 20.5, 30.5, 40.5, 50.5]
    }
    df = pd.DataFrame(data)
    
    # 创建临时文件
    temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_file.close()
    
    # 写入Excel
    df.to_excel(temp_file.name, index=False)
    return temp_file.name

def test_api_issues():
    """测试5个API问题"""
    excel_file = create_test_excel()
    sheet_name = "Sheet1"
    
    print(f"测试文件: {excel_file}")
    print("=" * 60)
    
    excel_ops = ExcelOperations()
    
    issues_found = []
    
    try:
        # 测试1: excel_get_range 范围查询（参数顺序颠倒）
        print("\n1. 测试 excel_get_range 范围查询（参数顺序颠倒）...")
        try:
            result = excel_ops.get_range(
                file_path=excel_file,
                sheet_name=sheet_name,
                range_expression="A1:C3"  # 正常顺序
            )
            print(f"正常顺序结果: {result}")
            
            # 颠倒的range_expression应该也能工作，或者给出明确的错误
            result2 = excel_ops.get_range(
                file_path=excel_file,
                sheet_name=sheet_name,
                range_expression="C3:A1"  # 颠倒顺序
            )
            print(f"颠倒顺序结果: {result2}")
            
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("excel_get_range参数顺序处理不当")
        
        # 测试2: excel_format_cells 缺少必要参数
        print("\n2. 测试 excel_format_cells 缺少必要参数...")
        try:
            result = excel_ops.format_cells(
                file_path=excel_file,
                sheet_name=sheet_name,
                range_address="A1:A5",
                # 故意不传bold等参数，看看是否优雅处理
            )
            print(f"结果: {result}")
            if "error" in str(result).lower():
                issues_found.append("excel_format_cells缺少参数处理不当")
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("excel_format_cells缺少参数处理不当")
        
        # 测试3: excel_set_formula 缺少formula参数
        print("\n3. 测试 excel_set_formula 缺少formula参数...")
        try:
            result = excel_ops.set_formula(
                file_path=excel_file,
                sheet_name=sheet_name,
                cell="A1",
                # 故意不传formula参数，应该优雅校验
            )
            print(f"结果: {result}")
            if "error" in str(result).lower():
                issues_found.append("excel_set_formula缺少参数校验")
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("excel_set_formula缺少参数校验")
        
        # 测试4: excel_get_range 搜索逻辑（sheet_name混淆）
        print("\n4. 测试 excel_get_range 搜索逻辑（sheet_name混淆）...")
        try:
            # 模拟sheet_name与其他参数混淆的情况
            result = excel_ops.get_range(
                file_path=excel_file,
                range_expression="A1:B3"  # 不传sheet_name
            )
            print(f"不传sheet_name结果: {result}")
            if "error" in str(result).lower() or "missing" in str(result).lower():
                issues_found.append("excel_get_range sheet_name参数混淆")
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("excel_get_range sheet_name参数混淆")
        
        # 测试5: excel_write_only_override 数据格式不匹配
        print("\n5. 测试 excel_write_only_override 数据格式不匹配...")
        try:
            # 传入错误的数据格式
            result = excel_ops.write_only_override(
                file_path=excel_file,
                sheet_name=sheet_name,
                data="invalid_format",  # 应该是list of lists
                range_address="A1:A5"
            )
            print(f"错误格式结果: {result}")
            if "error" in str(result).lower():
                issues_found.append("excel_write_only_override数据格式校验不当")
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("excel_write_only_override数据格式校验不当")
    
    finally:
        # 清理临时文件
        os.unlink(excel_file)
    
    print("\n" + "=" * 60)
    print("问题总结:")
    if issues_found:
        for i, issue in enumerate(issues_found, 1):
            print(f"{i}. {issue}")
    else:
        print("✅ 所有API测试通过，未发现问题")
    
    return issues_found

if __name__ == "__main__":
    test_api_issues()