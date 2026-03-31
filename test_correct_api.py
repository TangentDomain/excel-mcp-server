#!/usr/bin/env python3
"""
使用正确的函数名测试API问题
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
                range_expression="C1:B3"  # 故意颠倒范围
            )
            print(f"结果: {result}")
            if "error" in str(result).lower():
                issues_found.append("excel_get_range范围参数顺序问题")
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("excel_get_range范围参数顺序问题")
        
        # 测试2: excel_format_cells 缺少必要参数
        print("\n2. 测试 excel_format_cells 缺少必要参数...")
        try:
            result = excel_ops.format_cells(
                file_path=excel_file,
                sheet_name=sheet_name,
                range_address="A1:A5",
                # 故意不传bold等参数
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
                # 故意不传formula参数
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
            result = excel_ops.get_range(
                file_path=excel_file,
                # 故意不传sheet_name，看看会不会和其他参数混淆
                range_expression="A1:B3"
            )
            print(f"结果: {result}")
            if "error" in str(result).lower() or "missing" in str(result).lower():
                issues_found.append("excel_get_range sheet_name参数混淆")
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("excel_get_range sheet_name参数混淆")
        
        # 测试5: excel_write_only_override 数据格式不匹配
        print("\n5. 测试 excel_write_only_override 数据格式不匹配...")
        try:
            result = excel_ops.update_range(
                file_path=excel_file,
                sheet_name=sheet_name,
                range_expression="D1",
                # 故意传错误的数据格式
                data="single_string_instead_of_list"
            )
            print(f"结果: {result}")
            if "error" in str(result).lower():
                issues_found.append("excel_update_range数据格式处理不当")
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("excel_update_range数据格式处理不当")
        
    except Exception as e:
        print(f"测试过程中发生错误: {e}")
    
    # 输出总结
    print("\n" + "=" * 60)
    print("问题总结:")
    if issues_found:
        for i, issue in enumerate(issues_found, 1):
            print(f"{i}. {issue}")
    else:
        print("所有API都工作正常")
    
    # 清理临时文件
    os.unlink(excel_file)
    
    return issues_found

if __name__ == "__main__":
    issues = test_api_issues()
    print(f"\n发现 {len(issues)} 个API问题需要修复")