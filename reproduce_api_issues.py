#!/usr/bin/env python3
"""
复现监工发现的5个API问题
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

import tempfile
import json
from excel_mcp_server_fastmcp.server import main

def create_test_excel():
    """创建测试用的Excel文件"""
    import pandas as pd
    
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
    
    print(f"测试文件: {excel_file}")
    
    # 测试1: read_data_from_excel 范围查询
    print("\n1. 测试 read_data_from_excel 范围查询...")
    try:
        # 故意颠倒参数顺序测试
        result = main([
            "read_data_from_excel", 
            excel_file,
            "Sheet1",
            "C1",  # start_cell
            "B3"   # end_cell 
        ])
        print(f"结果: {result}")
    except Exception as e:
        print(f"错误: {e}")
    
    # 测试2: format_range 缺少参数
    print("\n2. 测试 format_range 缺少参数...")
    try:
        result = main([
            "format_range",
            excel_file,
            "Sheet1",
            "A1:A5",
            # 故意不传bold参数
        ])
        print(f"结果: {result}")
    except Exception as e:
        print(f"错误: {e}")
    
    # 测试3: apply_formula 缺少formula参数
    print("\n3. 测试 apply_formula 缺少参数...")
    try:
        result = main([
            "apply_formula",
            excel_file,
            "Sheet1",
            "A1",
            # 故意不传formula参数
        ])
        print(f"结果: {result}")
    except Exception as e:
        print(f"错误: {e}")
    
    # 测试4: read_data_from_excel 搜索逻辑
    print("\n4. 测试 read_data_from_excel 搜索逻辑...")
    try:
        result = main([
            "read_data_from_excel",
            excel_file,
            # 故意不传sheet_name，看看会不会和其他参数混淆
            "A1:B3"
        ])
        print(f"结果: {result}")
    except Exception as e:
        print(f"错误: {e}")
    
    # 测试5: write_data_to_excel 数据格式
    print("\n5. 测试 write_data_to_excel 数据格式...")
    try:
        # 故意传错误的数据格式
        result = main([
            "write_data_to_excel",
            excel_file,
            "Sheet1",
            "D1",
            # 故意传字符串而不是list of lists
            "single_string_instead_of_list"
        ])
        print(f"结果: {result}")
    except Exception as e:
        print(f"错误: {e}")
    
    # 清理临时文件
    os.unlink(excel_file)

if __name__ == "__main__":
    test_api_issues()