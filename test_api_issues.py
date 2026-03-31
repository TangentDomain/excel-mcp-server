#!/usr/bin/env python3
"""
测试监工报告中的5个API问题
"""

import subprocess
import json
import os

def test_api_issues():
    """测试5个API问题"""
    
    # 测试1: read_data_from_excel 范围查询 - 参数顺序可能颠倒
    print("=== 测试1: read_data_from_excel 范围查询 ===")
    try:
        result = subprocess.run([
            'python3', '-m', 'excel_mcp_server_fastmcp', 'read_data_from_excel',
            '--filepath', 'test_api_issues.xlsx',
            '--sheet_name', 'Sheet1',
            '--start_cell', 'C1',
            '--end_cell', 'C3'
        ], capture_output=True, text=True, timeout=10)
        print(f"正常调用: {result.returncode}")
        print(f"输出: {result.stdout}")
        if result.stderr:
            print(f"错误: {result.stderr}")
    except Exception as e:
        print(f"异常: {e}")
    
    # 测试参数颠倒
    try:
        result = subprocess.run([
            'python3', '-m', 'excel_mcp_server_fastmcp', 'read_data_from_excel',
            '--filepath', 'test_api_issues.xlsx',
            '--sheet_name', 'Sheet1',
            '--start_cell', 'C3',
            '--end_cell', 'C1'
        ], capture_output=True, text=True, timeout=10)
        print(f"参数颠倒调用: {result.returncode}")
        print(f"输出: {result.stdout}")
        if result.stderr:
            print(f"错误: {result.stderr}")
    except Exception as e:
        print(f"异常: {e}")
    
    print("\n" + "="*50 + "\n")

    # 测试2: format_range - 缺少必要参数
    print("=== 测试2: format_range 缺少参数 ===")
    try:
        result = subprocess.run([
            'python3', '-m', 'excel_mcp_server_fastmcp', 'format_range',
            '--filepath', 'test_api_issues.xlsx',
            '--sheet_name', 'Sheet1',
            '--start_cell', 'A1',
            '--end_cell', 'A5'
        ], capture_output=True, text=True, timeout=10)
        print(f"缺少bold等参数: {result.returncode}")
        print(f"输出: {result.stdout}")
        if result.stderr:
            print(f"错误: {result.stderr}")
    except Exception as e:
        print(f"异常: {e}")
    
    print("\n" + "="*50 + "\n")

    # 测试3: apply_formula - 缺少formula参数
    print("=== 测试3: apply_formula 缺少参数 ===")
    try:
        result = subprocess.run([
            'python3', '-m', 'excel_mcp_server_fastmcp', 'apply_formula',
            '--filepath', 'test_api_issues.xlsx',
            '--sheet_name', 'Sheet1',
            '--cell', 'A1'
            # 故意不提供 --formula
        ], capture_output=True, text=True, timeout=10)
        print(f"缺少formula参数: {result.returncode}")
        print(f"输出: {result.stdout}")
        if result.stderr:
            print(f"错误: {result.stderr}")
    except Exception as e:
        print(f"异常: {e}")
    
    print("\n" + "="*50 + "\n")

    # 测试4: read_data_from_excel 搜索逻辑
    print("=== 测试4: read_data_from_excel 搜索逻辑 ===")
    try:
        result = subprocess.run([
            'python3', '-m', 'excel_mcp_server_fastmcp', 'read_data_from_excel',
            '--filepath', 'test_api_issues.xlsx',
            '--sheet_name', 'Sheet1',
            '--start_cell', 'A1',
            '--end_cell', 'C5',
            '--search', 'b'  # 搜索参数
        ], capture_output=True, text=True, timeout=10)
        print(f"带搜索: {result.returncode}")
        print(f"输出: {result.stdout}")
        if result.stderr:
            print(f"错误: {result.stderr}")
    except Exception as e:
        print(f"异常: {e}")
    
    print("\n" + "="*50 + "\n")

    # 测试5: write_data_to_excel 数据格式不匹配
    print("=== 测试5: write_data_to_excel 数据格式不匹配 ===")
    try:
        # 错误的数据格式 - 不是list of lists
        result = subprocess.run([
            'python3', '-m', 'excel_mcp_server_fastmcp', 'write_data_to_excel',
            '--filepath', 'test_api_issues.xlsx',
            '--sheet_name', 'Sheet1',
            '--start_cell', 'D1',
            '--data', '["wrong", "format"]'  # 错误格式
        ], capture_output=True, text=True, timeout=10)
        print(f"错误数据格式: {result.returncode}")
        print(f"输出: {result.stdout}")
        if result.stderr:
            print(f"错误: {result.stderr}")
    except Exception as e:
        print(f"异常: {e}")

if __name__ == "__main__":
    test_api_issues()