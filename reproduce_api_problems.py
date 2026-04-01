#!/usr/bin/env python3
"""
复现监工发现的5个API问题
"""

import asyncio
import json
import sys
from pathlib import Path

# 添加路径以便导入
sys.path.append('src')
from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations
from excel_mcp_server_fastmcp.server import main

async def test_api_problems():
    """测试5个监工发现的API问题"""
    
    print("=" * 60)
    print("开始复现监工发现的5个API问题")
    print("=" * 60)
    
    # 创建测试文件
    test_file = Path("test_api_issues.xlsx")
    if test_file.exists():
        test_file.unlink()
    
    # 创建一个简单的测试Excel文件
    print("📝 创建测试Excel文件...")
    
    ops = ExcelOperations()
    
    # 创建测试数据
    headers = ["ID", "Name", "Value", "Active"]
    test_data = [
        [1, "Item1", 100, True],
        [2, "Item2", 200, False],
        [3, "Item3", 300, True],
    ]
    
    # 写入测试数据
    await ops.create_workbook("test_api_issues.xlsx")
    await ops.create_worksheet("test_api_issues", "Sheet1")
    await ops.write_data_to_excel("test_api_issues.xlsx", "Sheet1", [headers], "A1")
    await ops.write_data_to_excel("test_api_issues.xlsx", "Sheet1", test_data, "A2")
    
    print("✅ 测试文件创建完成")
    print()
    
    # 测试1: read_data_from_excel 范围查询 - 参数顺序可能颠倒
    print("🔍 测试1: read_data_from_excel 范围查询（参数顺序验证）")
    try:
        # 测试正常参数顺序
        result1 = await ops.read_data_from_excel("test_api_issues.xlsx", "Sheet1", "A1:C3")
        print(f"   正常顺序 (A1:C3): {len(result1.get('data', []))} 行数据")
        
        # 测试可能的颠倒参数
        result2 = await ops.read_data_from_excel("test_api_issues.xlsx", "Sheet1", "C3:A1")
        print(f"   颠倒顺序 (C3:A1): {len(result2.get('data', []))} 行数据")
        
        if result1 == result2:
            print("   ⚠️  问题：参数顺序颠倒时结果相同，可能参数验证不严格")
        else:
            print("   ✅ 参数顺序处理正确")
            
    except Exception as e:
        print(f"   ❌ 错误: {e}")
    print()
    
    # 测试2: format_range - 缺少bold等必要参数时的处理
    print("🔍 测试2: format_range（参数缺失处理）")
    try:
        # 测试只传部分参数
        result = await ops.format_range(
            "test_api_issues.xlsx", 
            "Sheet1", 
            "A1:A3", 
            bold=True,
            # 故意不传font_size等参数
        )
        print("   ✅ 部分参数处理成功")
        
        # 测试不传任何样式参数
        result2 = await ops.format_range(
            "test_api_issues.xlsx", 
            "Sheet1", 
            "B1:B3"
        )
        print("   ✅ 无参数处理成功")
        
    except Exception as e:
        print(f"   ❌ 错误: {e}")
    print()
    
    # 测试3: apply_formula - 缺少formula参数时的处理
    print("🔍 测试3: apply_formula（必填参数校验）")
    try:
        # 故意不传formula参数
        result = await ops.apply_formula(
            "test_api_issues.xlsx", 
            "Sheet1", 
            "A1",
            # 故意不传formula参数
        )
        print("   ❌ 问题：缺少必填参数formula应该报错但没有")
        
    except Exception as e:
        print(f"   ✅ 正确报错: {e}")
    print()
    
    # 测试4: read_data_from_excel 搜索逻辑 - sheet_name参数混淆
    print("🔍 测试4: read_data_from_excel搜索逻辑（sheet_name参数验证）")
    try:
        # 测试不存在的sheet
        result = await ops.read_data_from_excel("test_api_issues.xlsx", "NonExistentSheet", "A1:C3")
        print("   ❌ 问题：不存在的sheet应该报错但没有")
        
    except Exception as e:
        print(f"   ✅ 正确报错: {e}")
        
    # 测试有效的sheet
    try:
        result = await ops.read_data_from_excel("test_api_issues.xlsx", "Sheet1", "A1:C3")
        print(f"   ✅ 有效sheet读取成功: {len(result.get('data', []))} 行数据")
        
    except Exception as e:
        print(f"   ❌ 有效sheet读取失败: {e}")
    print()
    
    # 测试5: write_data_to_excel - 数据格式不匹配处理
    print("🔍 测试5: write_data_to_excel（数据格式验证）")
    try:
        # 测试错误的数据格式（不是list of lists）
        invalid_data = "not a list"
        result = await ops.write_data_to_excel("test_api_issues.xlsx", "Sheet1", invalid_data, "D1")
        print("   ❌ 问题：错误的数据格式应该报错但没有")
        
    except Exception as e:
        print(f"   ✅ 正确报错: {e}")
        
    # 测试正确的数据格式
    try:
        valid_data = [["Test", "Data"]]
        result = await ops.write_data_to_excel("test_api_issues.xlsx", "Sheet1", valid_data, "D1")
        print("   ✅ 正确数据格式写入成功")
        
    except Exception as e:
        print(f"   ❌ 正确数据格式写入失败: {e}")
    print()
    
    # 清理测试文件
    if test_file.exists():
        test_file.unlink()
        print("🧹 测试文件已清理")
    
    print("=" * 60)
    print("复测完成")
    print("=" * 60)

if __name__ == "__main__":
    asyncio.run(test_api_problems())