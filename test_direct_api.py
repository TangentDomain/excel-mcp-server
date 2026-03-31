#!/usr/bin/env python3
"""
直接测试API函数来复现5个问题
"""
import sys
from pathlib import Path

# 添加项目路径
project_root = Path(__file__).parent
src_path = project_root / "src"
sys.path.insert(0, str(src_path))

# 导入MCP函数直接测试（绕过API类）
from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations

def test_api_issues():
    """测试5个API问题"""
    ops = ExcelOperations()
    test_file = "test_api_issues.xlsx"
    
    print("🧪 测试监工报告中的5个API问题")
    print("=" * 60)
    
    # 测试1: excel_get_range 范围查询 - 参数顺序可能颠倒
    print("\n🔍 测试1: excel_get_range 范围查询")
    print("-" * 40)
    try:
        # 正常调用
        result1 = ops.get_range(test_file, "Sheet1", "C1:C3")
        print(f"✅ 正常调用 C1:C3: {len(result1)} 行")
        
        # 参数颠倒 - C3:C1 
        result2 = ops.get_range(test_file, "Sheet1", "C3:C1")
        print(f"⚠️ 参数颠倒 C3:C1: {len(result2)} 行")
        
        if len(result1) != len(result2):
            print("❌ 问题：参数颠倒时结果不一致！")
        else:
            print("✅ 参数颠倒时结果正常")
            
    except Exception as e:
        print(f"❌ get_range测试失败: {e}")
    
    # 测试2: excel_format_cells 缺少格式参数
    print("\n🎨 测试2: excel_format_cells 缺少格式参数")
    print("-" * 40)
    try:
        # 检查实际的excel_format_cells函数签名，故意不提供formatting参数
        result = ops.format_cells(
            file_path=test_file,
            sheet_name="Sheet1",
            range="A1:A5"
            # 故意不提供formatting参数
        )
        print(f"⚠️ 缺少格式参数: {result}")
        
        if "error" in str(result).lower() or "fail" in str(result).lower():
            print("❌ 问题：缺少参数时报错不合理")
        else:
            print("✅ 缺少参数时处理正常")
            
    except Exception as e:
        print(f"❌ format_cells测试失败: {e}")
    
    # 测试3: excel_set_formula 缺少formula参数
    print("\n📝 测试3: excel_set_formula 缺少参数")
    print("-" * 40)
    try:
        # 检查实际的excel_set_formula函数签名，故意不提供formula参数
        result = ops.set_formula(
            file_path=test_file,
            sheet_name="Sheet1", 
            cell_range="A1"
            # 故意不提供 formula 参数
        )
        print(f"⚠️ 缺少formula参数: {result}")
        
        if "error" in str(result).lower() or "missing" in str(result).lower():
            print("✅ 缺少必填参数时正确报错")
        else:
            print("❌ 问题：缺少必填参数时未报错")
            
    except Exception as e:
        print(f"❌ set_formula测试失败: {e}")
    
    # 测试4: excel_search 搜索逻辑
    print("\n🔍 测试4: excel_search 搜索逻辑")
    print("-" * 40)
    try:
        # 测试excel_search函数，使用正确的参数名
        result = ops.search(
            file_path=test_file,
            pattern="b"  # 搜索参数
        )
        print(f"⚠️ 带搜索调用: {result}")
        
        # 检查是否正确处理搜索参数
        if result and len(result) > 0:
            print("✅ 搜索功能正常工作")
        else:
            print("⚠️ 搜索功能可能有问题")
            
    except Exception as e:
        print(f"❌ 搜索测试失败: {e}")
    
    # 测试5: excel_update_range 数据格式不匹配
    print("\n✍️ 测试5: excel_update_range 数据格式")
    print("-" * 40)
    try:
        # 错误的数据格式 - 不是二维列表
        wrong_data = ["wrong", "format"]  # 应该是 [["wrong", "format"]] 或 [["wrong"], ["format"]]
        result = ops.update_range(
            file_path=test_file,
            range_expression="Sheet1!E1:F2",  # 必须包含工作表名
            data=wrong_data
            # 故意不提供insert_mode参数等
        )
        print(f"⚠️ 错误数据格式: {result}")
        
        if "error" in str(result).lower() or "invalid" in str(result).lower():
            print("✅ 错误数据格式时正确报错")
        else:
            print("❌ 问题：错误数据格式时未报错")
            
    except Exception as e:
        print(f"❌ update_range测试失败: {e}")

if __name__ == "__main__":
    test_api_issues()