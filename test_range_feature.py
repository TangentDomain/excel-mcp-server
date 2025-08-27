#!/usr/bin/env python3
"""
Excel正则搜索范围功能测试

验证 range_expression 参数的不同使用方式
"""

import json

def test_range_search():
    """测试范围搜索功能的各种用法"""
    
    print("🧪 Excel正则搜索 - 范围功能测试")
    print("=" * 50)
    
    test_cases = [
        {
            "name": "测试1: 搜索整个文件 (无范围限制)",
            "params": {
                "file_path": "d:\\excel-mcp-server\\test_range_search.xlsx",
                "pattern": "@"
            },
            "expected": 5
        },
        {
            "name": "测试2: 使用完整范围表达式 (Sheet1!A1:C6)",
            "params": {
                "file_path": "d:\\excel-mcp-server\\test_range_search.xlsx",
                "pattern": "@",
                "range_expression": "Sheet1!A1:C6"
            },
            "expected": 2
        },
        {
            "name": "测试3: 使用分离格式 (range + sheet_name)",
            "params": {
                "file_path": "d:\\excel-mcp-server\\test_range_search.xlsx",
                "pattern": "@",
                "range_expression": "A1:C6",
                "sheet_name": "Sheet1"
            },
            "expected": 2
        },
        {
            "name": "测试4: 扩大范围 (A1:D8) - 应该包含更多匹配",
            "params": {
                "file_path": "d:\\excel-mcp-server\\test_range_search.xlsx",
                "pattern": "@",
                "range_expression": "A1:D8",
                "sheet_name": "Sheet1"
            },
            "expected": 4
        },
        {
            "name": "测试5: 小范围 (B5:C5) - 只包含第5行的B和C列",
            "params": {
                "file_path": "d:\\excel-mcp-server\\test_range_search.xlsx",
                "pattern": "@",
                "range_expression": "B5:C5",
                "sheet_name": "Sheet1"
            },
            "expected": 2
        }
    ]
    
    # 这里应该调用 MCP 工具，但为了演示我们只打印测试用例
    for test_case in test_cases:
        print(f"\n📋 {test_case['name']}")
        print(f"   参数: {json.dumps(test_case['params'], ensure_ascii=False, indent=8)}")
        print(f"   预期匹配数: {test_case['expected']}")
        print(f"   状态: ✅ 测试用例已定义")

if __name__ == "__main__":
    test_range_search()
