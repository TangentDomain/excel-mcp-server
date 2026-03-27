#!/usr/bin/env python3
"""
简化MCP测试脚本
"""
import sys
import os
import json
from pathlib import Path

# 添加路径
sys.path.insert(0, '/root/.openclaw/workspace/excel-mcp-server/src')

# 直接导入函数
try:
    from excel_mcp_server_fastmcp.server import (
        excel_list_sheets, excel_get_headers, excel_get_range,
        excel_query, excel_find_last_row, excel_batch_insert_rows,
        excel_delete_rows, excel_describe_table
    )
    print("✅ 成功导入所有函数")
except ImportError as e:
    print(f"❌ 导入失败: {e}")
    sys.exit(1)

def test_function(func_name, func, args, description):
    """测试单个函数"""
    print(f"\n{description} ({func_name})")
    print("-" * 40)
    
    try:
        result = func(**args)
        print(f"✅ 成功: {result.get('message', '操作成功')}")
        return True
    except Exception as e:
        print(f"❌ 失败: {str(e)}")
        return False

def main():
    test_file = '/tmp/excel_mcp_test_fs5p18el/test_game_data.xlsx'
    
    tests = [
        {
            'func': excel_list_sheets,
            'args': {'file_path': test_file},
            'description': '列出工作表'
        },
        {
            'func': excel_get_headers,
            'args': {'file_path': test_file, 'sheet_name': '角色'},
            'description': '获取表头'
        },
        {
            'func': excel_get_range,
            'args': {'file_path': test_file, 'range': '角色!A1:E6'},
            'description': '获取数据范围'
        },
        {
            'func': excel_query,
            'args': {'file_path': test_file, 'query_expression': 'SELECT * FROM 角色 WHERE 职业="战士"'},
            'description': '简单查询'
        },
        {
            'func': excel_query,
            'args': {'file_path': test_file, 'query_expression': 'SELECT * FROM 角色 JOIN 技能 ON 职业=职业限制'},
            'description': 'JOIN查询'
        },
        {
            'func': excel_query,
            'args': {'file_path': test_file, 'query_expression': 'SELECT 职业, COUNT(*) as 人数 FROM 角色 GROUP BY 职业'},
            'description': 'GROUP BY查询'
        },
        {
            'func': excel_find_last_row,
            'args': {'file_path': test_file, 'sheet_name': '角色'},
            'description': '查找最后一行'
        },
        {
            'func': excel_describe_table,
            'args': {'file_path': test_file, 'sheet_name': '角色'},
            'description': '描述表结构'
        }
    ]
    
    print("🧪 MCP真实验证开始（第145轮）")
    print("=" * 60)
    
    results = []
    passed = 0
    failed = 0
    
    for i, test in enumerate(tests, 1):
        success = test_function(
            test['func'].__name__, 
            test['func'], 
            test['args'],
            test['description']
        )
        
        if success:
            passed += 1
        else:
            failed += 1
            
        results.append({
            'test': test['description'],
            'success': success
        })
    
    print("\n" + "=" * 60)
    print(f"📊 测试结果: {passed} 通过 / {failed} 失败")
    
    # 保存结果
    with open('/tmp/mcp_test_results_145.json', 'w', encoding='utf-8') as f:
        json.dump({
            'round': 145,
            'tests': results,
            'summary': {'passed': passed, 'failed': failed}
        }, f, indent=2, ensure_ascii=False)
    
    return passed, failed

if __name__ == "__main__":
    passed, failed = main()
    if failed > 0:
        print(f"\n⚠️ 发现 {failed} 个问题，需要修复")
        sys.exit(1)
    else:
        print(f"\n🎉 所有测试通过！")
        sys.exit(0)