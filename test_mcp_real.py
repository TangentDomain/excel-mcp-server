#!/usr/bin/env python3
"""
MCP真实验证脚本 - 测试12项核心功能
"""
import json
import sys
import os
import subprocess
from pathlib import Path

# 模拟MCP调用
def test_mcp_function(name, args):
    """模拟MCP函数调用"""
    cmd = [
        'python3', '-c', 
        f"""
import sys
sys.path.insert(0, 'src')
from excel_mcp_server_fastmcp.server import excel_{name.replace('-', '_')}
try:
    result = excel_{name.replace('-', '_')}(**{json.dumps(args)})
    print(json.dumps(result, indent=2, ensure_ascii=False))
except Exception as e:
    print(json.dumps({{
        'success': False,
        'message': str(e),
        'data': None,
        'meta': {{'error_code': 'EXECUTION_ERROR', 'error_details': str(e)}}
    }}, indent=2, ensure_ascii=False))
"""
    ]
    
    result = subprocess.run(cmd, cwd='/root/.openclaw/workspace/excel-mcp-server', 
                          capture_output=True, text=True, timeout=30)
    
    if result.returncode == 0:
        try:
            return json.loads(result.stdout)
        except json.JSONDecodeError:
            return {
                'success': False,
                'message': f'JSON解析失败: {result.stdout}',
                'data': None,
                'meta': {'error_code': 'JSON_ERROR'}
            }
    else:
        return {
            'success': False,
            'message': f'执行失败: {result.stderr}',
            'data': None,
            'meta': {'error_code': 'EXECUTION_ERROR'}
        }

def run_mcp_tests():
    """执行MCP真实验证"""
    test_file = '/tmp/excel_mcp_test_fs5p18el/test_game_data.xlsx'
    
    tests = [
        {
            'name': 'list_sheets',
            'args': {'file_path': test_file},
            'description': '列出工作表'
        },
        {
            'name': 'get_headers',
            'args': {'file_path': test_file, 'sheet_name': '角色'},
            'description': '获取表头'
        },
        {
            'name': 'get_range',
            'args': {'file_path': test_file, 'sheet_name': '角色', 'range_ref': 'A1:E6'},
            'description': '获取数据范围'
        },
        {
            'name': 'query',
            'args': {'file_path': test_file, 'sheet_name': '角色', 'sql': 'SELECT * FROM 角色 WHERE 职业="战士"'},
            'description': '简单查询'
        },
        {
            'name': 'query',
            'args': {'file_path': test_file, 'sheet_name': '角色', 'sql': 'SELECT * FROM 角色 JOIN 技能 ON 职业=职业限制'},
            'description': 'JOIN查询'
        },
        {
            'name': 'query',
            'args': {'file_path': test_file, 'sheet_name': '角色', 'sql': 'SELECT 职业, COUNT(*) as 人数 FROM 角色 GROUP BY 职业'},
            'description': 'GROUP BY查询'
        },
        {
            'name': 'query',
            'args': {'file_path': test_file, 'sheet_name': '技能', 'sql': 'SELECT * FROM 技能 WHERE 伤害 > 100'},
            'description': 'WHERE查询'
        },
        {
            'name': 'find_last_row',
            'args': {'file_path': test_file, 'sheet_name': '角色'},
            'description': '查找最后一行'
        },
        {
            'name': 'batch_insert_rows',
            'args': {
                'file_path': test_file, 
                'sheet_name': '角色',
                'data': [
                    {'ID': 6, '名称': '游侠', '职业': '游侠', '等级': 46, '属性': '自然'},
                    {'ID': 7, '名称': '圣骑士', '职业': '圣骑士', '等级': 52, '属性': '光'}
                ]
            },
            'description': '批量插入行'
        },
        {
            'name': 'delete_rows',
            'args': {
                'file_path': test_file,
                'sheet_name': '角色', 
                'condition': 'ID > 5'
            },
            'description': '删除行'
        },
        {
            'name': 'describe_table',
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
        print(f"\n{i}. {test['description']} ({test['name']})")
        print("-" * 40)
        
        try:
            result = test_mcp_function(test['name'], test['args'])
            
            if result.get('success', False):
                print(f"✅ 通过: {result.get('message', '成功')}")
                passed += 1
            else:
                print(f"❌ 失败: {result.get('message', '未知错误')}")
                failed += 1
                
            results.append({
                'test': test['description'],
                'result': result
            })
            
        except Exception as e:
            print(f"❌ 异常: {str(e)}")
            failed += 1
            results.append({
                'test': test['description'],
                'result': {'success': False, 'message': str(e)}
            })
    
    print("\n" + "=" * 60)
    print(f"📊 测试结果: {passed} 通过 / {failed} 失败")
    
    # 保存详细结果
    with open('/tmp/mcp_test_results_145.json', 'w', encoding='utf-8') as f:
        json.dump({
            'round': 145,
            'tests': results,
            'summary': {'passed': passed, 'failed': failed}
        }, f, indent=2, ensure_ascii=False)
    
    return passed, failed

if __name__ == "__main__":
    try:
        passed, failed = run_mcp_tests()
        if failed > 0:
            print(f"\n⚠️ 发现 {failed} 个问题，需要修复")
            sys.exit(1)
        else:
            print(f"\n🎉 所有测试通过！")
            sys.exit(0)
    except Exception as e:
        print(f"测试执行异常: {str(e)}")
        sys.exit(1)