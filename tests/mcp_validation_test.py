#!/usr/bin/env python3
"""
MCP真实验证测试脚本
测试12项核心功能的实际可用性
"""

import subprocess
import json
import os
from pathlib import Path

def run_mcp_command(command):
    """运行MCP命令并返回结果"""
    try:
        result = subprocess.run(
            command,
            shell=True,
            capture_output=True,
            text=True,
            timeout=30,
            cwd="."
        )
        return {
            "success": result.returncode == 0,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "returncode": result.returncode
        }
    except subprocess.TimeoutExpired:
        return {
            "success": False,
            "stdout": "",
            "stderr": "命令超时",
            "returncode": -1
        }
    except Exception as e:
        return {
            "success": False,
            "stdout": "",
            "stderr": str(e),
            "returncode": -1
        }

def run_core_functions_validation():
    """测试12项核心功能"""
    test_file = "tests/test_data/mcp_verify.xlsx"
    results = {}
    
    # 1. list_sheets
    print("测试 1/12: list_sheets")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{{"name":"excel_list_sheets","arguments":{{"file_path":"{test_file}"}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["list_sheets"] = run_mcp_command(cmd)
    
    # 2. get_range
    print("测试 2/12: get_range")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{{"name":"excel_get_range","arguments":{{"file_path":"{test_file}","sheet_name":"角色配置","range":"A1:C5"}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["get_range"] = run_mcp_command(cmd)
    
    # 3. query WHERE
    print("测试 3/12: query WHERE")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{{"name":"excel_query","arguments":{{"file_path":"{test_file}","query":"SELECT * FROM 角色配置 WHERE 职业=\\"法师\\" AND 等级>5"}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["query_where"] = run_mcp_command(cmd)
    
    # 4. query JOIN
    print("测试 4/12: query JOIN")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{{"name":"excel_query","arguments":{{"file_path":"{test_file}","query":"SELECT r.名称,r.职业,e.名称 AS 装备 FROM 角色配置 r LEFT JOIN 装备配置 e ON r.等级>=e.等级要求 WHERE r.等级>8"}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["query_join"] = run_mcp_command(cmd)
    
    # 5. query GROUP BY
    print("测试 5/12: query GROUP BY")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{{"name":"excel_query","arguments":{{"file_path":"{test_file}","query":"SELECT 职业, AVG(等级) AS 平均等级, SUM(生命值) AS 总生命值 FROM 角色配置 GROUP BY 职业"}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["query_groupby"] = run_mcp_command(cmd)
    
    # 6. query子查询
    print("测试 6/12: query子查询")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":6,"method":"tools/call","params":{{"name":"excel_query","arguments":{{"file_path":"{test_file}","query":"SELECT * FROM (SELECT * FROM 角色配置 WHERE 等级>10) AS 高等级角色 WHERE 职业=\\"战士\\""}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["query_subquery"] = run_mcp_command(cmd)
    
    # 7. query FROM子查询
    print("测试 7/12: query FROM子查询")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":7,"method":"tools/call","params":{{"name":"excel_query","arguments":{{"file_path":"{test_file}","query":"SELECT * FROM (SELECT ID, 名称, 等级 FROM 角色配置 WHERE 等级>=5) AS 高等级 WHERE 等级<=12"}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["query_from_subquery"] = run_mcp_command(cmd)
    
    # 8. get_headers
    print("测试 8/12: get_headers")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":8,"method":"tools/call","params":{{"name":"excel_get_headers","arguments":{{"file_path":"{test_file}","sheet_name":"技能配置"}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["get_headers"] = run_mcp_command(cmd)
    
    # 9. find_last_row
    print("测试 9/12: find_last_row")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":9,"method":"tools/call","params":{{"name":"excel_find_last_row","arguments":{{"file_path":"{test_file}","sheet_name":"角色配置"}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["find_last_row"] = run_mcp_command(cmd)
    
    # 10. batch_insert_rows
    print("测试 10/12: batch_insert_rows")
    insert_data = '[["6", "骑士", "近战", 20, 1500, 200, 160, 90, 110, "2026-03-28 04:35:00"]]'
    cmd = f'echo \'{{"jsonrpc":"2.0","id":10,"method":"tools/call","params":{{"name":"excel_batch_insert_rows","arguments":{{"file_path":"{test_file}","sheet_name":"角色配置","data":{insert_data},"start_row":6}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["batch_insert_rows"] = run_mcp_command(cmd)
    
    # 11. delete_rows
    print("测试 11/12: delete_rows")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":11,"method":"tools/call","params":{{"name":"excel_delete_rows","arguments":{{"file_path":"{test_file}","sheet_name":"角色配置","start_row":6,"count":1}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["delete_rows"] = run_mcp_command(cmd)
    
    # 12. describe_table
    print("测试 12/12: describe_table")
    cmd = f'echo \'{{"jsonrpc":"2.0","id":12,"method":"tools/call","params":{{"name":"excel_describe_table","arguments":{{"file_path":"{test_file}","sheet_name":"角色配置"}}}}}}}}\' | uvx excel-mcp-server-fastmcp'
    results["describe_table"] = run_mcp_command(cmd)
    
    return results

def main():
    print("开始MCP真实验证测试...")
    print("=" * 50)
    
    results = run_core_functions_validation()
    
    print("\n" + "=" * 50)
    print("测试结果汇总:")
    
    passed = 0
    total = len(results)
    
    for test_name, result in results.items():
        status = "✅ PASS" if result["success"] else "❌ FAIL"
        print(f"{test_name:<15} {status}")
        if not result["success"]:
            print(f"  错误: {result['stderr']}")
        passed += 1 if result["success"] else 0
    
    print(f"\n总计: {passed}/{total} 通过 ({passed/total*100:.1f}%)")
    
    # 保存详细结果
    result_file = "mcp_validation_results.json"
    with open(result_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"详细结果已保存到: {result_file}")
    
    return passed == total

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)