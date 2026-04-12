#!/usr/bin/env python3
"""
边缘案例测试 T416-T435 - 第267轮
"""

import subprocess
import json
import tempfile
import os
from pathlib import Path

# 创建测试文件
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Sheet1"
ws.append(["Name", "Age", "Score", "Active"])
for i in range(10):
    ws.append([f"User{i}", 20+i, 80+i*2, "true" if i%2==0 else "false"])

ws2 = wb.create_sheet("Sheet2")
ws2.append(["Key", "Value"])
ws2.append(["A", 100])

test_file = "/tmp/edge_test_base.xlsx"
wb.save(test_file)
print(f"测试文件已创建: {test_file}")

# MCP服务器
MCP_SERVER = ["excel-mcp-server-fastmcp"]

def call_mcp(tool_name, params):
    """调用MCP工具"""
    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": tool_name,
            "arguments": params
        }
    }

    try:
        proc = subprocess.Popen(
            MCP_SERVER,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        request_str = json.dumps(request)
        proc.stdin.write(request_str + "\n")
        proc.stdin.flush()

        # 读取响应（多行）
        response_lines = []
        for _ in range(10):
            line = proc.stdout.readline()
            if not line:
                break
            response_lines.append(line)

        proc.terminate()
        proc.wait(timeout=2)

        # 解析响应
        full_response = "".join(response_lines)
        if "result" in full_response:
            result = json.loads(full_response)
            return result.get("result", {})
        return {"raw_response": full_response[:200]}
    except Exception as e:
        return {"error": str(e)}

# 测试用例
tests = [
    (416, "基本SQL查询", "excel_query_sql", {"file_path": test_file, "query": "SELECT * FROM Sheet1 LIMIT 5"}),
    (417, "空文件路径", "excel_get_file_info", {"file_path": ""}),
    (418, "不存在的文件", "excel_list_sheets", {"file_path": "/nonexistent.xlsx"}),
    (419, "列出工作表", "excel_list_sheets", {"file_path": test_file}),
    (420, "获取数据范围", "excel_get_range", {"file_path": test_file, "sheet_name": "Sheet1", "range": "A1:D5"}),
    (421, "获取表头", "excel_get_headers", {"file_path": test_file, "sheet_name": "Sheet1"}),
    (422, "获取文件信息", "excel_get_file_info", {"file_path": test_file}),
    (423, "查找最后行", "excel_find_last_row", {"file_path": test_file, "sheet_name": "Sheet1", "column": "A"}),
    (424, "描述表结构", "excel_describe_table", {"file_path": test_file, "sheet_name": "Sheet1"}),
    (425, "空SQL查询", "excel_query_sql", {"file_path": test_file, "query": ""}),
    (426, "无效SQL语法", "excel_query_sql", {"file_path": test_file, "query": "INVALID"}),
    (427, "WHERE条件查询", "excel_query_sql", {"file_path": test_file, "query": "SELECT * FROM Sheet1 WHERE Age > 25"}),
    (428, "聚合查询COUNT", "excel_query_sql", {"file_path": test_file, "query": "SELECT COUNT(*) FROM Sheet1"}),
    (429, "GROUP BY查询", "excel_query_sql", {"file_path": test_file, "query": "SELECT Active, COUNT(*) FROM Sheet1 GROUP BY Active"}),
    (430, "ORDER BY查询", "excel_query_sql", {"file_path": test_file, "query": "SELECT * FROM Sheet1 ORDER BY Age DESC"}),
    (431, "查询不存在的列", "excel_query_sql", {"file_path": test_file, "query": "SELECT NonExistent FROM Sheet1"}),
    (432, "不存在的工作表", "excel_get_range", {"file_path": test_file, "sheet_name": "NonExistent", "range": "A1:B5"}),
    (433, "无效范围格式", "excel_get_range", {"file_path": test_file, "sheet_name": "Sheet1", "range": "INVALID"}),
    (434, "空范围字符串", "excel_get_range", {"file_path": test_file, "sheet_name": "Sheet1", "range": ""}),
    (435, "超大行号范围", "excel_get_range", {"file_path": test_file, "sheet_name": "Sheet1", "range": "A1:B1048577"}),
]

# 运行测试
print("\n=== 第267轮边缘案例测试 T416-T435 ===\n")

results = {"PASS": 0, "FAIL": 0, "INFO": 0, "ERROR": 0}

for num, desc, tool, params in tests:
    print(f"测试T{num}: {desc} ... ", end="", flush=True)

    try:
        result = call_mcp(tool, params)

        if result.get("success"):
            print("PASS")
            results["PASS"] += 1
        elif "success" in result and not result["success"]:
            print("FAIL")
            results["FAIL"] += 1
        elif "error" in result:
            print(f"ERROR: {result['error'][:50]}")
            results["ERROR"] += 1
        else:
            print("INFO")
            results["INFO"] += 1
    except Exception as e:
        print(f"ERROR: {str(e)[:50]}")
        results["ERROR"] += 1

# 统计
print("\n=== 第267轮统计 ===")
print(f"总计: {len(tests)}个边缘案例(T416-T435)")
print(f"通过: {results['PASS']} 个")
print(f"失败: {results['FAIL']} 个")
print(f"错误: {results['ERROR']} 个")
print(f"信息: {results['INFO']} 个")

# 清理
os.remove(test_file)

# 输出JSON
summary = {
    "total": len(tests),
    "pass": results["PASS"],
    "fail": results["FAIL"],
    "error": results["ERROR"],
    "info": results["INFO"]
}
print(json.dumps(summary))
