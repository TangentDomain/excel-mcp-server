#!/usr/bin/env python3
"""
边缘案例测试脚本 - 第267轮 - T416-T435
"""

import subprocess
import json
import tempfile
import os
from pathlib import Path

class EdgeCaseTester:
    def __init__(self):
        self.results = []
        self.test_file = None

    def create_test_file(self):
        """创建测试用Excel文件"""
        self.test_file = tempfile.mktemp(suffix=".xlsx")
        # 使用 MCP 创建基础文件
        return self.test_file

    def call_mcp(self, method, params):
        """调用MCP工具"""
        proc = subprocess.Popen(
            ["uvx", "--from", ".", "excel-mcp-server-fastmcp"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        request = json.dumps({
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/call",
            "params": {
                "name": method,
                "arguments": params
            }
        })

        try:
            proc.stdin.write(request + "\n")
            proc.stdin.flush()
            response = proc.stdout.readline()
            proc.terminate()

            result = json.loads(response)
            return result.get("result", {})
        except Exception as e:
            return {"error": str(e)}

    def test(self, num, description, operation, expected_pass=True):
        """执行单个测试"""
        print(f"测试T{num}: {description}")
        try:
            result = operation()
            if expected_pass:
                if result.get("success") or result.get("data"):
                    status = "PASS"
                else:
                    status = "FAIL"
            else:
                if not result.get("success"):
                    status = "PASS"
                else:
                    status = "FAIL"

            self.results.append({
                "num": num,
                "desc": description,
                "status": status,
                "result": str(result)[:200]
            })
            print(f"  结果: {status} - {str(result)[:100]}")
        except Exception as e:
            self.results.append({
                "num": num,
                "desc": description,
                "status": "ERROR",
                "result": str(e)
            })
            print(f"  错误: {e}")

    def run_tests(self):
        """运行所有测试"""
        # T416: 超长公式嵌套
        self.test(416, "超长公式嵌套(10层IF)", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": self.create_test_file(), "query": "SELECT CASE WHEN CASE WHEN CASE WHEN CASE WHEN CASE WHEN CASE WHEN CASE WHEN CASE WHEN CASE WHEN CASE WHEN 1=1 THEN 1 ELSE 0 END=1 THEN 1 ELSE 0 END=1 THEN 1 ELSE 0 END=1 THEN 1 ELSE 0 END=1 THEN 1 ELSE 0 END=1 THEN 1 ELSE 0 END=1 THEN 1 ELSE 0 END=1 THEN 1 ELSE 0 END as x FROM Sheet1"}
        ), expected_pass=False)

        # T417: 空白工作表名
        self.test(417, "空白工作表名", lambda: self.call_mcp(
            "excel_create_workbook",
            {"file_path": "/tmp/test417.xlsx", "sheet_names": ["", "Sheet1"]}
        ), expected_pass=False)

        # T418: 重复工作表名
        self.test(418, "重复工作表名", lambda: self.call_mcp(
            "excel_create_workbook",
            {"file_path": "/tmp/test418.xlsx", "sheet_names": ["Data", "Data"]}
        ), expected_pass=False)

        # T419: 特殊Unicode字符作为列名
        self.test(419, "特殊Unicode字符列名", lambda: self.call_mcp(
            "excel_get_headers",
            {"file_path": "/tmp/test419.xlsx", "sheet_name": "Sheet1"}
        ), expected_pass=True)

        # T420: 超大数字精度
        self.test(420, "超大数字精度(10^15)", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test420.xlsx", "query": "SELECT * FROM Sheet1 WHERE A > 999999999999999"}
        ), expected_pass=True)

        # T421: 负数索引访问
        self.test(421, "负数行列索引", lambda: self.call_mcp(
            "excel_get_range",
            {"file_path": "/tmp/test421.xlsx", "sheet_name": "Sheet1", "range": "A-1:B-5"}
        ), expected_pass=False)

        # T422: 混合引用样式
        self.test(422, "混合引用样式(R1C1+A1)", lambda: self.call_mcp(
            "excel_get_range",
            {"file_path": "/tmp/test422.xlsx", "sheet_name": "Sheet1", "range": "R1C1:A5"}
        ), expected_pass=False)

        # T423: 空字符串查询
        self.test(423, "空字符串WHERE条件", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test423.xlsx", "query": "SELECT * FROM Sheet1 WHERE name = ''"}
        ), expected_pass=True)

        # T424: NULL值比较
        self.test(424, "NULL值比较", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test424.xlsx", "query": "SELECT * FROM Sheet1 WHERE col IS NULL"}
        ), expected_pass=True)

        # T425: 跨工作表引用
        self.test(425, "跨工作表公式引用", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test425.xlsx", "query": "SELECT * FROM Sheet1 WHERE A = (SELECT B FROM Sheet2)"}
        ), expected_pass=True)

        # T426: 日期边界值
        self.test(426, "日期边界值(1900-01-01)", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test426.xlsx", "query": "SELECT * FROM Sheet1 WHERE date < '1900-01-01'"}
        ), expected_pass=True)

        # T427: 时间精度(毫秒)
        self.test(427, "时间精度毫秒级", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test427.xlsx", "query": "SELECT * FROM Sheet1 WHERE time LIKE '%.%'"}
        ), expected_pass=True)

        # T428: 布尔值转换
        self.test(428, "布尔值字符串转换", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test428.xlsx", "query": "SELECT * FROM Sheet1 WHERE active = 'true'"}
        ), expected_pass=True)

        # T429: 科学计数法解析
        self.test(429, "科学计数法数字", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test429.xlsx", "query": "SELECT * FROM Sheet1 WHERE value > 1E-10"}
        ), expected_pass=True)

        # T430: 正则表达式匹配
        self.test(430, "正则表达式模式", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test430.xlsx", "query": "SELECT * FROM Sheet1 WHERE name REGEXP '^A.*'"}
        ), expected_pass=False)

        # T431: 列名SQL关键字冲突
        self.test(431, "列名SQL关键字(SELECT)", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test431.xlsx", "query": "SELECT [SELECT], [FROM], [WHERE] FROM Sheet1"}
        ), expected_pass=False)

        # T432: 中文列名查询
        self.test(432, "中文列名SQL查询", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test432.xlsx", "query": "SELECT 姓名, 年龄 FROM Sheet1"}
        ), expected_pass=False)

        # T433: 列名包含点号
        self.test(433, "列名包含点号", lambda: self.call_mcp(
            "excel_get_headers",
            {"file_path": "/tmp/test433.xlsx", "sheet_name": "Sheet1"}
        ), expected_pass=True)

        # T434: 超长文本值
        self.test(434, "超长文本值(>32000字符)", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test434.xlsx", "query": "SELECT * FROM Sheet1 WHERE LENGTH(text) > 32000"}
        ), expected_pass=False)

        # T435: 非打印字符
        self.test(435, "非打印字符处理", lambda: self.call_mcp(
            "excel_query_sql",
            {"file_path": "/tmp/test435.xlsx", "query": "SELECT * FROM Sheet1 WHERE name LIKE '%\\t%'"}
        ), expected_pass=True)

    def print_summary(self):
        """打印测试总结"""
        pass_count = sum(1 for r in self.results if r["status"] == "PASS")
        fail_count = sum(1 for r in self.results if r["status"] == "FAIL")
        error_count = sum(1 for r in self.results if r["status"] == "ERROR")
        info_count = sum(1 for r in self.results if r["status"] == "INFO")

        print("\n=== 第267轮统计 ===")
        print(f"总计: {len(self.results)}个边缘案例")
        print(f"通过: {pass_count}个")
        print(f"失败: {fail_count}个")
        print(f"错误: {error_count}个")
        print(f"信息: {info_count}个")

        return {
            "total": len(self.results),
            "pass": pass_count,
            "fail": fail_count,
            "error": error_count,
            "info": info_count
        }

if __name__ == "__main__":
    tester = EdgeCaseTester()
    tester.run_tests()
    summary = tester.print_summary()

    # 输出结果为JSON以便解析
    print(json.dumps(summary))
