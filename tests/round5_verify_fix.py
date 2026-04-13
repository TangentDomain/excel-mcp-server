"""快速验证 Bug 1 修复: COUNT(*) 空结果返回 [0]"""

import sys

sys.path.insert(0, "/root/workspace/excel-mcp-server/src")
from openpyxl import Workbook

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_delete_query,
    execute_advanced_sql_query,
)

# 创建简单测试文件
wb = Workbook()
ws = wb.active
ws.title = "test"
ws.append(["ID", "Name"])
ws.append([1, "aaa"])
ws.append([2, "bbb"])
wb.save("/tmp/r5_fix_test.xlsx")

fp = "/tmp/r5_fix_test.xlsx"

print("=" * 60)
print("验证1: COUNT(*) 正常情况")
r = execute_advanced_sql_query(fp, "SELECT COUNT(*) as cnt FROM test")
print(f"  result: {r['data']}")

print("\n验证2: 删除后 COUNT(*) 空结果")
execute_advanced_delete_query(fp, "DELETE FROM test WHERE ID <= 2")
r = execute_advanced_sql_query(fp, "SELECT COUNT(*) as cnt FROM test")
print(f"  result: {r['data']}")
print(f"  success: {r['success']}")
if r["data"] and len(r["data"]) >= 2:
    print(f"  ✅ cnt value: {r['data'][1][0]}")
else:
    print("  ❌ 只有header，没有数据行!")

print("\n验证3: SUM 空结果 (应返回 NULL)")
r = execute_advanced_sql_query(fp, "SELECT SUM(ID) as total FROM test")
print(f"  result: {r['data']}")
if r["data"] and len(r["data"]) >= 2:
    print(f"  ✅ SUM 空值: {r['data'][1][0]}")
else:
    print("  ❌ 缺少默认行")

print("\n验证4: AVG 空结果 (应返回 NULL)")
r = execute_advanced_sql_query(fp, "SELECT AVG(ID) as avg_id FROM test")
print(f"  result: {r['data']}")
if r["data"] and len(r["data"]) >= 2:
    print(f"  ✅ AVG 空值: {r['data'][1][0]}")

print("\n验证5: 多聚合函数空结果")
r = execute_advanced_sql_query(fp, "SELECT COUNT(*) as cnt, SUM(ID) as total, AVG(ID) as avg_id FROM test")
print(f"  result: {r['data']}")
if r["data"] and len(r["data"]) >= 2:
    print(f"  ✅ 多聚合: cnt={r['data'][1][0]}, total={r['data'][1][1]}, avg={r['data'][1][2]}")

print("\n验证6: GROUP BY + 无匹配 (应返回空，不是默认行)")
r = execute_advanced_sql_query(fp, "SELECT Name, COUNT(*) as cnt FROM test GROUP BY Name")
print(f"  result: {r['data']}")
if len(r["data"]) <= 1:
    print("  ✅ GROUP BY空结果正确返回空集(只有header)")

print("\n" + "=" * 60)
