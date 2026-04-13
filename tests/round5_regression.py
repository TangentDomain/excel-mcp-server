"""回归测试: 确认 Round 5 修复不破坏已有功能"""

import sys

sys.path.insert(0, "/root/workspace/excel-mcp-server/src")
from openpyxl import Workbook

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
)

# 创建简单测试文件
wb = Workbook()
ws = wb.active
ws.title = "t"
ws.append(["ID", "Name", "Score", "Rate"])
ws.append([1, "Alice", 95.5, 0.123])
ws.append([2, "Bob", 87.25, 0.05])
ws.append([3, "Charlie", 0.0, 0.999])
wb.save("/tmp/r5_regression.xlsx")

fp = "/tmp/r5_regression.xlsx"
passed = 0
failed = 0


def check(name, result, condition, detail=""):
    global passed, failed
    if condition:
        print(f"✅ {name}")
        passed += 1
    else:
        print(f"❌ {name} — {detail}")
        print(f"   data={result.get('data', 'N/A')}")
        failed += 1


print("=" * 60)
print("  回归测试: Round 5 修复验证")
print("=" * 60)

# === Bug 1 回归: 聚合空结果 ===
r = execute_advanced_sql_query(fp, "SELECT COUNT(*) as cnt FROM t")
check(
    "正常 COUNT(*)",
    r,
    r["success"] and r["data"][1][0] == 3,
    f"期望3, 实际{r.get('data')}",
)

r = execute_advanced_sql_query(fp, "SELECT AVG(Score) as avg_s FROM t")
check("正常 AVG()", r, r["success"] and r["data"][1][0] is not None)

r = execute_advanced_sql_query(fp, "SELECT SUM(Score) as total FROM t WHERE ID > 99")
check(
    "SUM 空结果→NULL",
    r,
    r["success"] and len(r["data"]) >= 2 and r["data"][1][0] is None,
    f"期望[None], 实际{r.get('data')}",
)

r = execute_advanced_sql_query(fp, "SELECT COUNT(*) as cnt FROM t WHERE ID > 99")
check(
    "COUNT 空结果→0",
    r,
    r["success"] and len(r["data"]) >= 2 and r["data"][1][0] == 0,
    f"期望[0], 实际{r.get('data')}",
)

# GROUP BY 空结果仍应返回空集(不是默认行)
r = execute_advanced_sql_query(fp, "SELECT Name, COUNT(*) as cnt FROM t WHERE ID > 99 GROUP BY Name")
check(
    "GROUP BY 空结果→空集",
    r,
    r["success"] and len(r["data"]) <= 1,
    f"期望仅header或空, 实际{r.get('data')}",
)

# === Bug 2 回归: 浮点数舍入 ===
r = execute_advanced_sql_query(fp, "SELECT Score, Rate FROM t WHERE ID = 1")
check(
    "常规浮点保留2位",
    r,
    r["success"] and r["data"][1][0] == 95.5,
    f"Score=95.5, 实际{r.get('data')}",
)
check(
    "中等精度浮点",
    r,
    r["success"] and abs(r["data"][1][1] - 0.12) < 0.01,
    f"Rate≈0.12, 实际{r.get('data')}",
)

r = execute_advanced_sql_query(fp, "SELECT Score FROM t WHERE ID = 3")
check(
    "零值不变",
    r,
    r["success"] and r["data"][1][0] == 0,
    f"Score=0, 实际{r.get('data')}",
)

# 极小值测试 (Bug 2 的核心修复场景)
wb2 = Workbook()
ws2 = wb2.active
ws2.title = "tiny"
ws2.append(["ID", "Val"])
ws2.append([1, 0.000001])
ws2.append([2, 8.88e-2])
ws2.append([3, 0.005])
wb2.save("/tmp/r5_tiny.xlsx")

r = execute_advanced_sql_query("/tmp/r5_tiny.xlsx", "SELECT Val FROM tiny WHERE ID = 1")
check(
    "极小值 0.000001 不截断",
    r,
    r["success"] and r["data"][1][0] != 0 and r["data"][1][0] < 0.001,
    f"期望~0.000001, 实际{r.get('data')}",
)

r = execute_advanced_sql_query("/tmp/r5_tiny.xlsx", "SELECT Val FROM tiny WHERE ID = 2")
check(
    "小数值 0.0888 保留",
    r,
    r["success"] and abs(r["data"][1][0] - 0.0888) < 0.001,
    f"期望~0.0888, 实际{r.get('data')}",
)

r = execute_advanced_sql_query("/tmp/r5_tiny.xlsx", "SELECT Val FROM tiny WHERE ID = 3")
check(
    "中 小值 0.005 保留",
    r,
    r["success"] and abs(r["data"][1][0] - 0.005) < 0.0001,
    f"期望~0.005, 实际{r.get('data')}",
)

# 整数不应变浮点
r = execute_advanced_sql_query(fp, "SELECT ID FROM t WHERE ID = 1")
check(
    "整数保持int",
    r,
    r["success"] and isinstance(r["data"][1][0], int) and r["data"][1][0] == 1,
    f"期望int(1), 实际{type(r['data'][1][0]).__name__}({r.get('data')})",
)

print("\n" + "=" * 60)
total = passed + failed
print(f"  回归测试: {passed}/{total} 通过 ({passed / total * 100:.0f}%)" if total > 0 else "  无测试")
if failed:
    print(f"  ⚠️  {failed} 个失败!")
else:
    print("  ✅ 全部通过，无回归问题")
print("=" * 60)
