"""诊断: 空表上 SUM/AVG 返回空结果的根因"""

import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "src"))
import traceback

from openpyxl import Workbook

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
)

# 创建空表测试文件
wb = Workbook()
ws = wb.active
ws.title = "test"
ws.append(["ID", "Name", "Val"])
empty_test_file = str(Path(tempfile.gettempdir()) / "r5_empty_test.xlsx")
wb.save(empty_test_file)

fp = empty_test_file

print("=" * 60)
print("空表聚合查询诊断")
print("=" * 60)

for sql in [
    "SELECT COUNT(*) as cnt FROM test",
    "SELECT SUM(ID) as total FROM test",
    "SELECT AVG(ID) as avg_id FROM test",
]:
    print(f"\n--- SQL: {sql} ---")
    try:
        r = execute_advanced_sql_query(fp, sql)
        print(f"  success={r.get('success')}")
        print(f"  data={r.get('data')}")
        if r.get("message"):
            print(f"  message={r['message'][:300]}")
        if r.get("error_code"):
            print(f"  error_code={r['error_code']}")
    except Exception as e:
        print(f"  EXCEPTION: {e}")
        traceback.print_exc()

# 也测试有数据但 WHERE 过滤掉所有行的情况
print("\n" + "=" * 60)
print("WHERE 过滤后空结果")
print("=" * 60)

wb2 = Workbook()
ws2 = wb2.active
ws2.title = "test2"
ws2.append(["ID", "Val"])
ws2.append([1, 10])
ws2.append([2, 20])
where_empty_file = str(Path(tempfile.gettempdir()) / "r5_where_empty.xlsx")
wb2.save(where_empty_file)

fp2 = where_empty_file

for sql in [
    "SELECT COUNT(*) as cnt FROM test2 WHERE ID > 99",
    "SELECT SUM(Val) as total FROM test2 WHERE ID > 99",
    "SELECT AVG(Val) as avg_v FROM test2 WHERE ID > 99",
]:
    print(f"\n--- SQL: {sql} ---")
    try:
        r = execute_advanced_sql_query(fp2, sql)
        print(f"  success={r.get('success')}")
        print(f"  data={r.get('data')}")
        if r.get("message"):
            print(f"  message={r['message'][:200]}")
    except Exception as e:
        print(f"  EXCEPTION: {e}")
