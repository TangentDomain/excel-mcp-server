"""诊断 SUM/AVG 空结果的完整执行路径"""

import sys

sys.path.insert(0, "/root/workspace/excel-mcp-server/src")
from openpyxl import Workbook

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

# 创建空表测试文件
wb = Workbook()
ws = wb.active
ws.title = "test"
ws.append(["ID", "Name", "Val"])
wb.save("/tmp/r5_empty_test.xlsx")

fp = "/tmp/r5_empty_test.xlsx"
engine = AdvancedSQLQueryEngine()

print("=" * 60)
print("诊断: 空表上的聚合查询")
print("=" * 60)

# 测试各种聚合
for sql in [
    "SELECT COUNT(*) as cnt FROM test",
    "SELECT SUM(ID) as total FROM test",
    "SELECT AVG(ID) as avg_id FROM test",
    "SELECT MAX(ID) as mx FROM test",
    "SELECT MIN(ID) as mn FROM test",
    "SELECT COUNT(*) as cnt, SUM(ID) as total FROM test",
]:
    print(f"\nSQL: {sql}")
    try:
        parsed = engine._parse_sql(sql)
        print(f"  parsed OK, expressions: {[type(e).__name__ for e in parsed.expressions]}")

        # 手动执行完整流程看中间状态
        worksheets_data = engine._load_workbook(fp)
        sheet_name = engine._extract_table_name(parsed)
        df = worksheets_data[sheet_name]
        print(f"  loaded df: {len(df)} rows, cols={list(df.columns)}")

        result = execute_advanced_sql_query(fp, sql)
        print(f"  final result: {result['data']}, success={result.get('success')}")
        if result.get("message"):
            print(f"  message: {result['message'][:200]}")
    except Exception:
        import traceback

        traceback.print_exc()
