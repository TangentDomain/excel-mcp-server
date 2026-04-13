#!/usr/bin/env python3
"""诊断 Round 8 假失败用例"""
import sys
sys.path.insert(0, '/root/workspace/excel-mcp-server/src')
from openpyxl import Workbook
import tempfile, os
from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

wb = Workbook()
ws = wb.active
ws.title = "装备"
ws.append(["ID", "Name", "BaseAtk", "AtkBonus", "Price", "Rarity", "Category"])
ws.append([1, "Item-1", 100, 10.5, 500.0, "Common", "Weapon"])
ws.append([999, "NULL测试装", None, None, None, None, None])
ws.append([1000, "A" * 500, 1, 1.0, 1.0, "Common", "Weapon"])

tmpfile = tempfile.mktemp(suffix='.xlsx')
wb.save(tmpfile)

tests = [
    ("A1_diag", "WITH 高价值装备 AS (SELECT ID, Name, Price FROM 装备 WHERE Price > 100) SELECT * FROM 高价值装备 ORDER BY Price DESC LIMIT 5"),
    ("D2_diag", "SELECT Price FROM 装备 WHERE ID = 1"),
    ("F2_diag", "SELECT ID, LENGTH(Name) as name_len FROM 装备 WHERE ID = 1000"),
    ("G2_diag", "SELECT * FROM 装备 WHERE ID = 2001"),
]

for name, sql in tests:
    r = execute_advanced_sql_query(tmpfile, sql)
    print(f"=== {name} ===")
    print(f"  success: {r.get('success')}")
    dt = r.get('data')
    print(f"  data type: {type(dt).__name__}")
    if isinstance(dt, list):
        print(f"  data len: {len(dt)}")
        if dt:
            print(f"  data[0] type: {type(dt[0]).__name__}")
            print(f"  data[0]: {str(dt[0])[:200]}")
    elif hasattr(dt, 'shape'):
        print(f"  shape: {dt.shape}")
    else:
        print(f"  data value: {str(dt)[:200]}")
    print(f"  message: {r.get('message', '')[:100]}")
    # Test the verify logic
    if name == "A1_diag":
        v = r['success'] and len(r.get('data', [])) <= 5
        print(f"  VERIFY A1 logic: success={r.get('success')}, len(data)={len(r.get('data', [])) if isinstance(r.get('data'), list) else 'N/A'}, result={v}")
    elif name == "D2_diag":
        v = r['success'] and len(r.get('data', [])) == 1
        print(f"  VERIFY D2 logic: success={r.get('success')}, len(data)={len(r.get('data', [])) if isinstance(r.get('data'), list) else 'N/A'}, result={v}")
    print()

os.unlink(tmpfile)
