#!/usr/bin/env python3
"""快速验证修复"""
import os
import sys
import pandas as pd
import tempfile

sys.path.insert(0, 'src')
from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

# 测试 1: 同文件 JOIN
print("\n" + "="*60)
print("测试 P0: 同文件多表 JOIN")
print("="*60)

tmpdir = tempfile.mkdtemp()
file1 = os.path.join(tmpdir, "test1.xlsx")

with pd.ExcelWriter(file1, engine='openpyxl') as w:
    pd.DataFrame({'id': [1, 2], 'name': ['A', 'B']}).to_excel(w, sheet_name='Sheet1', index=False)
    pd.DataFrame({'id': [1, 3], 'value': [100, 300]}).to_excel(w, sheet_name='Sheet2', index=False)

engine = AdvancedSQLQueryEngine()
result = engine.execute_sql_query(
    file1,
    "SELECT a.name, b.value FROM Sheet1 a JOIN Sheet2 b ON a.id = b.id",
    sheet_name='Sheet1'
)

if result['success']:
    print("✅ JOIN 测试通过")
    print(f"   结果: {result['data']}")
else:
    print(f"❌ JOIN 测试失败: {result.get('message')}")

# 测试 2: GROUP_CONCAT 复杂表达式
print("\n" + "="*60)
print("测试 P1: GROUP_CONCAT 复杂表达式")
print("="*60)

file2 = os.path.join(tmpdir, "test2.xlsx")
pd.DataFrame({
    'cls': ['A', 'A', 'B'],
    'val': [1, 2, 3]
}).to_excel(file2, sheet_name='Data', index=False)

result = engine.execute_sql_query(
    file2,
    "SELECT cls, GROUP_CONCAT(CASE WHEN val > 1 THEN 'high' ELSE 'low' END) as tags FROM Data GROUP BY cls"
)

if result['success']:
    print("✅ GROUP_CONCAT 测试通过")
    print(f"   结果: {result['data']}")
else:
    print(f"❌ GROUP_CONCAT 测试失败: {result.get('message')}")

# 清理
import shutil
shutil.rmtree(tmpdir)
print("\n测试完成")
