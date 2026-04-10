#!/usr/bin/env python3
import sys
sys.path.insert(0, '.')
import pandas as pd
import tempfile
import os
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

# 创建测试数据
df1 = pd.DataFrame({
    '技能ID': [1, 2, 3],
    '技能名称': ['火球术', '冰霜箭', '雷电术'],
    '等级限制': [5, 10, 15]
})
df2 = pd.DataFrame({
    '角色ID': [101, 102],
    '角色名称': ['战士', '法师'],
    '等级': [12, 18]
})

with tempfile.NamedTemporaryFile(mode='wb', delete=False, suffix='.xlsx') as f:
    df1.to_excel(f, sheet_name='技能表', index=False)
    df2.to_excel(f, sheet_name='角色表', index=False)
    excel_file = f.name

try:
    engine = AdvancedSQLQueryEngine()
    # 测试非等值连接 <=
    sql = """
    SELECT s.技能名称, s.等级限制, r.角色名称, r.等级
    FROM 技能表 s
    JOIN 角色表 r ON s.等级限制 <= r.等级
    ORDER BY s.等级限制, r.等级
    """
    print("测试 SQL:", sql.strip())
    result = engine.execute_sql_query(excel_file, sql)
    
    if result['success']:
        print("✓ 查询成功！")
        print(f"  返回 {len(result['data'])} 行数据")
        if result['data']:
            for row in result['data']:
                print(f"    {row}")
    else:
        print("✗ 查询失败:", result['message'])
finally:
    os.unlink(excel_file)
