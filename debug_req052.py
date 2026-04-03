#!/usr/bin/env python3
"""调试 REQ-052 GROUP BY 聚合 bug"""
import sys
import pandas as pd

sys.path.insert(0, 'src')

# 1. 加载数据，检查原始数据
print("=== 1. 原始数据检查 ===")
try:
    df_original = pd.read_excel('/tmp/MapEvent.xlsx')
    print(f"原始数据行数: {len(df_original)}")
    print(f"原始数据列: {list(df_original.columns)}")
    if '显示路径ID' in df_original.columns:
        print(f"显示路径ID唯一值: {df_original['显示路径ID'].unique()}")
        print(f"显示路径ID范围: {df_original['显示路径ID'].min()} - {df_original['显示路径ID'].max()}")
    if '显示位置ID' in df_original.columns:
        print(f"显示位置ID唯一值: {df_original['显示位置ID'].unique()[:10]}...")
        print(f"显示位置ID范围: {df_original['显示位置ID'].min()} - {df_original['显示位置ID'].max()}")
except Exception as e:
    print(f"加载数据失败: {e}")
    sys.exit(1)

# 2. 执行 WHERE 过滤
print("\n=== 2. WHERE 过滤后数据 ===")
try:
    df_filtered = df_original[
        (df_original['显示路径ID'].isin([1, 2])) &
        (df_original['显示位置ID'] < 100)
    ]
    print(f"WHERE 后数据行数: {len(df_filtered)}")
    print(f"WHERE 后显示路径ID唯一值: {df_filtered['显示路径ID'].unique()}")
    print(f"WHERE 后显示位置ID唯一值: {df_filtered['显示位置ID'].unique()}")
except Exception as e:
    print(f"WHERE 过滤失败: {e}")
    sys.exit(1)

# 3. 使用 pandas groupby（预期正确结果）
print("\n=== 3. Pandas groupby 结果（预期正确）===")
try:
    grouped = df_filtered.groupby(['显示路径ID', '显示位置ID'], observed=True).size().reset_index(name='cnt')
    print(f"分组后行数: {len(grouped)}")
    print(f"显示路径ID范围: {grouped['显示路径ID'].min()} - {grouped['显示路径ID'].max()}")
    print(f"显示位置ID范围: {grouped['显示位置ID'].min()} - {grouped['显示位置ID'].max()}")
    print(f"前5行:\n{grouped.head()}")
except Exception as e:
    print(f"Pandas groupby 失败: {e}")
    sys.exit(1)

# 4. 使用 excel_mcp_server 执行（可能错误结果）
print("\n=== 4. Excel MCP Server 执行结果（可能有 bug）===")
try:
    from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
    result = execute_advanced_sql_query(
        '/tmp/MapEvent.xlsx',
        'SELECT 显示路径ID, 显示位置ID, COUNT(*) as cnt FROM MapEvent WHERE 显示路径ID IN (1, 2) AND 显示位置ID < 100 GROUP BY 显示路径ID, 显示位置ID'
    )
    data = result['data']
    print(f"返回数据行数（含表头）: {len(data)}")
    if len(data) > 1:
        df_result = pd.DataFrame(data[1:], columns=data[0])
        # 排除 TOTAL 行（用于检查异常值）
        df_data = df_result[df_result['显示路径ID'] != 'TOTAL']
        print(f"分组后行数: {len(df_data)}")
        print(f"显示路径ID范围: {df_data['显示路径ID'].min()} - {df_data['显示路径ID'].max()}")
        print(f"显示位置ID范围: {df_data['显示位置ID'].min()} - {df_data['显示位置ID'].max()}")
        print(f"前5行:\n{df_data.head()}")
        # 检查异常行
        bad_rows = df_data[(~df_data['显示路径ID'].isin([1, 2])) | (df_data['显示位置ID'] >= 100)]
        if len(bad_rows) > 0:
            print(f"异常行数: {len(bad_rows)}")
            print(f"异常行:\n{bad_rows}")
        else:
            print("无异常行")
except Exception as e:
    print(f"Excel MCP Server 执行失败: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# 5. 检查 SQL 解析
print("\n=== 5. SQL 解析检查 ===")
try:
    from sqlglot import parse
    from sqlglot import exp
    sql = 'SELECT 显示路径ID, 显示位置ID, COUNT(*) as cnt FROM MapEvent WHERE 显示路径ID IN (1, 2) AND 显示位置ID < 100 GROUP BY 显示路径ID, 显示位置ID'
    parsed = parse(sql)[0]

    print("SELECT 表达式:")
    for i, select_expr in enumerate(parsed.expressions):
        print(f"  [{i}] 类型: {type(select_expr).__name__}", end="")
        if isinstance(select_expr, exp.Column):
            print(f", name: {select_expr.name}", end="")
            if hasattr(select_expr, 'table') and select_expr.table:
                print(f", table: {select_expr.table}", end="")
        print()

    print("\nGROUP BY 表达式:")
    group_clause = parsed.args.get('group')
    if group_clause:
        for i, group_expr in enumerate(group_clause.expressions):
            print(f"  [{i}] 类型: {type(group_expr).__name__}", end="")
            if isinstance(group_expr, exp.Column):
                print(f", name: {group_expr.name}", end="")
                if hasattr(group_expr, 'table') and group_expr.table:
                    print(f", table: {group_expr.table}", end="")
            print()
except Exception as e:
    print(f"SQL 解析失败: {e}")
    import traceback
    traceback.print_exc()
