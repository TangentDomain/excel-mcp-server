import sys
sys.path.insert(0, 'src')
import pandas as pd
from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

# Test the basic data first
print("=== Testing basic data ===")
r = execute_advanced_sql_query('./tests/test_data/comprehensive_test.xlsx', 'SELECT * FROM 员工数据 LIMIT 5')
print(f"Basic query result: {r['data']}")

# Test GROUP BY 
print("\n=== Testing GROUP BY ===")
r = execute_advanced_sql_query('./tests/test_data/comprehensive_test.xlsx', 'SELECT 部门 FROM 员工数据 GROUP BY 部门 ORDER BY 部门')
print(f"GROUP BY 结果: {r['data']}")
print(f"行数: {len(r['data'])}")

# Test with SUM
print("\n=== Testing GROUP BY with SUM ===")
r = execute_advanced_sql_query('./tests/test_data/comprehensive_test.xlsx', 'SELECT 部门, SUM(工资) as total_salary FROM 员工数据 GROUP BY 部门 ORDER BY 部门')
print(f"GROUP BY SUM 结果: {r['data']}")
print(f"行数: {len(r['data'])}")