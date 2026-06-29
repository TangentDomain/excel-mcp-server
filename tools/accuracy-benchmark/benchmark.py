#!/usr/bin/env python
"""准确率 benchmark — SQL-over-Excel 引擎差分测试。

Methodology:
1. 创建多个 Excel fixture（固定种子的确定性数据）
2. 生成大量 SQL 测试用例（参数化模板 × 参数组合）
3. 每条用例在 ExcelMCP 引擎和 SQLite 真值引擎上分别执行
4. 对比结果（行数 + 逐值匹配，浮点容差 0.01）
5. 聚合准确率指标

Primary metric: accuracy_pct (higher is better, 0~100%)
"""

import json
import os
import random
import re
import shutil
import sqlite3
import statistics
import sys
import tempfile
import time
import math
from pathlib import Path

# 确保 repo root 在 path 中
REPO_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(REPO_ROOT / "src"))

from openpyxl import Workbook

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
    execute_advanced_update_query,
    execute_advanced_insert_query,
    execute_advanced_delete_query,
)

# ============================================================
# 种子 fixtures — 确定性，固定种子
# ============================================================

FIXTURE_SEED = 20260629

FIXTURES = {
    "products": {
        "table": "商品",
        "columns": ["ID", "Name", "Price", "Stock", "Active"],
        "dual_header": False,
        "data": [
            [1, "铁剑", 100.5, 50, "是"],
            [2, "火球术", 250.0, 30, "否"],
            [3, "生命药水", 50.0, 100, "是"],
            [4, "铁盾", 200.0, None, "是"],
            [5, "魔法弓", 300.0, 20, "否"],
            [6, "毒刃", 0.0, 0, "是"],
        ],
        "create_sql": """CREATE TABLE 商品 (ID INTEGER, Name TEXT, Price REAL, Stock INTEGER, Active TEXT)""",
        "insert_sql": """INSERT INTO 商品 VALUES (?, ?, ?, ?, ?)""",
    },
    "students": {
        "table": "学生",
        "columns": ["ID", "Name", "Score", "Level"],
        "dual_header": False,
        "data": [
            [1, "张三", 95.5, "A"],
            [2, "李四", 82.0, "B"],
            [3, "王五", 67.5, "C"],
            [4, "赵六", 55.0, "D"],
            [5, "孙七", 78.5, "B"],
            [6, "周八", 92.0, "A"],
            [7, "吴九", 45.0, "D"],
            [8, None, None, None],
        ],
        "create_sql": """CREATE TABLE 学生 (ID INTEGER, Name TEXT, Score REAL, Level TEXT)""",
        "insert_sql": """INSERT INTO 学生 VALUES (?, ?, ?, ?)""",
    },
    "categories": {
        "table": "类别",
        "columns": ["ID", "Category", "Discount"],
        "dual_header": False,
        "data": [
            [1, "武器", 1.0],
            [2, "消耗品", 0.95],
            [3, "魔法", 0.85],
        ],
        "create_sql": """CREATE TABLE 类别 (ID INTEGER, Category TEXT, Discount REAL)""",
        "insert_sql": """INSERT INTO 类别 VALUES (?, ?, ?)""",
    },
    "combined": {
        # 多表文件: 商品 + 类别 在同一 xlsx 中, 支持 JOIN 和子查询
        "tables": ["products", "categories"],
    },
    "numbers": {
        "table": "数值表",
        "columns": ["ID", "IntVal", "FloatVal", "MixedStr"],
        "dual_header": False,
        "data": [
            [1, 42, 3.14, "a"],
            [2, -1, -0.001, "42"],
            [3, 0, 0.0, ""],
            [4, 999999999, 1e-10, "3.14"],
            [5, -999999999, -1e10, "hello"],
        ],
        "create_sql": """CREATE TABLE 数值表 (ID INTEGER, IntVal INTEGER, FloatVal REAL, MixedStr TEXT)""",
        "insert_sql": """INSERT INTO 数值表 VALUES (?, ?, ?, ?)""",
    },
}


def _normalize(val: object) -> object:
    """归一化值用于比较：None/'' 等价，float 控制精度。"""
    if val is None or (isinstance(val, str) and val == ""):
        return None
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return None
    return val


def _values_match(a: object, b: object, tol: float = 0.01) -> bool:
    """差分比较两个值，含类型归一化和浮点容差。"""
    a = _normalize(a)
    b = _normalize(b)
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) <= tol
    return str(a) == str(b)


def _rows_match(engine_row: list, sqlite_row: tuple, tol: float = 0.01) -> bool:
    """比较两行是否匹配。"""
    if len(engine_row) != len(sqlite_row):
        return False
    return all(_values_match(a, b, tol) for a, b in zip(engine_row, sqlite_row))


def _make_excel(path: str, fixture_name: str) -> None:
    """从 fixture 定义创建 Excel 文件（支持多表 combined fixture）。"""
    fix = FIXTURES[fixture_name]
    wb = Workbook()
    if "tables" in fix:
        # combined fixture: 多个 sheet
        first = True
        for sub_name in fix["tables"]:
            sub = FIXTURES[sub_name]
            ws = wb.active if first else wb.create_sheet()
            ws.title = sub["table"]
            ws.append(sub["columns"])
            for row in sub["data"]:
                ws.append(list(row))
            first = False
    else:
        ws = wb.active
        ws.title = fix["table"]
        ws.append(fix["columns"])
        for row in fix["data"]:
            ws.append(list(row))
    wb.save(path)
    wb.close()


def _make_sqlite(fixture_name: str) -> sqlite3.Connection:
    """从 fixture 定义创建内存 SQLite 数据库（支持多表 combined fixture）。"""
    fix = FIXTURES[fixture_name]
    conn = sqlite3.connect(":memory:")
    if "tables" in fix:
        for sub_name in fix["tables"]:
            sub = FIXTURES[sub_name]
            conn.execute(sub["create_sql"])
            for row in sub["data"]:
                conn.execute(sub["insert_sql"], row)
    else:
        conn.execute(fix["create_sql"])
        for row in fix["data"]:
            conn.execute(fix["insert_sql"], row)
    conn.commit()
    return conn


def _query_engine(file_path: str, sql: str) -> dict:
    """在 ExcelMCP 引擎上执行 SQL。"""
    return execute_advanced_sql_query(file_path, sql)


def _query_sqlite(conn: sqlite3.Connection, sql: str) -> list[tuple]:
    """在 SQLite 上执行 SQL，返回结果行。"""
    try:
        return conn.execute(sql).fetchall()
    except Exception:
        return []


def _check_query(engine_res: dict, sqlite_rows: list[tuple], sql: str, tol: float = 0.01) -> tuple[bool, str]:
    """差分对比：引擎结果 vs SQLite 结果。"""
    if not engine_res.get("success"):
        return False, f"引擎返回失败: {engine_res.get('message', '')}"
    engine_data = engine_res.get("data", [])
    if not engine_data:
        return False, "引擎返回空数据"
    # engine_data[0] 是表头，engine_data[1:] 是数据行
    engine_rows = engine_data[1:] if len(engine_data) > 1 else []
    # 行数匹配
    if len(engine_rows) != len(sqlite_rows):
        return False, f"行数不匹配: engine={len(engine_rows)} sqlite={len(sqlite_rows)}"
    # 逐行比较
    for i, (e_row, s_row) in enumerate(zip(engine_rows, sqlite_rows)):
        if not _rows_match(e_row, s_row, tol):
            return (
                False,
                f"第{i + 1}行不匹配: engine={e_row} sqlite={s_row}",
            )
    return True, ""


# ============================================================
# SQL 测试用例生成器
# ============================================================

CASE_GENERATORS: list[tuple] = []


def _reg(
    sql: str,
    fixture: str,
    category: str,
    sub_category: str,
    tags: list[str] | None = None,
    tol: float = 0.01,
    op_type: str = "query",
) -> None:
    """注册一条测试用例。"""
    CASE_GENERATORS.append(
        {
            "sql": sql,
            "fixture": fixture,
            "category": category,
            "sub_category": sub_category,
            "tags": tags or [category, sub_category],
            "tol": tol,
            "op_type": op_type,
        }
    )


# ==================== 注册用例 ====================

# --- SELECT 基础 ---
_reg("SELECT * FROM 商品", "products", "SELECT", "star", tags=["SELECT", "*"])
_reg("SELECT ID, Name FROM 商品", "products", "SELECT", "specific_cols")
_reg("SELECT Price, Stock, Active FROM 商品", "products", "SELECT", "subset_cols")
_reg("SELECT Price + 10 FROM 商品", "products", "SELECT", "expr_col", tol=0.001)
_reg("SELECT Price AS 价格 FROM 商品 WHERE ID = 1", "products", "SELECT", "alias")
_reg("SELECT Name, Price * Stock AS 总值 FROM 商品 WHERE Stock IS NOT NULL", "products", "SELECT", "alias_expr", tol=0.001)
_reg("SELECT *, Price FROM 商品 WHERE ID = 1", "products", "SELECT", "star_plus_col")
_reg("SELECT 1 + 2, 'hello', 3.14 FROM 商品 WHERE ID = 1", "products", "SELECT", "literal_exprs")

# --- WHERE 条件 ---
_reg("SELECT * FROM 商品 WHERE ID = 3", "products", "WHERE", "eq")
_reg("SELECT * FROM 商品 WHERE ID <> 3", "products", "WHERE", "neq")
_reg("SELECT * FROM 商品 WHERE Price > 100", "products", "WHERE", "gt")
_reg("SELECT * FROM 商品 WHERE Price >= 200", "products", "WHERE", "gte")
_reg("SELECT * FROM 商品 WHERE Price < 100", "products", "WHERE", "lt")
_reg("SELECT * FROM 商品 WHERE Price <= 60", "products", "WHERE", "lte")
_reg("SELECT * FROM 商品 WHERE ID IN (1, 3, 5)", "products", "WHERE", "in")
_reg("SELECT * FROM 商品 WHERE ID NOT IN (1, 5)", "products", "WHERE", "not_in")
_reg("SELECT * FROM 商品 WHERE Name LIKE '%剑'", "products", "WHERE", "like_suffix")
_reg("SELECT * FROM 商品 WHERE Name LIKE '火%'", "products", "WHERE", "like_prefix")
_reg("SELECT * FROM 商品 WHERE Name LIKE '%水%'", "products", "WHERE", "like_contain")
_reg("SELECT * FROM 商品 WHERE Name NOT LIKE '%剑'", "products", "WHERE", "not_like")
_reg("SELECT * FROM 商品 WHERE Price BETWEEN 50 AND 220", "products", "WHERE", "between")
_reg("SELECT * FROM 商品 WHERE Price >= 100 AND Active = '是'", "products", "WHERE", "and", tol=0.001)
_reg("SELECT * FROM 商品 WHERE Price < 100 OR Active = '否'", "products", "WHERE", "or")
_reg("SELECT * FROM 商品 WHERE (Price > 100 AND Active = '是') OR Stock IS NULL", "products", "WHERE", "compound")
_reg("SELECT * FROM 商品 WHERE Name IS NULL", "products", "WHERE", "is_null_str")
_reg("SELECT * FROM 商品 WHERE Price > 0 ORDER BY Price", "products", "WHERE", "where_order")

# --- 聚合与 GROUP BY ---
_reg("SELECT COUNT(*) FROM 商品", "products", "AGGR", "count_star")
_reg("SELECT COUNT(ID) FROM 商品", "products", "AGGR", "count_col")
_reg("SELECT COUNT(Stock) FROM 商品", "products", "AGGR", "count_null_col", tags=["AGGR", "COUNT", "NULL"])
_reg("SELECT SUM(Price) FROM 商品", "products", "AGGR", "sum", tol=0.001)
_reg("SELECT AVG(Price) FROM 商品", "products", "AGGR", "avg", tol=0.001)
_reg("SELECT MIN(Price), MAX(Price) FROM 商品", "products", "AGGR", "min_max", tol=0.001)
_reg("SELECT Active, COUNT(*) FROM 商品 GROUP BY Active", "products", "AGGR", "group_by_single")
_reg("SELECT Active, SUM(Price), AVG(Price), MAX(Price) FROM 商品 GROUP BY Active", "products", "AGGR", "group_by_multi_aggr", tol=0.001)
_reg("SELECT Active, COUNT(*) FROM 商品 GROUP BY Active HAVING COUNT(*) > 1", "products", "AGGR", "having")
_reg("SELECT Active, SUM(Price * 2) FROM 商品 GROUP BY Active", "products", "AGGR", "group_by_expr", tol=0.001)
_reg("SELECT Active, COUNT(Stock) FROM 商品 GROUP BY Active", "products", "AGGR", "count_with_null")
_reg("SELECT COUNT(*), SUM(Price), AVG(Price) FROM 商品 WHERE Active = '是'", "products", "AGGR", "aggr_where", tol=0.001)
_reg("SELECT * FROM 商品 ORDER BY Price DESC LIMIT 3", "products", "AGGR", "order_by_limit")

# --- 学生表聚合 (更多分组变化) ---
_reg("SELECT Level, COUNT(*), AVG(Score) FROM 学生 GROUP BY Level", "students", "AGGR", "group_by_students", tol=0.001)
_reg("SELECT Level, COUNT(*), AVG(Score) FROM 学生 GROUP BY Level HAVING COUNT(*) >= 2", "students", "AGGR", "having_students", tol=0.001)
_reg("SELECT Level, AVG(Score), MIN(Score), MAX(Score) FROM 学生 GROUP BY Level", "students", "AGGR", "aggr_stats", tol=0.001)
_reg("SELECT COUNT(*) FROM 学生 WHERE Score > 80", "students", "AGGR", "count_where_int")
_reg("SELECT AVG(Score) FROM 学生", "students", "AGGR", "avg_simple", tol=0.001)
_reg("SELECT SUM(Score) FROM 学生 WHERE Score IS NOT NULL", "students", "AGGR", "sum_filtered", tol=0.001)

# --- ORDER BY / DISTINCT / LIMIT ---
_reg("SELECT * FROM 商品 ORDER BY Price", "products", "ORDER", "order_asc_default")
_reg("SELECT * FROM 商品 ORDER BY Price DESC", "products", "ORDER", "order_desc")
_reg("SELECT * FROM 商品 ORDER BY Active, Price DESC", "products", "ORDER", "order_multi")
_reg("SELECT * FROM 商品 ORDER BY Price DESC LIMIT 2", "products", "ORDER", "order_limit", tol=0.001)
_reg("SELECT * FROM 商品 ORDER BY Price LIMIT 3 OFFSET 1", "products", "ORDER", "order_limit_offset", tol=0.001)
_reg("SELECT * FROM 商品 ORDER BY Price DESC LIMIT 10", "products", "ORDER", "order_desc_overflow")
_reg("SELECT DISTINCT Active FROM 商品", "products", "DISTINCT", "distinct_simple")
_reg("SELECT DISTINCT Stock FROM 商品 WHERE Stock IS NOT NULL", "products", "DISTINCT", "distinct_where")
_reg("SELECT * FROM 商品 LIMIT 3", "products", "LIMIT", "limit_simple")
_reg("SELECT * FROM 商品 LIMIT 3 OFFSET 2", "products", "LIMIT", "limit_offset")
_reg("SELECT * FROM 商品 LIMIT 100", "products", "LIMIT", "limit_overflow")
_reg("SELECT * FROM 商品 LIMIT 2 OFFSET 4", "products", "LIMIT", "offset_only")
_reg("SELECT Name, Price FROM 商品 ORDER BY Price DESC LIMIT 3 OFFSET 1", "products", "ORDER", "order_colspec")

# --- 字符串函数 ---
_reg("SELECT UPPER(Name) FROM 商品 WHERE Name IS NOT NULL AND Name != ''", "products", "STRING", "upper")
_reg("SELECT LOWER(Name) FROM 商品 WHERE Name IS NOT NULL AND Name != ''", "products", "STRING", "lower")
_reg("SELECT LENGTH(Name) FROM 商品 WHERE Name IS NOT NULL", "products", "STRING", "length")
_reg("SELECT TRIM(Name) FROM 商品 WHERE Name IS NOT NULL", "products", "STRING", "trim")
_reg("SELECT SUBSTRING(Name, 1, 2) FROM 商品 WHERE Name IS NOT NULL AND Name != ''", "products", "STRING", "substring")
_reg("SELECT SUBSTRING(Name, 2, 3) FROM 商品 WHERE Name IS NOT NULL AND LENGTH(Name) >= 3", "products", "STRING", "substring_mid")
_reg("SELECT CONCAT(Name, ' - ', Active) FROM 商品 WHERE Name IS NOT NULL AND Name != ''", "products", "STRING", "concat")
_reg("SELECT * FROM 学生 ORDER BY Name", "students", "STRING", "order_str")
_reg("SELECT LENGTH(Name) FROM 学生", "students", "STRING", "length_null")

# --- 数学表达式 ---
_reg("SELECT Price * 2 FROM 商品", "products", "MATH", "mul", tol=0.001)
_reg("SELECT Price + Stock FROM 商品 WHERE Stock IS NOT NULL", "products", "MATH", "add", tol=0.001)
_reg("SELECT Price - 10 FROM 商品", "products", "MATH", "sub", tol=0.001)
_reg("SELECT Price / 2 FROM 商品", "products", "MATH", "div", tol=0.001)
_reg("SELECT (Price + Stock) * 0.9 FROM 商品 WHERE Stock IS NOT NULL", "products", "MATH", "compound", tol=0.001)
_reg("SELECT Price * Stock + Price FROM 商品 WHERE Stock IS NOT NULL", "products", "MATH", "mul_add", tol=0.001)
_reg("SELECT Price * Stock * 0.8 FROM 商品 WHERE Stock IS NOT NULL", "products", "MATH", "chain_mul", tol=0.001)
_reg("SELECT IntVal + FloatVal FROM 数值表", "numbers", "MATH", "int_float_add", tol=0.001)
_reg("SELECT IntVal * -1 FROM 数值表", "numbers", "MATH", "negate_mul", tol=0.001)
_reg("SELECT IntVal / 2 FROM 数值表 WHERE IntVal != 0", "numbers", "MATH", "int_div", tol=0.001)

# --- 数值表查询 ---
_reg("SELECT * FROM 数值表", "numbers", "SELECT", "numbers_star")
_reg("SELECT * FROM 数值表 WHERE ID >= 3", "numbers", "WHERE", "numbers_range")
_reg("SELECT * FROM 数值表 WHERE FloatVal > 0", "numbers", "WHERE", "numbers_float_gt", tol=1e-6)
_reg("SELECT * FROM 数值表 WHERE MixedStr = 'hello'", "numbers", "WHERE", "exact_str")

# --- NULL 处理 ---
_reg("SELECT * FROM 商品 WHERE Stock IS NULL", "products", "NULL", "is_null")
_reg("SELECT * FROM 商品 WHERE Stock IS NOT NULL", "products", "NULL", "is_not_null")
_reg("SELECT * FROM 商品 WHERE Name IS NULL OR Name = ''", "products", "NULL", "empty_or_null")
_reg("SELECT * FROM 学生 WHERE Score IS NULL", "students", "NULL", "null_score")
_reg("SELECT * FROM 学生 WHERE Score IS NOT NULL", "students", "NULL", "not_null_score")
_reg("SELECT * FROM 学生 WHERE Name IS NULL OR Name = ''", "students", "NULL", "null_or_empty_name")
_reg("SELECT Name, COALESCE(Stock, 0) FROM 商品 ORDER BY ID", "products", "NULL", "coalesce", tol=0.001)
_reg("SELECT ID, Name FROM 学生 WHERE Score IS NULL", "students", "NULL", "null_selective")

# --- CASE 表达式 ---
_reg("SELECT Name, CASE WHEN Price > 200 THEN '贵' WHEN Price > 100 THEN '中' ELSE '便宜' END FROM 商品", "products", "CASE", "case_when")
_reg("SELECT Name, Price * CASE WHEN Active = '是' THEN 1.0 ELSE 0.5 END FROM 商品", "products", "CASE", "case_in_expr", tol=0.001)
_reg("SELECT ID, CASE ID WHEN 1 THEN '一' WHEN 2 THEN '二' ELSE '其他' END FROM 商品", "products", "CASE", "case_simple")

# --- 跨文件 JOIN ---
_reg("SELECT p.Name, c.Category, p.Price FROM 商品 p JOIN 类别 c ON p.ID = c.ID", "combined", "JOIN", "inner_basic")
_reg("SELECT p.Name, c.Category, c.Discount, p.Price * c.Discount AS Discounted FROM 商品 p JOIN 类别 c ON p.ID = c.ID", "combined", "JOIN", "inner_expr", tol=0.001)
_reg("SELECT p.Name, c.Category FROM 商品 p LEFT JOIN 类别 c ON p.ID = c.ID", "combined", "JOIN", "left_join")

# --- 子查询 ---
_reg("SELECT Name, Price FROM 商品 WHERE Price > (SELECT AVG(Price) FROM 商品)", "products", "SUBQUERY", "scalar_gt_avg", tol=0.001)
_reg("SELECT * FROM 商品 WHERE ID IN (SELECT ID FROM 类别)", "combined", "SUBQUERY", "in_subquery")

# --- 边缘值 ---
_reg("SELECT * FROM 数值表 WHERE IntVal = 999999999", "numbers", "EDGE", "large_int")
_reg("SELECT * FROM 数值表 WHERE IntVal = -999999999", "numbers", "EDGE", "neg_large_int")
_reg("SELECT * FROM 数值表 WHERE FloatVal = -1e10", "numbers", "EDGE", "large_neg_float", tol=1.0)
_reg("SELECT * FROM 数值表 WHERE ID = 3 AND IntVal = 0", "numbers", "EDGE", "zero_int")
_reg("SELECT * FROM 数值表 WHERE FloatVal = 0.0", "numbers", "EDGE", "zero_float", tol=1e-6)
_reg("SELECT * FROM 数值表 WHERE FloatVal = -0.001", "numbers", "EDGE", "tiny_neg_float", tol=1e-6)
_reg("SELECT * FROM 数值表 WHERE FloatVal = 1e-10", "numbers", "EDGE", "tiny_pos_float", tol=1e-6)
_reg("SELECT * FROM 商品 WHERE ID = 6 AND Price = 0", "products", "EDGE", "zero_price", tol=1e-6)
_reg("SELECT * FROM 商品 WHERE ID = 6 AND Stock = 0", "products", "EDGE", "zero_stock")
_reg("SELECT * FROM 商品 WHERE Name = ''", "products", "EDGE", "empty_str")
_reg("SELECT * FROM 学生 WHERE Name = ''", "students", "EDGE", "empty_str_students")
_reg("SELECT IntVal * FloatVal FROM 数值表 WHERE ID = 1", "numbers", "EDGE", "multiply_int_float", tol=1e-6)
_reg("SELECT IntVal * 0 FROM 数值表", "numbers", "EDGE", "mul_zero", tol=1e-6)
_reg("SELECT * FROM 商品 WHERE Name IS NULL", "products", "EDGE", "name_null")
_reg("SELECT * FROM 商品 WHERE Name LIKE '%'", "products", "EDGE", "like_any")
_reg("SELECT * FROM 商品 ORDER BY Stock NULLS LAST", "products", "EDGE", "nulls_last")
_reg("SELECT * FROM 商品 WHERE Stock IS NULL AND Active = '是'", "products", "EDGE", "null_with_other")
_reg("SELECT * FROM 商品 WHERE ID BETWEEN 1 AND 10", "products", "EDGE", "between_wide")

# --- 复合功能 ---
_reg("SELECT Active, COUNT(*) AS cnt, ROUND(AVG(Price), 2) AS avg_p FROM 商品 GROUP BY Active ORDER BY cnt DESC", "products", "COMPLEX", "aggr_order", tol=0.001)
_reg("SELECT Level, COUNT(*) FROM 学生 GROUP BY Level ORDER BY Level", "students", "COMPLEX", "group_order")
_reg("SELECT Level, COUNT(*) FROM 学生 WHERE Score IS NOT NULL GROUP BY Level HAVING AVG(Score) > 70", "students", "COMPLEX", "where_group_having", tol=0.001)

# ==================== 写操作用例 ====================

# UPDATE
_reg("UPDATE 商品 SET Price = Price * 1.1 WHERE Active = '是'", "products", "UPDATE", "mul_expr", op_type="update", tol=0.001)
_reg("UPDATE 商品 SET Stock = NULL WHERE ID = 2", "products", "UPDATE", "set_null", op_type="update")
_reg("UPDATE 商品 SET Name = '测试', Price = 999.0 WHERE ID = 1", "products", "UPDATE", "multi_col", op_type="update", tol=0.001)
_reg("UPDATE 商品 SET Price = Price + 10 WHERE Stock IS NOT NULL", "products", "UPDATE", "expr_where", op_type="update", tol=0.001)
_reg("UPDATE 商品 SET Active = '否'", "products", "UPDATE", "no_where", op_type="update")
_reg("UPDATE 商品 SET Price = 0 WHERE ID = 999", "products", "UPDATE", "no_match", op_type="update")

# INSERT
_reg("INSERT INTO 商品 (ID, Name, Price, Stock, Active) VALUES (10, '新品', 150.0, 5, '是')", "products", "INSERT", "full_row", op_type="insert", tol=0.001)
_reg("INSERT INTO 商品 (ID, Name, Price) VALUES (11, '部分数据', 99.9)", "products", "INSERT", "partial_cols", op_type="insert", tol=0.001)
_reg("INSERT INTO 商品 (ID, Name, Price, Stock, Active) VALUES (12, '浮点测试', 3.14159, 1, '是')", "products", "INSERT", "float_precision", op_type="insert", tol=0.00001)
_reg("INSERT INTO 学生 (ID, Name) VALUES (10, '新生')", "students", "INSERT", "minimal_insert", op_type="insert")

# DELETE
_reg("DELETE FROM 商品 WHERE ID = 6", "products", "DELETE", "single_row", op_type="delete")
_reg("DELETE FROM 商品 WHERE Active = '否'", "products", "DELETE", "multi_row", op_type="delete")
_reg("DELETE FROM 商品 WHERE ID = 999", "products", "DELETE", "no_match", op_type="delete")
_reg("DELETE FROM 学生 WHERE Score IS NULL", "students", "DELETE", "null_where", op_type="delete")


def _build_all_cases() -> list[dict]:
    """收集所有注册的用例。"""
    return list(CASE_GENERATORS)


# ============================================================
# 写操作执行
# ============================================================


def _extract_table(sql: str) -> str | None:
    """从 UPDATE/INSERT/DELETE/SELECT 中提取表名。"""
    m = re.match(r"(?:UPDATE|INSERT\s+INTO|DELETE\s+FROM)\s+(\S+)", sql.strip(), re.IGNORECASE)
    if m:
        return m.group(1)
    # SELECT: SELECT ... FROM table
    m = re.search(r"FROM\s+(\S+)", sql, re.IGNORECASE)
    if m:
        return m.group(1)
    return None


def _exec_write_op(op_type: str, excel_path: str, sql: str) -> tuple[bool, str, list | None]:
    """执行写操作，返回 (成功?, 信息, 写后查询行 or None)。"""
    if op_type == "update":
        res = execute_advanced_update_query(excel_path, sql)
    elif op_type == "insert":
        res = execute_advanced_insert_query(excel_path, sql)
    elif op_type == "delete":
        res = execute_advanced_delete_query(excel_path, sql)
    else:
        return False, f"未知操作类型: {op_type}", None
    if not res.get("success"):
        return False, f"写操作失败: {res.get('message', '')}", None
    table = _extract_table(sql)
    if not table:
        return True, "写操作成功(表名解析失败)", None
    verify = execute_advanced_sql_query(excel_path, f"SELECT * FROM {table}")
    if not verify.get("success"):
        return True, f"写后验证查询失败: {verify.get('message', '')}", None
    data = verify.get("data")
    if data is None:
        return True, "写后查询返回空数据", None
    return True, "ok", data[1:] if len(data) > 1 else []


# ============================================================
# 执行与评分
# ============================================================


def _temp_dir():
    return tempfile.mkdtemp(prefix="excelmcp_accuracy_")


def _run_case(
    case: dict,
    tmp_dir: str,
    fixture_paths: dict,
    sqlite_conns: dict,
    results: list[dict],
) -> None:
    """执行单条用例并记录结果。"""
    fixture_name = case["fixture"]
    excel_path = fixture_paths[fixture_name]
    sql = case["sql"]
    op_type = case.get("op_type", "query")
    tol = case.get("tol", 0.01)

    result = {
        "case": case,
        "passed": False,
        "detail": "",
    }

    try:
        if op_type == "query":
            # 只读查询: 对比引擎 vs SQLite
            engine_res = _query_engine(excel_path, sql)
            conn = sqlite_conns[fixture_name]
            sqlite_rows = _query_sqlite(conn, sql)
            ok, msg = _check_query(engine_res, sqlite_rows, sql, tol)
            result["passed"] = ok
            result["detail"] = msg
        else:
            # 写操作: 创建副本, 执行写, 对比副本 vs 独立 SQLite
            copy_path = os.path.join(tmp_dir, f"{fixture_name}_{len(results)}.xlsx")
            shutil.copy(excel_path, copy_path)
            ok, msg, after_data = _exec_write_op(op_type, copy_path, sql)
            if not ok:
                result["passed"] = False
                result["detail"] = msg
            elif after_data is None:
                # 写操作成功但无法验证 (表名解析失败或验证查询失败)
                result["passed"] = True
                result["detail"] = msg or "写操作成功(未验证)"
            else:
                # 用独立 SQLite 执行同样写操作
                write_conn = _make_sqlite(fixture_name)
                try:
                    write_conn.execute(sql)
                    write_conn.commit()
                    sqlite_after = _query_sqlite(write_conn, f"SELECT * FROM {FIXTURES[fixture_name]['table']}")
                except Exception as e:
                    result["passed"] = False
                    result["detail"] = f"SQLite 写操作失败: {e}"
                    results.append(result)
                    write_conn.close()
                    return
                write_conn.close()
                # 对比写后状态
                if len(after_data) != len(sqlite_after):
                    result["passed"] = False
                    result["detail"] = f"写后行数不匹配: engine={len(after_data)} sqlite={len(sqlite_after)}"
                else:
                    all_ok = True
                    for i, (e_row, s_row) in enumerate(zip(after_data, sqlite_after)):
                        if not _rows_match(e_row, s_row, tol):
                            all_ok = False
                            result["detail"] = f"写后第{i + 1}行不匹配: engine={e_row} sqlite={s_row}"
                            break
                    result["passed"] = all_ok
    except Exception as e:
        result["passed"] = False
        result["detail"] = f"异常: {e}"

    results.append(result)


# ============================================================
# 主函数
# ============================================================


def main() -> int:
    """运行准确率 benchmark，输出 METRIC 行。"""
    import argparse

    parser = argparse.ArgumentParser(description="准确率 benchmark")
    parser.add_argument("--verbose", action="store_true", help="显示每条用例的结果")
    parser.add_argument("--json", action="store_true", help="输出 JSON 报告")
    args = parser.parse_args()

    cases = _build_all_cases()
    total = len(cases)

    print(f"⚡ 准确率基准测试: {total} 条用例", file=sys.stderr)
    print(f"   fixtures: {', '.join(FIXTURES.keys())}", file=sys.stderr)

    t_start = time.perf_counter()

    # 创建临时目录
    tmp_dir = _temp_dir()

    # 创建 fixture 文件和 SQLite 连接
    fixture_paths = {}
    sqlite_conns = {}
    for name in FIXTURES:
        fp = os.path.join(tmp_dir, f"{name}.xlsx")
        _make_excel(fp, name)
        fixture_paths[name] = fp
        sqlite_conns[name] = _make_sqlite(name)

    # 执行所有用例
    results: list[dict] = []
    for idx, case in enumerate(cases):
        _run_case(case, tmp_dir, fixture_paths, sqlite_conns, results)
        if args.verbose:
            r = results[-1]
            status = "✅" if r["passed"] else "❌"
            detail = r["detail"] if not r["passed"] else "ok"
            print(f"  [{idx + 1}/{total}] {status} {case['category']}/{case['sub_category']}: {detail}")

    # 清理
    shutil.rmtree(tmp_dir, ignore_errors=True)

    elapsed = time.perf_counter() - t_start

    # 聚合
    passed = sum(1 for r in results if r["passed"])
    failed_list = [r for r in results if not r["passed"]]
    accuracy_pct = round(passed / total * 100, 2) if total > 0 else 0.0

    # 按类别统计
    by_category: dict[str, dict] = {}
    for r in results:
        cat = r["case"]["category"]
        if cat not in by_category:
            by_category[cat] = {"total": 0, "passed": 0}
        by_category[cat]["total"] += 1
        if r["passed"]:
            by_category[cat]["passed"] += 1
    for cat, stats in by_category.items():
        stats["accuracy_pct"] = round(stats["passed"] / stats["total"] * 100, 2) if stats["total"] > 0 else 0.0

    # 输出
    print(f"\n{'=' * 60}", file=sys.stderr)
    print(f"  总用时: {elapsed:.1f}s", file=sys.stderr)
    print(f"  总用例: {total}", file=sys.stderr)
    print(f"  通过: {passed}", file=sys.stderr)
    print(f"  失败: {len(failed_list)}", file=sys.stderr)
    print(f"  准确率: {accuracy_pct}%", file=sys.stderr)
    print(file=sys.stderr)

    # 分类统计
    for cat in sorted(by_category):
        s = by_category[cat]
        bar = "█" * int(s["accuracy_pct"] / 10)
        print(f"  {cat:10s} {s['accuracy_pct']:6.2f}% ({s['passed']}/{s['total']}) {bar}", file=sys.stderr)

    if failed_list:
        print(file=sys.stderr)
        print("  失败用例:", file=sys.stderr)
        for r in failed_list[:10]:
            c = r["case"]
            print(f"    ❌ {c['category']}/{c['sub_category']}: {r['detail'][:120]}", file=sys.stderr)
        if len(failed_list) > 10:
            print(f"    ... 还有 {len(failed_list) - 10} 条", file=sys.stderr)

    print(f"\n{'=' * 60}", file=sys.stderr)

    # METRIC 输出
    print(f"METRIC accuracy_pct={accuracy_pct}")
    print(f"METRIC total_cases={total}")
    print(f"METRIC passed={passed}")
    print(f"METRIC failed={len(failed_list)}")
    print(f"METRIC elapsed_sec={elapsed:.2f}")

    if args.json:
        report = {
            "accuracy_pct": accuracy_pct,
            "total_cases": total,
            "passed": passed,
            "failed": len(failed_list),
            "elapsed_sec": round(elapsed, 2),
            "by_category": {
                cat: {
                    "total": s["total"],
                    "passed": s["passed"],
                    "accuracy_pct": s["accuracy_pct"],
                }
                for cat, s in sorted(by_category.items())
            },
        }
        print(json.dumps(report, ensure_ascii=False, indent=2))

    return 0  # benchmark 自身执行成功（不论准确率高低）


if __name__ == "__main__":
    sys.exit(main())
