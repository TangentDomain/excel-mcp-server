"""L4 限制消除测试：验证已修复的 SQL 引擎限制。

之前 SKILL.md 中记录为「不支持」的限制，现已修复：
- SELECT alias.* FROM (subquery) AS alias
- WHERE 引用 SELECT 别名（非窗口）
- WHERE 引用窗口函数别名（自动重写为子查询）
- run-python 沙箱支持 class 定义和常见异常捕获
"""

from __future__ import annotations

import pytest

from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
from excel_mcp_server_fastmcp.api.script_runner import execute_python_script

from .conftest import get_data_rows, get_headers


class TestQualifiedStar:
    """SELECT t.* 支持测试"""

    def test_qualified_star_from_subquery(self, simple_file):
        """SELECT t.* FROM (SELECT ...) AS t — 子查询 qualified star"""
        result = execute_advanced_sql_query(
            simple_file,
            "SELECT t.* FROM (SELECT ID, Name FROM 数据) AS t",
        )
        assert result["success"], result.get("message", "")
        headers = get_headers(result)
        assert "ID" in headers
        assert "Name" in headers
        assert "Price" not in headers  # 子查询只选了 2 列

    def test_qualified_star_from_table(self, simple_file):
        """SELECT t.* FROM 数据 AS t — 普通表 qualified star"""
        result = execute_advanced_sql_query(
            simple_file,
            "SELECT t.* FROM 数据 AS t LIMIT 3",
        )
        assert result["success"], result.get("message", "")
        rows = get_data_rows(result)
        assert len(rows) == 3

    def test_qualified_star_bare_star_equivalence(self, simple_file):
        """SELECT t.* 和 SELECT * 返回相同列集"""
        sql_qualified = "SELECT t.* FROM (SELECT ID, Name, Price FROM 数据) AS t"
        sql_bare = "SELECT * FROM (SELECT ID, Name, Price FROM 数据) AS t"
        r1 = execute_advanced_sql_query(simple_file, sql_qualified)
        r2 = execute_advanced_sql_query(simple_file, sql_bare)
        assert r1["success"] and r2["success"]
        assert get_headers(r1) == get_headers(r2)
        assert len(get_data_rows(r1)) == len(get_data_rows(r2))


# ============================================================
# Fix 2: WHERE 引用窗口函数别名
# ============================================================


class TestWhereWindowAlias:
    """WHERE 引用窗口函数别名 — 自动重写为子查询"""

    def test_where_rank_filter(self, numbers_file):
        """WHERE rk <= 2 — RANK 别名过滤"""
        result = execute_advanced_sql_query(
            numbers_file,
            "SELECT Category, Value, RANK() OVER (ORDER BY Value DESC) AS rk FROM 数值 WHERE rk <= 2",
        )
        assert result["success"], result.get("message", "")
        rows = get_data_rows(result)
        # 值降序排列: 50(rk=1), 40(rk=2), 30(rk=3), 20(rk=4), 20(rk=4), 10(rk=6)
        # rk <= 2 → 只有 50 和 40
        values = [r[1] for r in rows]
        assert 50 in values
        assert 40 in values
        assert 30 not in values

    def test_where_row_number_filter(self, numbers_file):
        """WHERE rn = 1 — ROW_NUMBER 别名过滤"""
        result = execute_advanced_sql_query(
            numbers_file,
            "SELECT Category, Value, ROW_NUMBER() OVER (PARTITION BY Category ORDER BY Value DESC) AS rn FROM 数值 WHERE rn = 1",
        )
        assert result["success"], result.get("message", "")
        rows = get_data_rows(result)
        # 每个分区的第 1 行: A→20, B→50
        assert len(rows) == 2
        categories = {r[0] for r in rows}
        assert categories == {"A", "B"}

    def test_where_window_with_order_by(self, numbers_file):
        """WHERE + ORDER BY 同时引用窗口别名"""
        result = execute_advanced_sql_query(
            numbers_file,
            "SELECT Value, RANK() OVER (ORDER BY Value DESC) AS rk FROM 数值 WHERE rk <= 3 ORDER BY rk",
        )
        assert result["success"], result.get("message", "")
        rows = get_data_rows(result)
        rks = [r[1] for r in rows]
        assert rks == sorted(rks)  # ORDER BY rk 生效


# ============================================================
# Fix 3: WHERE 引用 SELECT 别名（非窗口）
# ============================================================


class TestWhereSelectAlias:
    """WHERE 引用 SELECT 别名 — 物化为临时列"""

    def test_where_arithmetic_alias(self, simple_file):
        """WHERE double_price > 200 — 算术表达式别名"""
        result = execute_advanced_sql_query(
            simple_file,
            "SELECT ID, Name, Price * 2 AS double_price FROM 数据 WHERE double_price > 200",
        )
        assert result["success"], result.get("message", "")
        rows = get_data_rows(result)
        # Price: 100.5*2=201, 250*2=500, 50*2=100, None, 999.99*2≈2000
        # > 200 → 201, 500, 2000
        assert len(rows) == 3

    def test_where_string_alias(self, simple_file):
        """WHERE upper_name LIKE '%火%' — 字符串函数别名"""
        result = execute_advanced_sql_query(
            simple_file,
            "SELECT ID, UPPER(Name) AS upper_name FROM 数据 WHERE upper_name LIKE '%火%'",
        )
        assert result["success"], result.get("message", "")
        rows = get_data_rows(result)
        # Name: 铁剑, 火球术, 生命药水, None, O'Brien's Sword
        # UPPER → 铁剑, 火球术, 生命药水, None, O'BRIEN'S SWORD
        # LIKE %火% → 火球术
        assert len(rows) == 1
        assert rows[0][0] == 2

    def test_where_alias_no_match(self, simple_file):
        """WHERE alias > 99999 — 无匹配行"""
        result = execute_advanced_sql_query(
            simple_file,
            "SELECT ID, Price * 2 AS dp FROM 数据 WHERE dp > 99999",
        )
        assert result["success"], result.get("message", "")
        rows = get_data_rows(result)
        assert len(rows) == 0


# ============================================================
# Fix 4: run-python 沙箱扩展
# ============================================================


class TestSandboxExpansion:
    """沙箱支持 class 定义和常见异常捕获"""

    def test_class_definition(self, simple_file):
        """class Foo: pass 不再报 NameError"""
        code = """
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y
p = Point(1, 2)
print(f"Point({p.x}, {p.y})")
"""
        result = execute_python_script(simple_file, code, timeout=10)
        assert result["success"], result.get("message", "")
        assert "Point(1, 2)" in result["data"]["stdout"]

    def test_zero_division_catch(self, simple_file):
        """try/except ZeroDivisionError 正常工作"""
        code = """
try:
    result = 1 / 0
except ZeroDivisionError:
    print("caught zero division")
"""
        result = execute_python_script(simple_file, code, timeout=10)
        assert result["success"], result.get("message", "")
        assert "caught zero division" in result["data"]["stdout"]

    def test_attribute_error_catch(self, simple_file):
        """try/except AttributeError 正常工作"""
        code = """
try:
    x = [1, 2, 3]
    x.nonexistent
except AttributeError:
    print("caught attribute error")
"""
        result = execute_python_script(simple_file, code, timeout=10)
        assert result["success"], result.get("message", "")
        assert "caught attribute error" in result["data"]["stdout"]

    def test_json_import(self, simple_file):
        """import json 仍然正常"""
        code = """
import json
data = {"name": "测试", "value": 42}
print(json.dumps(data, ensure_ascii=False))
"""
        result = execute_python_script(simple_file, code, timeout=10)
        assert result["success"], result.get("message", "")
        assert "测试" in result["data"]["stdout"]

    def test_statistics_import(self, simple_file):
        """import statistics 不再被级联阻断"""
        code = """
import statistics
print(statistics.mean([1, 2, 3, 4, 5]))
"""
        result = execute_python_script(simple_file, code, timeout=10)
        assert result["success"], result.get("message", "")
        assert "3" in result["data"]["stdout"]

    def test_class_inheritance(self, simple_file):
        """class 继承 + super() 正常工作"""
        code = """
class Animal:
    def __init__(self, name):
        self.name = name
    def speak(self):
        return f"{self.name} makes a sound"

class Dog(Animal):
    def __init__(self, name, breed):
        super().__init__(name)
        self.breed = breed
    def speak(self):
        return f"{self.name} barks"

d = Dog("Rex", "Labrador")
print(d.speak())
print(d.breed)
"""
        result = execute_python_script(simple_file, code, timeout=10)
        assert result["success"], result.get("message", "")
        assert "Rex barks" in result["data"]["stdout"]
        assert "Labrador" in result["data"]["stdout"]
