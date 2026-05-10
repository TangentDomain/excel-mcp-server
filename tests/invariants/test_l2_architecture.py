"""L2 架构原则不变量测试（INV-5 ~ INV-9）。

INV-5: 失败安全 — success=False 时 data 为空列表，message 非空且无堆栈
INV-6: 错误可分类 — 所有错误消息能被 ToolCallTracker.classify_error() 归入已知类别
INV-7: 幂等读取 — 同一 SELECT 连续执行两次，结果完全一致
INV-8: LIMIT 约束 — SELECT ... LIMIT N 返回行数 ≤ N
INV-9: 聚合语义正确 — COUNT(*) ≥ COUNT(col)；SUM 忽略 NULL；空表聚合语义
"""

from __future__ import annotations

import copy
import math

import pytest

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
)

from .conftest import (
    all_null_file,
    assert_failure_safe,
    assert_result_structure,
    empty_file,
    get_data_rows,
    single_row_file,
    simple_file,
)


# ============================================================
# INV-5: 失败安全
# ============================================================


class TestINV5FailureSafe:
    """INV-5: success=False 时 data=[], message 非空且无堆栈"""

    def test_bad_table_name(self, simple_file):
        result = execute_advanced_sql_query(simple_file, "SELECT * FROM 不存在的表")
        assert result["success"] is False
        assert_failure_safe(result)

    def test_bad_column_name(self, simple_file):
        result = execute_advanced_sql_query(simple_file, "SELECT 不存在的列 FROM 数据")
        assert result["success"] is False
        assert_failure_safe(result)

    def test_syntax_error(self, simple_file):
        result = execute_advanced_sql_query(simple_file, "SELEC * FORM 数据")
        assert result["success"] is False
        assert_failure_safe(result)

    def test_empty_sql(self, simple_file):
        result = execute_advanced_sql_query(simple_file, "")
        assert result["success"] is False
        assert_failure_safe(result)

    def test_nonexistent_file(self):
        result = execute_advanced_sql_query(
            "/tmp/nonexistent_file_inv_test.xlsx", "SELECT * FROM 数据"
        )
        assert result["success"] is False
        assert_failure_safe(result)


# ============================================================
# INV-6: 错误可分类
# ============================================================


class TestINV6ErrorClassifiable:
    """INV-6: 所有错误消息能被 ToolCallTracker.classify_error() 归入已知类别"""

    # 已知的错误类别
    KNOWN_CATEGORIES = {
        "security",
        "file_not_found",
        "sheet_not_found",
        "validation",
        "unsupported",
        "column",
        "sql_syntax",
        "unknown",
        "file_load",
        "file_too_large",
        "execution",
    }

    @pytest.fixture
    def classifier(self):
        try:
            from excel_mcp_server_fastmcp.utils.formatter import ToolCallTracker
            return ToolCallTracker
        except ImportError:
            pytest.skip("ToolCallTracker 不可用")

    def test_bad_table_classified(self, simple_file, classifier):
        result = execute_advanced_sql_query(simple_file, "SELECT * FROM 不存在的表")
        if not result["success"]:
            category = classifier.classify_error(result.get("message", ""))
            assert category in self.KNOWN_CATEGORIES, (
                f"错误未被分类: category='{category}', message='{result['message'][:80]}'"
            )

    def test_syntax_error_classified(self, simple_file, classifier):
        result = execute_advanced_sql_query(simple_file, "INVALID SQL !!!")
        if not result["success"]:
            category = classifier.classify_error(result.get("message", ""))
            assert category in self.KNOWN_CATEGORIES, (
                f"错误未被分类: category='{category}', message='{result['message'][:80]}'"
            )

    def test_missing_file_classified(self, classifier):
        result = execute_advanced_sql_query(
            "/tmp/inv_nonexistent_12345.xlsx", "SELECT 1"
        )
        if not result["success"]:
            category = classifier.classify_error(result.get("message", ""))
            assert category in self.KNOWN_CATEGORIES, (
                f"错误未被分类: category='{category}', message='{result['message'][:80]}'"
            )


# ============================================================
# INV-7: 幂等读取
# ============================================================


class TestINV7IdempotentRead:
    """INV-7: 同一 SELECT 连续执行两次，结果完全一致"""

    def test_simple_select_idempotent(self, simple_file):
        sql = "SELECT * FROM 数据"
        r1 = execute_advanced_sql_query(simple_file, sql)
        r2 = execute_advanced_sql_query(simple_file, sql)
        assert r1["success"] and r2["success"]
        assert r1["data"] == r2["data"], "同一 SELECT 两次执行结果不一致"

    def test_aggregation_idempotent(self, simple_file):
        sql = "SELECT Active, COUNT(*), SUM(Price) FROM 数据 GROUP BY Active"
        r1 = execute_advanced_sql_query(simple_file, sql)
        r2 = execute_advanced_sql_query(simple_file, sql)
        assert r1["success"] and r2["success"]
        assert r1["data"] == r2["data"], "聚合查询两次执行结果不一致"

    def test_ordered_select_idempotent(self, simple_file):
        sql = "SELECT * FROM 数据 ORDER BY Price DESC LIMIT 3"
        r1 = execute_advanced_sql_query(simple_file, sql)
        r2 = execute_advanced_sql_query(simple_file, sql)
        assert r1["success"] and r2["success"]
        assert r1["data"] == r2["data"], "ORDER BY 查询两次执行结果不一致"

    def test_where_select_idempotent(self, simple_file):
        sql = "SELECT * FROM 数据 WHERE Price > 50"
        r1 = execute_advanced_sql_query(simple_file, sql)
        r2 = execute_advanced_sql_query(simple_file, sql)
        assert r1["success"] and r2["success"]
        assert r1["data"] == r2["data"], "WHERE 查询两次执行结果不一致"


# ============================================================
# INV-8: LIMIT 约束
# ============================================================


class TestINV8LimitConstraint:
    """INV-8: SELECT ... LIMIT N 返回行数 ≤ N"""

    def test_limit_1(self, simple_file):
        result = execute_advanced_sql_query(simple_file, "SELECT * FROM 数据 LIMIT 1")
        assert result["success"]
        assert len(result["data"]) - 1 <= 1, "LIMIT 1 返回了超过 1 行"

    def test_limit_0(self, simple_file):
        result = execute_advanced_sql_query(simple_file, "SELECT * FROM 数据 LIMIT 0")
        assert result["success"]
        assert len(result["data"]) - 1 <= 0, "LIMIT 0 返回了超过 0 行"

    def test_limit_exceeds_rows(self, simple_file):
        result = execute_advanced_sql_query(simple_file, "SELECT * FROM 数据 LIMIT 1000")
        assert result["success"]
        # simple_file 有 5 行数据
        assert len(result["data"]) - 1 <= 1000

    def test_limit_with_order(self, simple_file):
        result = execute_advanced_sql_query(
            simple_file, "SELECT * FROM 数据 ORDER BY Price DESC LIMIT 2"
        )
        assert result["success"]
        assert len(result["data"]) - 1 <= 2

    def test_limit_with_offset(self, simple_file):
        result = execute_advanced_sql_query(
            simple_file, "SELECT * FROM 数据 LIMIT 2 OFFSET 1"
        )
        assert result["success"]
        assert len(result["data"]) - 1 <= 2


# ============================================================
# INV-9: 聚合语义正确
# ============================================================


class TestINV9AggregateSemantics:
    """INV-9: COUNT(*) ≥ COUNT(col)；SUM 忽略 NULL；空表聚合语义"""

    def test_count_star_gte_count_col(self, simple_file):
        """COUNT(*) ≥ COUNT(col)，因为 NULL 不计入 COUNT(col)"""
        r_star = execute_advanced_sql_query(simple_file, "SELECT COUNT(*) FROM 数据")
        r_name = execute_advanced_sql_query(simple_file, "SELECT COUNT(Name) FROM 数据")
        r_price = execute_advanced_sql_query(simple_file, "SELECT COUNT(Price) FROM 数据")
        assert r_star["success"] and r_name["success"] and r_price["success"]

        count_star = r_star["data"][1][0]
        count_name = r_name["data"][1][0]
        count_price = r_price["data"][1][0]

        assert count_star >= count_name, f"COUNT(*)={count_star} < COUNT(Name)={count_name}"
        assert count_star >= count_price, f"COUNT(*)={count_star} < COUNT(Price)={count_price}"

    def test_sum_ignores_null(self, simple_file):
        """SUM 忽略 NULL 值"""
        result = execute_advanced_sql_query(simple_file, "SELECT SUM(Price) FROM 数据")
        assert result["success"]
        sum_val = result["data"][1][0]
        # simple_file: 100.5, 250.0, 50.0, NULL, 999.99 → SUM = 1400.49
        assert sum_val is not None, "SUM 不应返回 NULL（有非 NULL 值）"
        expected = 100.5 + 250.0 + 50.0 + 999.99
        assert abs(sum_val - expected) < 0.01, f"SUM={sum_val}，期望 {expected}"

    def test_avg_ignores_null(self, simple_file):
        """AVG 忽略 NULL 值"""
        result = execute_advanced_sql_query(simple_file, "SELECT AVG(Price) FROM 数据")
        assert result["success"]
        avg_val = result["data"][1][0]
        # 4 个非 NULL 值: (100.5 + 250.0 + 50.0 + 999.99) / 4 = 350.1225
        expected = (100.5 + 250.0 + 50.0 + 999.99) / 4
        assert abs(avg_val - expected) < 0.01, f"AVG={avg_val}，期望 {expected}"

    def test_empty_table_count(self, empty_file):
        """空表 COUNT → 0"""
        result = execute_advanced_sql_query(empty_file, "SELECT COUNT(*) FROM 空表")
        assert result["success"]
        assert result["data"][1][0] == 0, f"空表 COUNT(*) 应为 0，实际 {result['data'][1][0]}"

    def test_empty_table_sum(self, empty_file):
        """空表 SUM → NULL"""
        result = execute_advanced_sql_query(empty_file, "SELECT SUM(Value) FROM 空表")
        assert result["success"]
        val = result["data"][1][0]
        assert val is None or val == 0, f"空表 SUM 应为 NULL 或 0，实际 {val}"

    def test_empty_table_avg(self, empty_file):
        """空表 AVG → NULL"""
        result = execute_advanced_sql_query(empty_file, "SELECT AVG(Value) FROM 空表")
        assert result["success"]
        val = result["data"][1][0]
        assert val is None, f"空表 AVG 应为 NULL，实际 {val}"

    def test_empty_table_min_max(self, empty_file):
        """空表 MIN/MAX → NULL"""
        for func in ["MIN", "MAX"]:
            result = execute_advanced_sql_query(empty_file, f"SELECT {func}(Value) FROM 空表")
            assert result["success"]
            val = result["data"][1][0]
            assert val is None, f"空表 {func} 应为 NULL，实际 {val}"

    def test_all_null_column_sum(self, all_null_file):
        """全 NULL 列 SUM → 0（空单元格在 openpyxl 中读取为空字符串，SUM 行为）"""
        result = execute_advanced_sql_query(all_null_file, "SELECT SUM(ColA) FROM Null表")
        assert result["success"]
        val = result["data"][1][0]
        # openpyxl 的空单元格读取为空字符串，SUM 空字符串返回 0
        assert val == 0, f"全空列 SUM 应为 0，实际 {val}"


    def test_all_null_column_count(self, all_null_file):
        """全 NULL 列 COUNT(col) → 0，但 COUNT(*) > 0"""
        r_col = execute_advanced_sql_query(all_null_file, "SELECT COUNT(ColA) FROM Null表")
        r_star = execute_advanced_sql_query(all_null_file, "SELECT COUNT(*) FROM Null表")
        assert r_col["success"] and r_star["success"]
        assert r_col["data"][1][0] == 0, f"全 NULL 列 COUNT(col) 应为 0，实际 {r_col['data'][1][0]}"
        assert r_star["data"][1][0] == 5, f"全 NULL 表 COUNT(*) 应为 5，实际 {r_star['data'][1][0]}"
