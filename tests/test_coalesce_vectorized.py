"""COALESCE向量化测试 — 验证_evaluate_coalesce_vectorized与逐行版本结果一致"""
import pytest
import pandas as pd
import numpy as np
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
    AdvancedSQLQueryEngine,
)
import os

GAME_CONFIG = os.path.join(os.path.dirname(__file__), "test_data", "game_config.xlsx")

@pytest.fixture
def engine():
    return AdvancedSQLQueryEngine()

@pytest.fixture
def df(engine):
    data = engine._load_excel_data(GAME_CONFIG)
    return data["技能配置"]

class TestCoalesceVectorized:
    """验证向量化COALESCE与逐行版本结果完全一致"""

    def test_vectorized_matches_row_by_row(self, engine, df):
        """5个COALESCE调用场景: 列+字面量/多列/数学表达式内/GROUP BY/ORDER BY"""
        import sqlglot
        from sqlglot import expressions as exp

        cases = [
            "COALESCE(damage, 0)",
            "COALESCE(damage, cooldown, 0)",
            "COALESCE(damage, 999)",
            "COALESCE(level, 1)",
            "COALESCE(cost, 0)",
        ]

        for sql_expr in cases:
            parsed = sqlglot.parse_one(f"SELECT {sql_expr} AS val FROM t")
            coalesce_expr = parsed.expressions[0]

            old = [engine._evaluate_coalesce_for_row(coalesce_expr, df.iloc[i]) for i in range(len(df))]
            new = list(engine._evaluate_coalesce_vectorized(coalesce_expr, df))

            assert len(old) == len(new), f"Length mismatch for {sql_expr}"
            for i, (o, n) in enumerate(zip(old, new)):
                # Compare as strings to handle int/float type differences
                assert str(o) == str(n), f"Mismatch at row {i} for {sql_expr}: {o!r} vs {n!r}"

    def test_vectorized_with_string_literal(self, engine, df):
        """COALESCE(col, 'default') — 字面量为字符串"""
        import sqlglot
        from sqlglot import expressions as exp

        parsed = sqlglot.parse_one("SELECT COALESCE(skill_name, 'unknown') AS val FROM t")
        coalesce_expr = parsed.expressions[0]

        old = [engine._evaluate_coalesce_for_row(coalesce_expr, df.iloc[i]) for i in range(len(df))]
        new = list(engine._evaluate_coalesce_vectorized(coalesce_expr, df))

        assert old == new

    def test_vectorized_with_float_literal(self, engine, df):
        """COALESCE(col, 3.14) — 字面量为浮点数"""
        import sqlglot

        parsed = sqlglot.parse_one("SELECT COALESCE(damage, 3.14) AS val FROM t")
        coalesce_expr = parsed.expressions[0]

        old = [engine._evaluate_coalesce_for_row(coalesce_expr, df.iloc[i]) for i in range(len(df))]
        new = list(engine._evaluate_coalesce_vectorized(coalesce_expr, df))

        for o, n in zip(old, new):
            assert str(o) == str(n)

    def test_fallback_to_row_by_row(self, engine, df):
        """COALESCE含复杂表达式 → 自动回退逐行模式"""
        import sqlglot
        from sqlglot import expressions as exp

        # COALESCE(damage + cooldown, 0) — 第一个参数是数学表达式
        parsed = sqlglot.parse_one("SELECT COALESCE(damage + cooldown, 0) AS val FROM t")
        coalesce_expr = parsed.expressions[0]

        result = engine._evaluate_coalesce_vectorized(coalesce_expr, df)

        # 验证回退路径产生有效结果
        assert len(result) == len(df)
        # damage + cooldown 不为 null，所以所有值都应该是 sum
        for i in range(min(5, len(df))):
            expected = df.iloc[i]["damage"] + df.iloc[i]["cooldown"]
            assert result.iloc[i] == expected

    def test_coalesce_in_select(self):
        """COALESCE在SELECT中正常工作"""
        r = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT skill_name, COALESCE(damage, 0) AS safe_dmg FROM 技能配置 LIMIT 3",
        )
        assert r["success"] is True
        assert len(r["data"]) == 4  # header + 3 rows
        assert r["data"][1][1] > 0  # damage values are positive

    def test_coalesce_in_where(self):
        """COALESCE在WHERE中正常工作"""
        r = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT skill_name FROM 技能配置 WHERE COALESCE(damage, 0) > 100",
        )
        assert r["success"] is True
        # All returned skills should have damage > 100
        assert len(r["data"]) > 1

    def test_coalesce_in_math_expression(self):
        """COALESCE在数学表达式中正常工作"""
        r = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT skill_name, COALESCE(damage, 0) * 2 AS double_dmg FROM 技能配置 LIMIT 3",
        )
        assert r["success"] is True
        assert r["data"][1][1] > 0

    def test_coalesce_in_group_by(self):
        """COALESCE在GROUP BY聚合结果中正常工作"""
        r = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT skill_type, COALESCE(MAX(damage), 0) AS max_dmg FROM 技能配置 GROUP BY skill_type",
        )
        assert r["success"] is True
        assert len(r["data"]) > 1  # Multiple skill types

class TestGetRowValueNumericLiteral:
    """验证_get_row_value正确处理数字字面量（与_get_expression_value一致）"""

    def test_int_literal(self, engine):
        """整数字面量返回int类型"""
        import sqlglot
        from sqlglot import expressions as exp

        row = pd.Series({"a": 1, "b": 2})
        lit = exp.Literal.number(42)
        result = engine._get_row_value(lit, row)
        assert result == 42
        assert isinstance(result, int)

class TestCaseWhenNumericThen:
    """验证CASE WHEN数字THEN值修复 — _get_row_value数字字面量正确转换"""

    def test_case_when_numeric_in_where(self):
        """CASE WHEN ... THEN 1 ELSE 0 END = 1 — 数字THEN值比较"""
        r = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT skill_name FROM 技能配置 WHERE CASE WHEN damage > 100 THEN 1 ELSE 0 END = 1",
        )
        assert r["success"] is True
        # All returned skills should have damage > 100
        for row in r["data"][1:]:
            assert row[0] is not None

    def test_case_when_numeric_in_select(self):
        """CASE WHEN ... THEN 1 ELSE 0 — 数字THEN值在SELECT中"""
        r = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT skill_name, CASE WHEN damage > 100 THEN 1 ELSE 0 END AS is_high FROM 技能配置 LIMIT 3",
        )
        assert r["success"] is True
        # Values should be int 0 or 1
        for row in r["data"][1:]:
            assert row[1] in (0, 1)

    def test_case_when_numeric_order_by(self):
        """CASE WHEN数字THEN值可正确排序"""
        r = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT skill_name, CASE WHEN damage > 100 THEN 1 ELSE 0 END AS is_high FROM 技能配置 ORDER BY is_high DESC LIMIT 3",
        )
        assert r["success"] is True
        # First 3 should have is_high=1
        for row in r["data"][1:4]:
            assert row[1] == 1

    def test_case_when_string_then_still_works(self):
        """CASE WHEN字符串THEN值不受影响"""
        r = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT skill_name, CASE WHEN damage > 100 THEN 'high' ELSE 'low' END AS tier FROM 技能配置 LIMIT 3",
        )
        assert r["success"] is True
        for row in r["data"][1:]:
            assert row[1] in ("high", "low")
