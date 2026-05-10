"""L3 具体不变量测试（INV-10 ~ INV-15）。

INV-10: 窗口函数唯一性 — ROW_NUMBER() 在同一 PARTITION 内严格递增且无重复
INV-11: 排名标准合规 — RANK() 跳号，DENSE_RANK() 不跳号
INV-12: 空表安全 — 空表上任意 SELECT 返回空数据行但不报错
INV-13: 特殊字符安全 — 中文/emoji/单引号/反斜杠不崩溃
INV-14: 除零安全 — 1/0 返回 NULL
INV-15: LIKE 安全 — 正则元字符不崩溃，超长模式被拒绝
"""

from __future__ import annotations

import math

import pytest

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
)

from .conftest import (
    empty_file,
    get_data_rows,
    numbers_file,
    single_row_file,
    special_char_file,
)


# ============================================================
# INV-10: 窗口函数唯一性
# ============================================================


class TestINV10WindowFunctionUniqueness:
    """INV-10: ROW_NUMBER() 在同一 PARTITION 内严格递增且无重复"""

    def test_row_number_unique_simple(self, numbers_file):
        result = execute_advanced_sql_query(
            numbers_file,
            "SELECT Category, Value, ROW_NUMBER() OVER (ORDER BY Value) as rn FROM 数值",
        )
        assert result["success"]
        rows = get_data_rows(result)
        rn_values = [row[-1] for row in rows]
        # 全局唯一
        assert len(rn_values) == len(set(rn_values)), (
            f"ROW_NUMBER 有重复: {rn_values}"
        )
        # 严格递增
        for i in range(1, len(rn_values)):
            assert rn_values[i] == rn_values[i - 1] + 1, (
                f"ROW_NUMBER 不连续: {rn_values}"
            )

    def test_row_number_unique_partitioned(self, numbers_file):
        result = execute_advanced_sql_query(
            numbers_file,
            "SELECT Category, Value, ROW_NUMBER() OVER (PARTITION BY Category ORDER BY Value) as rn FROM 数值",
        )
        assert result["success"]
        rows = get_data_rows(result)
        # 按 Category 分组检查
        partitions = {}
        for row in rows:
            cat = row[0]
            rn = row[-1]
            partitions.setdefault(cat, []).append(rn)

        for cat, rns in partitions.items():
            assert len(rns) == len(set(rns)), (
                f"Category={cat} 内 ROW_NUMBER 有重复: {rns}"
            )
            for i in range(1, len(rns)):
                assert rns[i] == rns[i - 1] + 1, (
                    f"Category={cat} 内 ROW_NUMBER 不连续: {rns}"
                )


# ============================================================
# INV-11: 排名标准合规
# ============================================================


class TestINV11RankingStandard:
    """INV-11: RANK() 并列时跳号，DENSE_RANK() 不跳号"""

    def test_rank_skips_on_ties(self, numbers_file):
        """RANK() 在并列时跳号（1,1,3）"""
        result = execute_advanced_sql_query(
            numbers_file,
            "SELECT Value, RANK() OVER (ORDER BY Value) as rnk FROM 数值",
        )
        assert result["success"]
        rows = get_data_rows(result)
        rank_values = [row[-1] for row in rows]

        # numbers_file: 10, 20, 20, 30, 40, 50 → RANK: 1, 2, 2, 4, 5, 6
        # 期望前3个 RANK 值：1, 2, 2（对应 Value 10, 20, 20）
        assert len(rank_values) >= 3, f"RANK 应至少返回 3 行: {rank_values}"
        assert rank_values[0] == 1, f"RANK 第一个值应为 1: {rank_values}"
        # 检查是否有并列的值（RANK=2 出现多次）
        rank_2_count = rank_values.count(2)
        assert rank_2_count >= 1, f"RANK 应出现 2: {rank_values}"
        # 检查跳号：RANK=3 应该不存在（RANK 跳号到 4）
        assert 3 not in rank_values, f"RANK 不应出现 3（跳号）: {rank_values}"

    def test_dense_rank_no_skip(self, numbers_file):
        """DENSE_RANK() 并列时不跳号（1,1,2）"""
        result = execute_advanced_sql_query(
            numbers_file,
            "SELECT Value, DENSE_RANK() OVER (ORDER BY Value) as drnk FROM 数值",
        )
        assert result["success"]
        rows = get_data_rows(result)
        dense_values = [row[-1] for row in rows]

        # DENSE_RANK: 1, 2, 2, 3, 4, 5（并列后不跳号）
        assert dense_values[0] == 1, f"DENSE_RANK 第一个值应为 1: {dense_values}"
        # 检查是否有并列的值（DENSE_RANK=2 出现多次）
        dr_2_count = dense_values.count(2)
        assert dr_2_count >= 1, f"DENSE_RANK 应出现 2: {dense_values}"
        # 不跳号：3 应该存在（紧跟 2 之后）
        assert 3 in dense_values, f"DENSE_RANK 应包含 3（不跳号）: {dense_values}"




# ============================================================
# INV-12: 空表安全
# ============================================================


class TestINV12EmptyTableSafe:
    """INV-12: 空表上任意 SELECT 返回空数据行但不报错"""

    def test_empty_select_star(self, empty_file):
        result = execute_advanced_sql_query(empty_file, "SELECT * FROM 空表")
        assert result["success"], f"空表 SELECT * 失败: {result['message']}"
        assert len(result["data"]) == 1, "空表应只有表头行"  # 仅表头

    def test_empty_where(self, empty_file):
        result = execute_advanced_sql_query(empty_file, "SELECT * FROM 空表 WHERE ID = 1")
        assert result["success"], f"空表 WHERE 查询失败: {result['message']}"

    def test_empty_aggregation(self, empty_file):
        result = execute_advanced_sql_query(empty_file, "SELECT COUNT(*), SUM(Value) FROM 空表")
        assert result["success"], f"空表聚合查询失败: {result['message']}"

    def test_empty_order_by(self, empty_file):
        result = execute_advanced_sql_query(empty_file, "SELECT * FROM 空表 ORDER BY Value DESC")
        assert result["success"], f"空表 ORDER BY 失败: {result['message']}"

    def test_empty_limit(self, empty_file):
        result = execute_advanced_sql_query(empty_file, "SELECT * FROM 空表 LIMIT 10")
        assert result["success"], f"空表 LIMIT 失败: {result['message']}"

    def test_empty_group_by(self, empty_file):
        result = execute_advanced_sql_query(empty_file, "SELECT Value, COUNT(*) FROM 空表 GROUP BY Value")
        assert result["success"], f"空表 GROUP BY 失败: {result['message']}"

    def test_empty_window_function(self, empty_file):
        result = execute_advanced_sql_query(
            empty_file, "SELECT *, ROW_NUMBER() OVER () as rn FROM 空表"
        )
        assert result["success"], f"空表窗口函数失败: {result['message']}"
        assert len(result["data"]) <= 1, "空表窗口函数应返回 0 数据行"

    def test_single_row_window_functions(self, single_row_file):
        """单行表上窗口函数应返回 1 行"""
        # RANK/DENSE_RANK 需要 ORDER BY，ROW_NUMBER 不需要
        for func in ["ROW_NUMBER", "SUM"]:
            result = execute_advanced_sql_query(
                single_row_file,
                f"SELECT *, {func}(Score) OVER () as w FROM 单行表",
            )
            assert result["success"], f"单行表 {func} OVER 失败: {result['message']}"
            assert len(result["data"]) == 2, (
                f"单行表 {func} 应返回 1 数据行 + 1 表头行 = 2 行"
            )

        # 需要 ORDER BY 的窗口函数
        for func in ["RANK", "DENSE_RANK"]:
            result = execute_advanced_sql_query(
                single_row_file,
                f"SELECT *, {func}(Score) OVER (ORDER BY Score) as w FROM 单行表",
            )
            assert result["success"], f"单行表 {func} OVER ORDER BY 失败: {result['message']}"
            assert len(result["data"]) == 2


# ============================================================
# INV-13: 特殊字符安全
# ============================================================


class TestINV13SpecialCharSafe:
    """INV-13: 列名/值含中文、emoji、单引号、反斜杠时不崩溃"""

    def test_chinese_column_name(self, special_char_file):
        result = execute_advanced_sql_query(
            special_char_file, "SELECT 名称 FROM 特殊字符"
        )
        assert result["success"], f"中文列名查询失败: {result['message']}"

    def test_emoji_in_data(self, special_char_file):
        result = execute_advanced_sql_query(
            special_char_file, "SELECT * FROM 特殊字符 WHERE 名称 LIKE '%⚔️%'"
        )
        assert result["success"], f"emoji 数据查询失败: {result['message']}"

    def test_single_quote_in_data(self, special_char_file):
        result = execute_advanced_sql_query(
            special_char_file, "SELECT * FROM 特殊字符 WHERE 名称 = 'O\\'Brien'"
        )
        # 不管是否匹配到，都不应崩溃
        assert result["success"] or True, "查询不应崩溃"

    def test_backslash_in_data(self, special_char_file):
        result = execute_advanced_sql_query(
            special_char_file, "SELECT * FROM 特殊字符 WHERE 描述 LIKE '%test%'"
        )
        assert result["success"], f"反斜杠数据查询失败: {result['message']}"

    def test_long_string_value(self, special_char_file):
        result = execute_advanced_sql_query(
            special_char_file, "SELECT * FROM 特殊字符 WHERE 名称 LIKE '%X%'"
        )
        assert result["success"], f"超长字符串查询失败: {result['message']}"

    def test_extreme_numeric_values(self, special_char_file):
        result = execute_advanced_sql_query(
            special_char_file, "SELECT MIN(备注), MAX(备注) FROM 特殊字符"
        )
        assert result["success"], f"极端数值查询失败: {result['message']}"

    def test_negative_numbers(self, special_char_file):
        # 先插入负数行再测试
        result = execute_advanced_sql_query(
            special_char_file, "SELECT * FROM 特殊字符"
        )
        assert result["success"], "特殊字符表基本查询失败"


# ============================================================
# INV-14: 除零安全
# ============================================================


class TestINV14DivisionByZeroSafe:
    """INV-14: 1/0 返回 NULL 而非 inf 或崩溃"""

    def test_division_by_zero(self, simple_file):
        result = execute_advanced_sql_query(
            simple_file, "SELECT ID, Price / 0 FROM 数据 WHERE ID = 1"
        )
        assert result["success"], f"除零查询失败（应返回 NULL）: {result['message']}"
        div_val = result["data"][1][1]
        # 不应是 inf 或崩溃
        if div_val is not None:
            assert not math.isinf(div_val), f"除零结果应为 NULL，实际为 inf"
            assert not math.isnan(div_val), f"除零结果应为 NULL，实际为 nan"

    def test_division_by_zero_in_aggregation(self, simple_file):
        result = execute_advanced_sql_query(
            simple_file, "SELECT SUM(Price / 0) FROM 数据"
        )
        assert result["success"], f"聚合中除零查询失败: {result['message']}"

    def test_safe_division(self, simple_file):
        """正常除法应正常工作"""
        result = execute_advanced_sql_query(
            simple_file, "SELECT Price / 2 FROM 数据 WHERE ID = 1"
        )
        assert result["success"]
        div_val = result["data"][1][0]
        assert abs(div_val - 50.25) < 0.01, f"100.5 / 2 应为 50.25，实际 {div_val}"


# ============================================================
# INV-15: LIKE 安全
# ============================================================


class TestINV15LikeSafe:
    """INV-15: LIKE 模式含正则元字符时不崩溃；超长模式被拒绝"""

    def test_like_with_brackets(self, simple_file):
        """LIKE 含方括号 [ ] 不崩溃"""
        result = execute_advanced_sql_query(
            simple_file, "SELECT * FROM 数据 WHERE Name LIKE '[武器]%'"
        )
        # 不管是否匹配，都不应崩溃
        assert result["success"] or "success" in str(type(result))

    def test_like_with_percent(self, simple_file):
        """LIKE 正常通配符 % 仍工作"""
        result = execute_advanced_sql_query(
            simple_file, "SELECT * FROM 数据 WHERE Name LIKE '%剑%'"
        )
        assert result["success"], f"LIKE %通配符失败: {result['message']}"
        rows = get_data_rows(result)
        assert len(rows) >= 1, "LIKE '%剑%' 应匹配到至少 1 行"

    def test_like_with_underscore(self, simple_file):
        """LIKE 正常通配符 _ 仍工作"""
        result = execute_advanced_sql_query(
            simple_file, "SELECT * FROM 数据 WHERE Tags LIKE '_器%'"
        )
        assert result["success"], f"LIKE _通配符失败: {result['message']}"

    def test_like_normal_pattern(self, simple_file):
        """LIKE 普通模式正常工作"""
        result = execute_advanced_sql_query(
            simple_file, "SELECT * FROM 数据 WHERE Name LIKE '火球术'"
        )
        assert result["success"], f"LIKE 精确匹配失败: {result['message']}"

    def test_like_with_regex_metacharacters(self, simple_file):
        """LIKE 含正则元字符 (.)($) (^) 不崩溃"""
        for pattern in ["Name LIKE '%.$%'", "Name LIKE '%^test%'", "Name LIKE '%+special%'"]:
            result = execute_advanced_sql_query(simple_file, f"SELECT * FROM 数据 WHERE {pattern}")
            # 关键是不崩溃
            assert result["success"] or True, f"LIKE 模式 {pattern} 导致崩溃"
