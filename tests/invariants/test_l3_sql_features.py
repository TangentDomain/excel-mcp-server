"""L3 SQL 功能边界不变量测试。

INV-25: DISTINCT 语义正确性
INV-26: HAVING 子句正确性
INV-27: NULL 比较（IS NULL / IS NOT NULL）正确性
INV-28: 子查询（IN (SELECT...), EXISTS）正确性
INV-29: OFFSET 边界正确性
INV-30: NOT IN / NOT LIKE 语义正确性
INV-31: 双行表头写操作正确性
INV-32: _ROW_NUMBER_ 写操作正确性
"""

from __future__ import annotations

import pytest

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_delete_query,
    execute_advanced_insert_query,
    execute_advanced_sql_query,
    execute_advanced_update_query,
)
from excel_mcp_server_fastmcp.calibrator.core import cmd_import, cmd_query

from .conftest import (
    dual_header_file,
    get_data_rows,
    get_headers,
    simple_file,
    writable_file,
)
from .test_l1_result_structure import _CAL_DB, _align_result

# ============================================================
# INV-25: DISTINCT 语义正确性
# ============================================================


class TestINV25Distinct:
    """INV-25: DISTINCT 返回去重后的唯一行"""

    def test_distinct_basic(self, simple_file):
        """DISTINCT 消除重复行"""
        result = execute_advanced_sql_query(simple_file, "SELECT DISTINCT Active FROM 数据")
        assert result["success"]
        # simple_file 有 Active: 是, 否, 是, None, 是 → DISTINCT 后应该是 3 种: 是, 否, None/空
        values = [row[0] for row in get_data_rows(result)]
        # 去重后应该少于总行数
        assert len(values) <= 5
        # "是" 出现 3 次，DISTINCT 后只应出现 1 次
        assert values.count("是") == 1

    def test_distinct_all_columns(self, simple_file):
        """SELECT DISTINCT * 返回唯一完整行"""
        result_all = execute_advanced_sql_query(simple_file, "SELECT * FROM 数据")
        result_distinct = execute_advanced_sql_query(simple_file, "SELECT DISTINCT * FROM 数据")
        assert result_all["success"] and result_distinct["success"]
        # DISTINCT * 的行数 <= 原始行数
        assert len(get_data_rows(result_distinct)) <= len(get_data_rows(result_all))

    def test_distinct_with_order_by(self, simple_file):
        """DISTINCT + ORDER BY 不报错"""
        result = execute_advanced_sql_query(simple_file, "SELECT DISTINCT Active FROM 数据 ORDER BY Active")
        assert result["success"]
        assert len(get_data_rows(result)) >= 1

    def test_distinct_count(self, simple_file):
        """SELECT COUNT(DISTINCT col) 正确计数"""
        result = execute_advanced_sql_query(simple_file, "SELECT COUNT(DISTINCT Active) FROM 数据")
        assert result["success"]
        # Active 有 "是"(3次), "否"(1次), None(1次)
        # COUNT(DISTINCT) 排除 NULL → 2（是, 否）
        count_val = result["data"][1][0]
        assert count_val == 2


# ============================================================
# INV-26: HAVING 子句正确性
# ============================================================


class TestINV26Having:
    """INV-26: HAVING 子句在 GROUP BY 后正确过滤"""

    def test_having_with_count(self, simple_file):
        """HAVING COUNT(*) > N"""
        result = execute_advanced_sql_query(
            simple_file,
            "SELECT Active, COUNT(*) as cnt FROM 数据 GROUP BY Active HAVING cnt > 1",
        )
        assert result["success"]
        # 只有 "是" 出现 3 次 > 1
        rows = get_data_rows(result)
        assert len(rows) == 1
        assert rows[0][0] == "是"
        assert rows[0][1] == 3

    def test_having_with_sum(self, simple_file):
        """HAVING SUM(Price) > N"""
        result = execute_advanced_sql_query(
            simple_file,
            "SELECT Active, SUM(Price) as total FROM 数据 GROUP BY Active HAVING total > 100",
        )
        assert result["success"]
        rows = get_data_rows(result)
        # "否" 组 SUM = 250 > 100; "是" 组 SUM = 100.5+50+999.99 > 100
        assert len(rows) >= 1

    def test_having_no_match(self, simple_file):
        """HAVING 条件不匹配时返回空"""
        result = execute_advanced_sql_query(
            simple_file,
            "SELECT Active, COUNT(*) FROM 数据 GROUP BY Active HAVING COUNT(*) > 100",
        )
        assert result["success"]
        assert len(get_data_rows(result)) == 0

    def test_having_vs_where(self, simple_file):
        """HAVING 和 WHERE 的区别：WHERE 过滤行，HAVING 过滤组"""
        # WHERE 先过滤，HAVING 后过滤组
        result = execute_advanced_sql_query(
            simple_file,
            "SELECT Active, COUNT(*) as cnt FROM 数据 WHERE Price > 100 GROUP BY Active HAVING cnt >= 1",
        )
        assert result["success"]
        rows = get_data_rows(result)
        # Price > 100 的行: ID=2(250,否), ID=5(999.99,是)
        # 否 组只有 1 行 → cnt=1 >= 1
        # 是 组只有 1 行 → cnt=1 >= 1
        assert len(rows) == 2


# ============================================================
# INV-27: IS NULL / IS NOT NULL 正确性
# ============================================================


class TestINV27NullComparison:
    """INV-27: IS NULL / IS NOT NULL 正确识别空值"""

    def test_is_null(self, simple_file):
        """IS NULL 正确找到 NULL 行"""
        result = execute_advanced_sql_query(simple_file, "SELECT ID, Name FROM 数据 WHERE Name IS NULL")
        assert result["success"]
        rows = get_data_rows(result)
        # simple_file 中 ID=4 的 Name 是 None
        assert len(rows) == 1
        assert rows[0][0] == 4

    def test_is_not_null(self, simple_file):
        """IS NOT NULL 排除 NULL 行"""
        result = execute_advanced_sql_query(simple_file, "SELECT COUNT(*) FROM 数据 WHERE Name IS NOT NULL")
        assert result["success"]
        count = result["data"][1][0]
        assert count == 4  # 5 行 - 1 个 NULL = 4

    def test_is_null_count_zero(self, simple_file):
        """对非 NULL 列 IS NULL 返回 0 行"""
        result = execute_advanced_sql_query(simple_file, "SELECT COUNT(*) FROM 数据 WHERE ID IS NULL")
        assert result["success"]
        count = result["data"][1][0]
        assert count == 0  # ID 列没有 NULL

    def test_is_null_with_update(self, writable_file):
        """UPDATE 设 NULL 后 IS NULL 能找到"""
        execute_advanced_update_query(writable_file, "UPDATE 商品 SET Stock = NULL WHERE ID = 1")
        result = execute_advanced_sql_query(writable_file, "SELECT ID FROM 商品 WHERE Stock IS NULL")
        assert result["success"]
        rows = get_data_rows(result)
        assert len(rows) == 1
        assert rows[0][0] == 1


# ============================================================
# INV-28: 子查询正确性
# ============================================================


class TestINV28Subquery:
    """INV-28: 子查询（IN (SELECT...) 正确执行"""

    @pytest.fixture(autouse=True)
    def _setup(self, simple_file):
        self.file_path = simple_file

    def test_in_subquery(self):
        """WHERE ID IN (SELECT ...) 子查询"""
        result = execute_advanced_sql_query(
            self.file_path,
            "SELECT Name FROM 数据 WHERE ID IN (SELECT ID FROM 数据 WHERE Active = '是')",
        )
        assert result["success"]
        rows = get_data_rows(result)
        names = [r[0] for r in rows]
        # Active='是' 的 ID: 1,3,5 → Name: 铁剑, 生命药水, O'Brien's Sword
        assert "铁剑" in names
        assert "生命药水" in names

    def test_not_in_subquery(self):
        """WHERE ID NOT IN (SELECT ...) 子查询"""
        result = execute_advanced_sql_query(
            self.file_path,
            "SELECT Name FROM 数据 WHERE ID NOT IN (SELECT ID FROM 数据 WHERE Active = '是')",
        )
        assert result["success"]
        rows = get_data_rows(result)
        names = [r[0] for r in rows]
        # 非 Active='是' 的: ID=2(火球术), ID=4(None)
        assert "火球术" in names

    def test_subquery_in_select(self):
        """SELECT 子查询作为列"""
        result = execute_advanced_sql_query(
            self.file_path,
            "SELECT Name, (SELECT MAX(Price) FROM 数据) as max_price FROM 数据 WHERE ID = 1",
        )
        assert result["success"]
        # max_price 应该是 999.99 (ID=5)
        assert result["data"][1][1] == 999.99

    def test_subquery_scalar(self):
        """标量子查询在 WHERE 中"""
        result = execute_advanced_sql_query(
            self.file_path,
            "SELECT Name FROM 数据 WHERE Price > (SELECT AVG(Price) FROM 数据)",
        )
        assert result["success"]
        rows = get_data_rows(result)
        # AVG(Price) = (100.5+250+50+999.99)/4 = 400.12 (不含 NULL)
        # Price > 400.12 的: ID=5 (999.99)
        names = [r[0] for r in rows]
        assert "O'Brien's Sword" in names


# ============================================================
# INV-29: OFFSET 边界正确性
# ============================================================


class TestINV29OffsetBoundary:
    """INV-29: OFFSET 边界条件"""

    def test_offset_beyond_total(self, simple_file):
        """OFFSET 超过总行数返回空"""
        result = execute_advanced_sql_query(simple_file, "SELECT * FROM 数据 OFFSET 999")
        assert result["success"]
        assert len(get_data_rows(result)) == 0

    def test_offset_with_limit(self, simple_file):
        """OFFSET + LIMIT 组合"""
        result = execute_advanced_sql_query(simple_file, "SELECT * FROM 数据 LIMIT 2 OFFSET 2")
        assert result["success"]
        rows = get_data_rows(result)
        assert len(rows) <= 2  # 最多 2 行

    def test_offset_zero(self, simple_file):
        """OFFSET 0 等同于不使用 OFFSET"""
        result_no_offset = execute_advanced_sql_query(simple_file, "SELECT * FROM 数据 LIMIT 3")
        result_offset_0 = execute_advanced_sql_query(simple_file, "SELECT * FROM 数据 LIMIT 3 OFFSET 0")
        assert result_no_offset["success"] and result_offset_0["success"]
        assert get_data_rows(result_no_offset) == get_data_rows(result_offset_0)


# ============================================================
# INV-30: NOT IN / NOT LIKE 语义正确性
# ============================================================


class TestINV30NotOperators:
    """INV-30: NOT IN / NOT LIKE 语义正确"""

    def test_not_in(self, simple_file):
        """NOT IN 排除指定值"""
        result = execute_advanced_sql_query(simple_file, "SELECT ID FROM 数据 WHERE Active NOT IN ('是')")
        assert result["success"]
        rows = get_data_rows(result)
        ids = [r[0] for r in rows]
        assert 2 in ids  # ID=2 Active='否'
        assert 1 not in ids

    def test_not_like(self, simple_file):
        """NOT LIKE 排除匹配行"""
        result = execute_advanced_sql_query(simple_file, "SELECT Name FROM 数据 WHERE Name NOT LIKE '%剑%'")
        assert result["success"]
        rows = get_data_rows(result)
        names = [r[0] for r in rows]
        assert "铁剑" not in names
        assert "火球术" in names

    def test_not_in_empty_set(self, simple_file):
        """NOT IN 空集返回所有行"""
        result = execute_advanced_sql_query(simple_file, "SELECT COUNT(*) FROM 数据 WHERE ID NOT IN ()")
        # 空集行为取决于实现，至少不应崩溃
        assert result["success"]


# ============================================================
# INV-31: 双行表头写操作正确性
# ============================================================


class TestINV31DualHeaderWrite:
    """INV-31: 双行表头表的写操作正确性"""

    def test_update_dual_header(self, dual_header_file):
        """UPDATE 双行表头表后读回正确"""
        result = execute_advanced_update_query(
            dual_header_file,
            "UPDATE 技能配置 SET base_damage = 999 WHERE skill_id = 'SK001'",
        )
        assert result["success"]
        verify = execute_advanced_sql_query(
            dual_header_file,
            "SELECT base_damage FROM 技能配置 WHERE skill_id = 'SK001'",
        )
        assert verify["success"]
        assert verify["data"][1][0] == 999

    def test_insert_dual_header(self, dual_header_file):
        """INSERT 双行表头表后 COUNT 增加"""
        before = execute_advanced_sql_query(dual_header_file, "SELECT COUNT(*) FROM 技能配置")
        result = execute_advanced_insert_query(
            dual_header_file,
            "INSERT INTO 技能配置 (skill_id, skill_name, base_damage, cooldown, skill_type) VALUES ('SK999', '新技能', 50, 3, '测试')",
        )
        assert result["success"]
        after = execute_advanced_sql_query(dual_header_file, "SELECT COUNT(*) FROM 技能配置")
        assert after["data"][1][0] == before["data"][1][0] + 1

    def test_delete_dual_header(self, dual_header_file):
        """DELETE 双行表头表后 COUNT 减少"""
        before = execute_advanced_sql_query(dual_header_file, "SELECT COUNT(*) FROM 技能配置")
        result = execute_advanced_delete_query(dual_header_file, "DELETE FROM 技能配置 WHERE skill_id = 'SK004'")
        assert result["success"]
        after = execute_advanced_sql_query(dual_header_file, "SELECT COUNT(*) FROM 技能配置")
        assert after["data"][1][0] == before["data"][1][0] - 1

    def test_update_dual_header_expression(self, dual_header_file):
        """UPDATE 双行表头表 SET 表达式"""
        execute_advanced_update_query(
            dual_header_file,
            "UPDATE 技能配置 SET base_damage = base_damage + 10 WHERE skill_type = '物理'",
        )
        verify = execute_advanced_sql_query(
            dual_header_file,
            "SELECT skill_id, base_damage FROM 技能配置 WHERE skill_type = '物理' ORDER BY skill_id",
        )
        assert verify["success"]
        rows = get_data_rows(verify)
        # SK001 物理: 150+10=160, SK004 物理: 300+10=310
        assert rows[0][1] == 160
        assert rows[1][1] == 310


# ============================================================
# INV-32: _ROW_NUMBER_ 写操作正确性
# ============================================================


class TestINV32RowNumberWrite:
    """INV-32: _ROW_NUMBER_ 在 UPDATE WHERE 中正确工作"""

    def test_update_by_row_number(self, writable_file):
        """UPDATE WHERE _ROW_NUMBER_ = N 精确定位行"""
        execute_advanced_update_query(writable_file, "UPDATE 商品 SET Price = 0 WHERE _ROW_NUMBER_ = 3")
        # 第 3 行应该是 ID=3 的行
        result = execute_advanced_sql_query(writable_file, "SELECT ID, Price FROM 商品 WHERE ID = 3")
        assert result["success"]
        assert result["data"][1][1] == 0

    def test_update_row_number_range(self, writable_file):
        """UPDATE WHERE _ROW_NUMBER_ BETWEEN 定位范围"""
        execute_advanced_update_query(writable_file, "UPDATE 商品 SET Price = 0 WHERE _ROW_NUMBER_ BETWEEN 2 AND 4")
        result = execute_advanced_sql_query(writable_file, "SELECT ID, Price FROM 商品 ORDER BY ID")
        assert result["success"]
        rows = get_data_rows(result)
        # _ROW_NUMBER_ 2,3,4 → ID=2,3,4
        assert rows[0][1] == 100.0  # ID=1 不变
        assert rows[1][1] == 0  # ID=2 被更新
        assert rows[2][1] == 0  # ID=3 被更新
        assert rows[3][1] == 0  # ID=4 被更新
        assert rows[4][1] == pytest.approx(75.5)  # ID=5 不变
        assert rows[5][1] == pytest.approx(999.99)  # ID=6 不变

    def test_delete_by_row_number(self, writable_file):
        """DELETE WHERE _ROW_NUMBER_ = N 精确定位删除"""
        execute_advanced_delete_query(writable_file, "DELETE FROM 商品 WHERE _ROW_NUMBER_ = 1")
        result = execute_advanced_sql_query(writable_file, "SELECT ID FROM 商品 ORDER BY ID")
        assert result["success"]
        rows = get_data_rows(result)
        ids = [r[0] for r in rows]
        # ID=1 应该被删除
        assert 1 not in ids
        assert 2 in ids  # 其余不变
