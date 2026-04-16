"""ORDER BY 函数表达式测试 — 验证UPPER/LENGTH/COALESCE/CASE WHEN等可在ORDER BY中使用"""
import os
import pytest

GAME_CONFIG = os.path.join(os.path.dirname(__file__), "test_data", "game_config.xlsx")


@pytest.fixture
def engine():
    from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
    return AdvancedSQLQueryEngine()


class TestOrderByStringFunctions:
    """ORDER BY 字符串函数"""

    def test_order_by_upper_asc(self, engine):
        """ORDER BY UPPER(col) — 按大写排序"""
        result = engine.execute_sql_query(GAME_CONFIG, "SELECT skill_name FROM 技能配置 ORDER BY UPPER(skill_name)")
        assert result['success'] is True
        names = [row[0] for row in result['data'][1:]]
        # UPPER排序应不区分大小写，冰(bing)在火(huo)前
        assert names[0] in ('冰冻术', '冰风暴')

    def test_order_by_upper_desc(self, engine):
        """ORDER BY UPPER(col) DESC"""
        result = engine.execute_sql_query(GAME_CONFIG, "SELECT skill_name FROM 技能配置 ORDER BY UPPER(skill_name) DESC")
        assert result['success'] is True
        names = [row[0] for row in result['data'][1:]]
        assert len(names) > 0

    def test_order_by_length(self, engine):
        """ORDER BY LENGTH(col) — 按名称长度排序"""
        result = engine.execute_sql_query(GAME_CONFIG, "SELECT skill_name FROM 技能配置 ORDER BY LENGTH(skill_name)")
        assert result['success'] is True
        names = [row[0] for row in result['data'][1:]]
        # 最短名字在前
        assert len(names[0]) <= len(names[-1])

    def test_order_by_length_desc(self, engine):
        """ORDER BY LENGTH(col) DESC — 最长名字在前"""
        result = engine.execute_sql_query(GAME_CONFIG, "SELECT skill_name FROM 技能配置 ORDER BY LENGTH(skill_name) DESC")
        assert result['success'] is True
        names = [row[0] for row in result['data'][1:]]
        assert len(names[0]) >= len(names[-1])


class TestOrderByComplexExpressions:
    """ORDER BY 复杂表达式"""

    def test_order_by_coalesce(self, engine):
        """ORDER BY COALESCE(col, 0)"""
        result = engine.execute_sql_query(GAME_CONFIG, "SELECT skill_name, damage FROM 技能配置 ORDER BY COALESCE(damage, 0)")
        assert result['success'] is True

    def test_order_by_case_when(self, engine):
        """ORDER BY CASE WHEN ... END"""
        result = engine.execute_sql_query(
            GAME_CONFIG,
            "SELECT skill_name, damage FROM 技能配置 ORDER BY CASE WHEN damage > 100 THEN 1 ELSE 0 END DESC"
        )
        assert result['success'] is True
        # 高伤害在前
        damages = [row[1] for row in result['data'][1:] if row[1] is not None]
        if damages:
            # 至少第一个应该>100
            assert float(damages[0]) > 100

    def test_order_by_math_expression(self, engine):
        """ORDER BY col1 * col2"""
        result = engine.execute_sql_query(
            GAME_CONFIG,
            "SELECT skill_name, damage, cooldown FROM 技能配置 ORDER BY damage / cooldown DESC"
        )
        assert result['success'] is True

    def test_order_by_function_with_where(self, engine):
        """ORDER BY function + WHERE 组合"""
        result = engine.execute_sql_query(
            GAME_CONFIG,
            "SELECT skill_name FROM 技能配置 WHERE skill_type = '法师' ORDER BY LENGTH(skill_name)"
        )
        assert result['success'] is True
        # 结果数应>0且全部按长度递增
        names = [row[0] for row in result['data'][1:]]
        assert len(names) > 0
        for i in range(1, len(names)):
            assert len(names[i]) >= len(names[i - 1])
