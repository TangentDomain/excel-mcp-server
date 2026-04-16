"""UNION / UNION ALL 功能测试"""
import pytest
import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

TEST_FILE = os.path.join(PROJECT_ROOT, "tests", "test_data", "game_config.xlsx")


@pytest.fixture
def engine():
    """创建SQL查询引擎实例"""
    return AdvancedSQLQueryEngine()


def _data_rows(result):
    """从结果中提取数据行（跳过表头行，include_headers=True时第一行是列名）"""
    if not result['success'] or not result['data']:
        return []
    # 第一行是列名（字符串），后续行是数据
    return result['data'][1:] if result['data'] and isinstance(result['data'][0][0], str) and not any(isinstance(v, (int, float)) for v in result['data'][0]) else result['data']


class TestUnionAll:
    """UNION ALL 基本功能测试"""

    def test_union_all_basic(self, engine):
        """基本 UNION ALL：合并两个不同类型的查询结果"""
        sql = """
        SELECT skill_name, damage FROM 技能配置 WHERE 技能类型='法师'
        UNION ALL
        SELECT skill_name, damage FROM 技能配置 WHERE 技能类型='战士'
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        # 法师4个 + 战士2个 = 6行数据
        assert len(rows) == 6
        names = [row[0] for row in rows]
        assert '火球术' in names
        assert '斩击' in names

    def test_union_all_with_limit(self, engine):
        """UNION ALL + LIMIT"""
        sql = """
        SELECT skill_name FROM 技能配置 WHERE 技能类型='法师'
        UNION ALL
        SELECT skill_name FROM 技能配置 WHERE 技能类型='战士'
        LIMIT 3
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) == 3

    def test_union_all_with_order_by(self, engine):
        """UNION ALL + ORDER BY（数值列排序）"""
        sql = """
        SELECT skill_name, damage FROM 技能配置 WHERE 技能类型='法师'
        UNION ALL
        SELECT skill_name, damage FROM 技能配置 WHERE 技能类型='战士'
        ORDER BY damage DESC
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        damages = [float(row[1]) for row in rows if row[1] != '']
        assert damages == sorted(damages, reverse=True)

    def test_union_all_with_order_by_and_limit(self, engine):
        """UNION ALL + ORDER BY + LIMIT"""
        sql = """
        SELECT skill_name, damage FROM 技能配置 WHERE 技能类型='法师'
        UNION ALL
        SELECT skill_name, damage FROM 技能配置 WHERE 技能类型='战士'
        ORDER BY damage DESC
        LIMIT 2
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) == 2
        # 第一行应该是最高伤害
        d1 = float(rows[0][1]) if rows[0][1] != '' else 0
        d2 = float(rows[1][1]) if rows[1][1] != '' else 0
        assert d1 >= d2

    def test_union_all_three_selects(self, engine):
        """三表 UNION ALL（链式 UNION）"""
        sql = """
        SELECT skill_name FROM 技能配置 WHERE 技能类型='法师'
        UNION ALL
        SELECT skill_name FROM 技能配置 WHERE 技能类型='战士'
        UNION ALL
        SELECT skill_name FROM 技能配置 WHERE 技能类型='刺客'
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        # 法师4 + 战士2 + 刺客2 = 8
        assert len(rows) == 8


class TestUnionDistinct:
    """UNION（去重）测试"""

    def test_union_dedup(self, engine):
        """UNION 自动去重"""
        sql = """
        SELECT 技能类型 FROM 技能配置 WHERE 等级=1
        UNION
        SELECT 技能类型 FROM 技能配置 WHERE 等级=5
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        types = [row[0] for row in rows]
        # 去重后应该有4种类型
        assert '法师' in types
        assert '战士' in types
        assert '刺客' in types
        assert '辅助' in types
        # 不应该有重复
        assert len(types) == len(set(types))

    def test_union_all_no_dedup(self, engine):
        """UNION ALL 不去重"""
        sql = """
        SELECT 技能类型 FROM 技能配置 WHERE 等级=1
        UNION ALL
        SELECT 技能类型 FROM 技能配置 WHERE 等级=5
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        # 等级1有5行 + 等级5有5行 = 10行（可能有重复类型）
        assert len(rows) == 10

    def test_union_vs_union_all(self, engine):
        """对比 UNION 和 UNION ALL 的行数差异"""
        # UNION ALL: 保留所有行
        sql_all = "SELECT 技能类型 FROM 技能配置 WHERE 等级=1 UNION ALL SELECT 技能类型 FROM 技能配置 WHERE 等级=1"
        result_all = engine.execute_sql_query(TEST_FILE, sql_all)
        # UNION: 去重
        sql_distinct = "SELECT 技能类型 FROM 技能配置 WHERE 等级=1 UNION SELECT 技能类型 FROM 技能配置 WHERE 等级=1"
        result_distinct = engine.execute_sql_query(TEST_FILE, sql_distinct)
        assert result_all['success'] is True
        assert result_distinct['success'] is True
        # UNION ALL 应该有更多行（因为有重复）
        assert len(_data_rows(result_all)) >= len(_data_rows(result_distinct))


class TestUnionEdgeCases:
    """UNION 边界情况测试"""

    def test_union_empty_result_from_one_side(self, engine):
        """一侧查询返回空结果"""
        sql = """
        SELECT skill_name FROM 技能配置 WHERE 技能类型='不存在'
        UNION ALL
        SELECT skill_name FROM 技能配置 WHERE 技能类型='法师'
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) == 4  # 只有法师的4个

    def test_union_different_columns(self, engine):
        """两侧 SELECT 不同列数应报错（R56修复：防止静默截断数据）"""
        sql = """
        SELECT skill_name, damage FROM 技能配置 WHERE 技能类型='法师'
        UNION ALL
        SELECT skill_name FROM 技能配置 WHERE 技能类型='战士'
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        # R56: 不同列数UNION应返回错误而非静默截断
        assert result['success'] is False
        msg = result.get('message', '')
        assert '列数' in msg or 'column' in msg.lower() or 'UNION' in msg, \
            f"错误信息不明确: {msg}"

    def test_union_with_aggregation(self, engine):
        """UNION 中包含聚合查询"""
        sql = """
        SELECT 技能类型, AVG(伤害) as avg_dmg FROM 技能配置 WHERE 技能类型='法师'
        UNION ALL
        SELECT 技能类型, AVG(伤害) as avg_dmg FROM 技能配置 WHERE 技能类型='战士'
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        # 聚合结果各1行 + 无TOTAL行（UNION不触发TOTAL）= 2行数据
        assert len(rows) == 2

    def test_union_with_where_and_order(self, engine):
        """UNION + WHERE条件 + ORDER BY排序"""
        sql = """
        SELECT skill_name, damage, 技能类型 FROM 技能配置 WHERE 伤害 > 100
        UNION ALL
        SELECT skill_name, damage, 技能类型 FROM 技能配置 WHERE 伤害 <= 100 AND 技能类型='战士'
        ORDER BY damage DESC
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        damages = [float(row[1]) for row in rows if row[1] != '']
        assert damages == sorted(damages, reverse=True)

    def test_union_cross_sheet(self, engine):
        """跨工作表 UNION（同文件内不同sheet）"""
        sql = """
        SELECT skill_name FROM 技能配置 WHERE 技能类型='法师'
        UNION ALL
        SELECT equip_name FROM 装备配置 WHERE quality=5
        """
        result = engine.execute_sql_query(TEST_FILE, sql)
        assert result['success'] is True
        rows = _data_rows(result)
        # 法师4个 + 传说品质装备
        assert len(rows) >= 4
        names = [row[0] for row in rows]
        assert '火球术' in names
