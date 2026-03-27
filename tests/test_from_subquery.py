"""FROM子查询测试 — REQ-028

支持 FROM (SELECT ...) AS alias 语法，允许将子查询结果作为虚拟表使用。
"""
import pytest
import pandas as pd
import tempfile
import os

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


@pytest.fixture
def engine():
    return AdvancedSQLQueryEngine()


@pytest.fixture
def test_file():
    """创建包含两个工作表的测试Excel文件"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        tmpfile = f.name

    with pd.ExcelWriter(tmpfile, engine='openpyxl') as writer:
        pd.DataFrame({
            '技能名称': ['火球', '冰箭', '雷电', '治愈', '毒雾'],
            '伤害': [100, 80, 120, 0, 60],
            '技能类型': ['法师', '法师', '法师', '牧师', '刺客'],
            '等级': [1, 2, 3, 1, 2]
        }).to_excel(writer, sheet_name='技能配置', index=False)
        pd.DataFrame({
            '技能名称': ['火球', '冰箭', '雷电', '治愈', '毒雾'],
            'MP消耗': [30, 20, 50, 15, 25],
            '冷却时间': [5, 3, 8, 2, 4]
        }).to_excel(writer, sheet_name='技能数值', index=False)

    yield tmpfile
    os.unlink(tmpfile)


def get_data(result):
    """从execute_sql_query结果中提取数据行（跳过header行）"""
    data = result['data']
    if not data:
        return []
    # data[0] 是header行，data[1:] 是数据行
    # 过滤掉空行（openpyxl可能产生）
    rows = [row for row in data[1:] if any(cell is not None for cell in row)]
    return rows


class TestBasicFromSubquery:
    """基础FROM子查询功能"""

    def test_basic_from_subquery(self, engine, test_file):
        """基础FROM子查询：过滤后作为虚拟表"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT * FROM (SELECT 技能名称, 伤害 FROM 技能配置 WHERE 伤害 > 80) AS t'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 2
        names = {row[0] for row in rows}
        assert names == {'火球', '雷电'}

    def test_from_subquery_with_outer_where(self, engine, test_file):
        """FROM子查询 + 外层WHERE过滤"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT * FROM (SELECT 技能名称, 伤害 FROM 技能配置) AS t WHERE 伤害 > 80'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 2

    def test_from_subquery_specific_columns(self, engine, test_file):
        """FROM子查询 + 外层选择特定列"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT t.技能名称 FROM (SELECT 技能名称, 伤害 FROM 技能配置 WHERE 伤害 > 80) AS t'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 2
        assert all(len(row) == 1 for row in rows)


class TestFromSubqueryWithAggregation:
    """FROM子查询包裹聚合"""

    def test_from_subquery_wraps_group_by(self, engine, test_file):
        """FROM子查询包裹GROUP BY聚合"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT * FROM (SELECT 技能类型, AVG(伤害) as avg_dmg FROM 技能配置 GROUP BY 技能类型) AS stats WHERE avg_dmg > 80'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 1
        assert rows[0][0] == '法师'

    def test_from_subquery_with_count(self, engine, test_file):
        """FROM子查询 + COUNT聚合"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT COUNT(*) as cnt FROM (SELECT 技能名称 FROM 技能配置 WHERE 伤害 > 80) AS t'
        )
        assert result['success']
        rows = get_data(result)
        assert int(rows[0][0]) == 2


class TestFromSubqueryWithOrderLimit:
    """FROM子查询 + ORDER BY/LIMIT"""

    def test_from_subquery_order_limit(self, engine, test_file):
        """FROM子查询内ORDER BY + LIMIT，外层再过滤"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT * FROM (SELECT 技能名称, 伤害 FROM 技能配置 ORDER BY 伤害 DESC LIMIT 3) AS top_skills WHERE 伤害 > 90'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 2
        # 验证排序保持（雷电120 > 火球100）
        assert rows[0][1] > rows[1][1]


class TestFromSubqueryWithJoin:
    """FROM子查询 + JOIN"""

    def test_from_subquery_join(self, engine, test_file):
        """FROM子查询结果与另一工作表JOIN"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT t.技能名称, v.MP消耗 FROM (SELECT 技能名称, 伤害 FROM 技能配置 WHERE 伤害 > 80) AS t JOIN 技能数值 v ON t.技能名称 = v.技能名称'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 2


class TestFromSubqueryEdgeCases:
    """FROM子查询边界情况"""

    def test_from_subquery_no_alias(self, engine, test_file):
        """FROM子查询无别名时使用默认别名"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT * FROM (SELECT 技能名称, 伤害 FROM 技能配置 WHERE 伤害 > 80)'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 2

    def test_from_subquery_empty_result(self, engine, test_file):
        """FROM子查询返回空结果"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT * FROM (SELECT 技能名称 FROM 技能配置 WHERE 伤害 > 999) AS empty'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 0

    def test_from_subquery_with_distinct(self, engine, test_file):
        """FROM子查询 + DISTINCT"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT DISTINCT 技能类型 FROM (SELECT 技能类型 FROM 技能配置) AS t'
        )
        assert result['success']
        rows = get_data(result)
        types = {row[0] for row in rows}
        assert types == {'法师', '牧师', '刺客'}

    def test_from_subquery_referencing_nonexistent_table(self, engine, test_file):
        """FROM子查询引用不存在的表应报错"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT * FROM (SELECT 技能名称 FROM 不存在的表) AS t'
        )
        assert not result['success']
        assert '不存在' in result['message']

    def test_nested_from_subquery_rejected(self, engine, test_file):
        """嵌套FROM子查询应被拒绝"""
        result = engine.execute_sql_query(
            test_file,
            "SELECT * FROM (SELECT * FROM (SELECT 技能名称 FROM 技能配置) AS inner_t) AS outer_t"
        )
        assert not result['success']
        assert '嵌套' in result['message']
