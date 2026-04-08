"""OFFSET语法测试 — 验证分页查询支持"""
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
    """创建包含5条记录的测试文件"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        tmpfile = f.name

    with pd.ExcelWriter(tmpfile, engine='openpyxl') as writer:
        pd.DataFrame({
            '技能名称': ['火球', '冰箭', '雷电', '治愈', '毒雾'],
            '伤害': [100, 80, 120, 0, 60],
        }).to_excel(writer, sheet_name='技能配置', index=False)

    yield tmpfile
    os.unlink(tmpfile)


def get_data(result):
    """从结果中提取数据行（跳过header）"""
    data = result.get('data', [])
    if not data:
        return []
    return [row for row in data[1:] if any(cell is not None for cell in row)]


class TestOffset:
    """OFFSET分页功能测试"""

    def test_offset_only(self, engine, test_file):
        """OFFSET跳过前N行"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT 技能名称, 伤害 FROM 技能配置 ORDER BY 伤害 DESC OFFSET 2'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 3
        # 伤害降序: 雷电120, 火球100, 冰箭80, 毒雾60, 治愈0
        # OFFSET 2 跳过雷电和火球
        assert rows[0][1] == 80  # 冰箭

    def test_offset_with_limit(self, engine, test_file):
        """OFFSET + LIMIT 分页"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT 技能名称, 伤害 FROM 技能配置 ORDER BY 伤害 DESC LIMIT 2 OFFSET 1'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 2
        # 跳过雷电120，取火球100和冰箭80
        assert rows[0][0] == '火球'
        assert rows[0][1] == 100

    def test_offset_zero(self, engine, test_file):
        """OFFSET 0 等于无偏移"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT 技能名称 FROM 技能配置 ORDER BY 伤害 DESC LIMIT 2 OFFSET 0'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 2
        assert rows[0][0] == '雷电'

    def test_offset_exceeds_total(self, engine, test_file):
        """OFFSET超过总行数返回空"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT 技能名称 FROM 技能配置 OFFSET 999'
        )
        assert result['success']
        rows = get_data(result)
        assert len(rows) == 0

    def test_offset_with_where(self, engine, test_file):
        """OFFSET + WHERE组合"""
        result = engine.execute_sql_query(
            test_file,
            'SELECT 技能名称, 伤害 FROM 技能配置 WHERE 伤害 > 50 ORDER BY 伤害 DESC OFFSET 1'
        )
        assert result['success']
        rows = get_data(result)
        # 伤害>50: 雷电120, 火球100, 冰箭80, 毒雾60
        # OFFSET 1 跳过雷电
        assert len(rows) == 3
        assert rows[0][1] == 100  # 火球
