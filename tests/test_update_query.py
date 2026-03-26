"""测试 excel_update_query 功能"""
import os
import shutil
import pytest
import pandas as pd


@pytest.fixture
def game_config_copy(tmp_path):
    """创建游戏配置表的临时副本"""
    src = os.path.join(os.path.dirname(__file__), 'test_data', 'game_config.xlsx')
    dst = str(tmp_path / 'test_update.xlsx')
    shutil.copy(src, dst)
    return dst


class TestUpdateQueryBasic:
    """基础UPDATE功能测试"""

    def test_update_single_column_arithmetic(self, game_config_copy):
        """单列算术更新：法师伤害*1.1"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET damage = damage * 1.1 WHERE skill_type = '法师'"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 4
        assert len(result['changes']) == 4

        # 验证写回
        df = pd.read_excel(game_config_copy, sheet_name='技能配置', header=1)
        mage_damages = df[df['skill_type'] == '法师']['damage'].tolist()
        assert mage_damages == [110, 88, 220, 198]

    def test_update_multi_column(self, game_config_copy):
        """多列SET同时更新"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET damage = damage + 10, cost = cost * 0.9 WHERE skill_type = '战士'"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 2
        assert len(result['changes']) == 4  # 2 rows * 2 columns

        # 验证战士数据
        df = pd.read_excel(game_config_copy, sheet_name='技能配置', header=1)
        warriors = df[df['skill_type'] == '战士']
        assert warriors['damage'].tolist() == [130, 160]
        assert warriors['cost'].tolist() == [18, 45]

    def test_update_column_reference(self, game_config_copy):
        """SET使用列引用：damage = cooldown * 20"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET damage = cooldown * 20 WHERE skill_type = '辅助'"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 2

        df = pd.read_excel(game_config_copy, sheet_name='技能配置', header=1)
        supports = df[df['skill_type'] == '辅助']
        assert supports['damage'].tolist() == [200, 300]

    def test_update_constant_value(self, game_config_copy):
        """SET使用常量值"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET description = '已审核' WHERE level = 1"
        )
        assert result['success'] is True

        df = pd.read_excel(game_config_copy, sheet_name='技能配置', header=1)
        level1 = df[df['level'] == 1]
        assert all(d == '已审核' for d in level1['description'])

    def test_dry_run_no_modify(self, game_config_copy):
        """预览模式不修改文件"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET damage = 999 WHERE skill_type = '法师'",
            dry_run=True
        )
        assert result['success'] is True
        assert result['dry_run'] is True
        assert result['affected_rows'] > 0

        # 文件未被修改
        df = pd.read_excel(game_config_copy, sheet_name='技能配置', header=1)
        assert df[df['skill_type'] == '法师']['damage'].tolist() == [100, 80, 200, 180]

    def test_update_no_where(self, game_config_copy):
        """无WHERE条件更新所有行"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET description = '待审核'"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 10

    def test_update_no_match(self, game_config_copy):
        """WHERE不匹配任何行"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET damage = 999 WHERE skill_type = '不存在'"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 0
        assert result['changes'] == []

    def test_update_value_unchanged(self, game_config_copy):
        """值未变化时不写入"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET damage = damage WHERE skill_type = '法师'"
        )
        assert result['success'] is True
        assert len(result['changes']) == 0
        assert '无变化' in result['message']

    def test_select_rejected(self, game_config_copy):
        """SELECT语句被拒绝"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "SELECT * FROM 技能配置"
        )
        assert result['success'] is False
        assert '只支持UPDATE' in result['message']

    def test_wrong_table_name(self, game_config_copy):
        """错误的表名给出建议"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 不存在的表 SET damage = 1"
        )
        assert result['success'] is False
        assert '不存在' in result['message']

    def test_wrong_column_name(self, game_config_copy):
        """错误的列名给出建议"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET 不存在的列 = 1"
        )
        assert result['success'] is False
        assert '不存在' in result['message']

    def test_empty_file_path(self):
        """空文件路径"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query('', "UPDATE t SET a=1")
        assert result['success'] is False

    def test_empty_sql(self, game_config_copy):
        """空SQL"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(game_config_copy, '')
        assert result['success'] is False

    def test_nonexistent_file(self):
        """不存在的文件"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            '/tmp/nonexistent_file.xlsx',
            "UPDATE t SET a=1"
        )
        assert result['success'] is False
        assert '不存在' in result['message']
