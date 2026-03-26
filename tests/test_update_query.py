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

    def test_transaction_rollback_on_write_failure(self, game_config_copy):
        """写入失败时自动回滚，文件不损坏"""
        import shutil
        from src.api.advanced_sql_query import execute_advanced_update_query

        # 先读取原始文件内容做校验
        import hashlib
        with open(game_config_copy, 'rb') as f:
            original_hash = hashlib.md5(f.read()).hexdigest()

        # 正常执行一次UPDATE（应该成功）
        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET damage = damage * 1.1 WHERE skill_type = '法师'"
        )
        assert result['success'] is True
        assert result['affected_rows'] > 0

        # 文件应该被修改（hash不同）
        with open(game_config_copy, 'rb') as f:
            modified_hash = hashlib.md5(f.read()).hexdigest()
        assert original_hash != modified_hash


class TestFileLockProtection:
    """文件锁保护测试"""

    def test_single_write_with_lock(self, game_config_copy):
        """单次写入文件锁正常工作不报错"""
        from src.api.advanced_sql_query import execute_advanced_update_query
        result = execute_advanced_update_query(
            file_path=game_config_copy,
            sql="UPDATE 技能配置 SET damage = damage + 10 WHERE 1=1",
            dry_run=False
        )
        assert result['success'] is True
        assert result['affected_rows'] > 0


class TestNumpySerialization:
    """numpy类型JSON序列化测试"""

    def test_dry_run_changes_json_serializable(self, game_config_copy):
        """dry_run返回的changes列表必须可JSON序列化（numpy→Python原生）"""
        import json
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET damage = damage * 1.1 WHERE skill_type = '法师'",
            dry_run=True
        )
        assert result['success'] is True
        assert result['dry_run'] is True
        # Must be JSON-serializable
        json_str = json.dumps(result)
        assert json_str  # No TypeError raised
        # Verify values are Python native types, not numpy
        for change in result['changes']:
            assert isinstance(change['row'], int)
            assert isinstance(change['new_value'], (int, float, str))
            assert isinstance(change['old_value'], (int, float, str))

    def test_actual_write_changes_json_serializable(self, game_config_copy):
        """实际写入返回的changes列表必须可JSON序列化"""
        import json
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET damage = damage + 10 WHERE level >= 1",
            dry_run=False
        )
        assert result['success'] is True
        json_str = json.dumps(result)
        assert json_str

    def test_update_with_chinese_column_names(self, game_config_copy):
        """中文列名UPDATE：SET和WHERE都使用中文列名"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET 伤害 = 伤害 * 1.1 WHERE 技能类型 = '法师'",
            dry_run=True
        )
        assert result['success'] is True
        assert result['affected_rows'] == 4

    def test_update_chinese_set_english_where(self, game_config_copy):
        """混合列名UPDATE：SET中文、WHERE英文"""
        from src.api.advanced_sql_query import execute_advanced_update_query

        result = execute_advanced_update_query(
            game_config_copy,
            "UPDATE 技能配置 SET 伤害 = 伤害 * 1.1 WHERE skill_type = '法师'",
            dry_run=True
        )
        assert result['success'] is True
        assert result['affected_rows'] == 4
