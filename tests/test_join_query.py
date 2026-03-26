"""JOIN查询测试 - 验证INNER JOIN和LEFT JOIN功能"""
import pytest
import os

TEST_FILE = os.path.join(os.path.dirname(__file__), 'test_data', 'join_test.xlsx')


def _rows(result):
    """返回数据行（不含表头，第一行总是表头）"""
    data = result.get('data', [])
    return data[1:] if data and len(data) > 0 else data


class TestInnerJoin:
    """INNER JOIN 基本功能测试"""

    def test_inner_join_basic(self):
        """基本INNER JOIN: 技能表 JOIN 装备表 ON equip_id"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT a.skill_name, b.equip_name FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id"
        )
        assert result['success'] is True
        rows = _rows(result)
        assert len(rows) == 5  # 5个技能都有匹配装备

    def test_inner_join_with_where(self):
        """INNER JOIN + WHERE: 查找攻击类技能的装备"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT a.skill_name, b.equip_name, a.damage FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id WHERE a.skill_type = 'attack'"
        )
        assert result['success'] is True
        rows = _rows(result)
        skill_names = [r[0] for r in rows]
        assert 'fireball' in skill_names
        assert 'slash' in skill_names
        assert 'thunder' in skill_names

    def test_inner_join_with_order_by(self):
        """INNER JOIN + ORDER BY: 按伤害排序"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT a.skill_name, a.damage, b.equip_name FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id ORDER BY a.damage DESC"
        )
        assert result['success'] is True
        rows = _rows(result)
        damages = [int(row[1]) for row in rows]
        assert damages == sorted(damages, reverse=True)

    def test_inner_join_select_star(self):
        """INNER JOIN + SELECT *: 返回所有列"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT * FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id"
        )
        assert result['success'] is True
        rows = _rows(result)
        assert len(rows) > 0

    def test_inner_join_no_match(self):
        """INNER JOIN无匹配: 返回空结果（只有表头）"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT a.skill_name, b.equip_name FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id WHERE a.skill_type = 'nonexistent'"
        )
        assert result['success'] is True
        rows = _rows(result)
        assert len(rows) == 0


class TestLeftJoin:
    """LEFT JOIN 测试"""

    def test_left_join_keeps_all_left(self):
        """LEFT JOIN: 保留左表所有行"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT a.skill_name, b.equip_name FROM 技能表 a LEFT JOIN 装备表 b ON a.equip_id = b.equip_id"
        )
        assert result['success'] is True
        rows = _rows(result)
        assert len(rows) == 5  # 所有5个技能


class TestJoinWithAggregate:
    """JOIN + 聚合函数测试"""

    def test_join_with_group_by(self):
        """JOIN + GROUP BY: 统计每种装备关联的技能数"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT b.equip_name, COUNT(a.skill_name) as skill_count, AVG(a.damage) as avg_dmg FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id GROUP BY b.equip_name"
        )
        assert result['success'] is True
        rows = _rows(result)
        assert len(rows) > 0


class TestJoinEdgeCases:
    """JOIN边界情况测试"""

    def test_join_nonexistent_table(self):
        """JOIN不存在的表: 应该报错"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT * FROM 技能表 a JOIN 不存在的表 b ON a.equip_id = b.equip_id"
        )
        assert result['success'] is False
        assert '不存在' in result.get('message', '')

    def test_join_missing_on(self):
        """JOIN缺少ON条件: 应该报错"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT * FROM 技能表 a JOIN 装备表 b"
        )
        assert result['success'] is False

    def test_join_on_nonexistent_column(self):
        """JOIN ON条件引用不存在的列: 应该报错"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT * FROM 技能表 a JOIN 装备表 b ON a.not_exist = b.equip_id"
        )
        assert result['success'] is False

    def test_join_with_limit(self):
        """JOIN + LIMIT"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT a.skill_name, b.equip_name FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id LIMIT 2"
        )
        assert result['success'] is True
        rows = _rows(result)
        assert len(rows) <= 2

    def test_join_qualified_column_in_where(self):
        """WHERE中使用限定列名"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT a.skill_name, b.equip_name FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id WHERE b.quality = 'legendary'"
        )
        assert result['success'] is True
        rows = _rows(result)
        # E001(烈焰法杖-传说): 火球术+治愈术, E003(斩龙剑-传说): 斩击+雷击 = 4行
        assert len(rows) == 4

    def test_three_table_join(self):
        """三表JOIN: 语法支持"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT a.skill_name FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id JOIN 怪物表 c ON b.equip_id = c.monster_name"
        )
        # 语法应该被解析（不报不支持JOIN）
        assert '不支持JOIN' not in result.get('message', '')

    def test_join_column_disambiguation(self):
        """JOIN列名消歧义: 两个表都有equip_id列"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT a.skill_name, a.equip_id, b.equip_name FROM 技能表 a JOIN 装备表 b ON a.equip_id = b.equip_id LIMIT 1"
        )
        assert result['success'] is True

    def test_join_without_alias(self):
        """JOIN不带表别名"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            file_path=TEST_FILE,
            sql="SELECT 技能表.skill_name, 装备表.equip_name FROM 技能表 JOIN 装备表 ON 技能表.equip_id = 装备表.equip_id LIMIT 2"
        )
        assert result['success'] is True
        rows = _rows(result)
        assert len(rows) <= 2
