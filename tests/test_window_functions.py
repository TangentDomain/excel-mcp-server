"""
窗口函数测试 - ROW_NUMBER, RANK, DENSE_RANK
"""
import os
import pytest
import pandas as pd


@pytest.fixture
def game_config(tmp_path):
    """游戏配置表 - 用于窗口函数测试"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '技能配置'
    # 双行表头
    ws.append(['技能名称', 'skill_name', '技能类型', 'skill_type', '伤害', 'damage', '冷却', 'cooldown', '等级', 'level'])
    ws.append(['火球术', 'fireball', '法师', 'mage', 250, 250, 5, 5, 10, 10])
    ws.append(['冰冻术', 'ice', '法师', 'mage', 180, 180, 3, 3, 8, 8])
    ws.append(['斩击', 'slash', '战士', 'warrior', 200, 200, 2, 2, 5, 5])
    ws.append(['治疗术', 'heal', '牧师', 'priest', 0, 0, 8, 8, 3, 3])
    ws.append(['火墙', 'firewall', '法师', 'mage', 220, 220, 6, 6, 10, 10])
    ws.append(['冰锥', 'ice_spike', '法师', 'mage', 150, 150, 4, 4, 7, 7])
    ws.append(['旋风斩', 'whirlwind', '战士', 'warrior', 190, 190, 3, 3, 5, 5])
    ws.append(['圣光术', 'holy', '牧师', 'priest', 0, 0, 10, 10, 3, 3])
    path = str(tmp_path / 'window_test.xlsx')
    wb.save(path)
    return path


@pytest.fixture
def game_config_with_dup(tmp_path):
    """含重复伤害值的配置表 - 用于RANK/DENSE_RANK对比"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '技能配置'
    ws.append(['技能名称', 'skill_name', '技能类型', 'skill_type', '伤害', 'damage', '冷却', 'cooldown', '等级', 'level'])
    ws.append(['火球术', 'fireball', '法师', 'mage', 250, 250, 5, 5, 10, 10])
    ws.append(['雷击', 'lightning', '法师', 'mage', 250, 250, 7, 7, 10, 10])  # 与火球术同伤害
    ws.append(['火墙', 'firewall', '法师', 'mage', 220, 220, 6, 6, 10, 10])
    ws.append(['斩击', 'slash', '战士', 'warrior', 200, 200, 2, 2, 5, 5])
    ws.append(['冰矛', 'ice_spear', '法师', 'mage', 200, 200, 3, 3, 8, 8])   # 与斩击同伤害
    ws.append(['旋风斩', 'whirlwind', '战士', 'warrior', 190, 190, 3, 3, 5, 5])
    ws.append(['冰冻术', 'ice', '法师', 'mage', 180, 180, 3, 3, 8, 8])
    ws.append(['冰锥', 'ice_spike', '法师', 'mage', 150, 150, 4, 4, 7, 7])
    ws.append(['治疗术', 'heal', '牧师', 'priest', 0, 0, 8, 8, 3, 3])
    ws.append(['圣光术', 'holy', '牧师', 'priest', 0, 0, 10, 10, 3, 3])
    path = str(tmp_path / 'window_dup_test.xlsx')
    wb.save(path)
    return path


def _get_rows(result):
    """提取数据行（返回dict列表，跳过表头行）"""
    if not result.get('success'):
        return []
    data = result.get('data', [])
    if not data or not isinstance(data[0], list):
        return []
    headers = data[0]
    return [dict(zip(headers, row)) for row in data[1:]]


class TestRowNumber:
    """ROW_NUMBER() 测试"""

    def test_row_number_basic(self, game_config):
        """基本ROW_NUMBER: 按伤害降序编号"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, ROW_NUMBER() OVER (ORDER BY damage DESC) as rn FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 8
        # 第一行应该是伤害最高的
        assert rows[0]['rn'] == 1
        assert rows[0]['damage'] == 250
        assert rows[-1]['rn'] == 8

    def test_row_number_partition(self, game_config):
        """ROW_NUMBER with PARTITION BY: 每个职业内按伤害排名"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, skill_type, damage, ROW_NUMBER() OVER (PARTITION BY skill_type ORDER BY damage DESC) as rn FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 8

        mages = [r for r in rows if r['skill_type'] == 'mage']
        assert len(mages) == 4
        mage_ranks = [r['rn'] for r in mages]
        assert sorted(mage_ranks) == [1, 2, 3, 4]
        assert mages[0]['skill_name'] == 'fireball'
        assert mages[0]['rn'] == 1

        warriors = [r for r in rows if r['skill_type'] == 'warrior']
        assert len(warriors) == 2
        warrior_ranks = [r['rn'] for r in warriors]
        assert sorted(warrior_ranks) == [1, 2]

    def test_row_number_with_limit(self, game_config):
        """ROW_NUMBER + LIMIT"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, ROW_NUMBER() OVER (ORDER BY damage DESC) as rn FROM 技能配置 LIMIT 5"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 5

    def test_row_number_order_by_asc(self, game_config):
        """ROW_NUMBER ORDER BY ASC: 升序编号"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, cooldown, ROW_NUMBER() OVER (ORDER BY cooldown ASC) as rn FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 找到cooldown最小的行，rn应该是1
        min_cd_rows = [r for r in rows if r['rn'] == 1]
        assert len(min_cd_rows) == 1
        assert min_cd_rows[0]['cooldown'] == 2  # 斩击冷却最短


class TestRank:
    """RANK() 测试"""

    def test_rank_basic(self, game_config_with_dup):
        """基本RANK: 相同伤害相同排名，跳过"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config_with_dup,
            "SELECT skill_name, damage, RANK() OVER (ORDER BY damage DESC) as r FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 10

        top_two = [r for r in rows if r['damage'] == 250]
        assert len(top_two) == 2
        assert all(r['r'] == 1 for r in top_two)
        # 下一个应该rank=3（跳过2）
        assert rows[2]['r'] == 3

    def test_rank_partition(self, game_config_with_dup):
        """RANK with PARTITION BY"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config_with_dup,
            "SELECT skill_name, skill_type, damage, RANK() OVER (PARTITION BY skill_type ORDER BY damage DESC) as r FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)

        mages = [r for r in rows if r['skill_type'] == 'mage']
        # 法师伤害: 250,250,220,200,180,150
        # RANK: 1,1,3,4,5,6
        mage_ranks = [r['r'] for r in mages]
        assert mage_ranks.count(1) == 2  # 两个250并列第1


class TestDenseRank:
    """DENSE_RANK() 测试"""

    def test_dense_rank_basic(self, game_config_with_dup):
        """基本DENSE_RANK: 相同伤害相同排名，不跳过"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config_with_dup,
            "SELECT skill_name, damage, DENSE_RANK() OVER (ORDER BY damage DESC) as dr FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)

        top_two = [r for r in rows if r['damage'] == 250]
        assert all(r['dr'] == 1 for r in top_two)
        assert rows[2]['dr'] == 2  # 220 → dense_rank=2（不跳过）

    def test_dense_rank_multiple_ties(self, game_config_with_dup):
        """DENSE_RANK 多组并列"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config_with_dup,
            "SELECT skill_name, damage, DENSE_RANK() OVER (ORDER BY damage DESC) as dr FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 10

        # 伤害排序: 250,250,220,200,200,190,180,150,0,0
        # DENSE_RANK: 1,1,2,3,3,4,5,6,7,7
        drs = [r['dr'] for r in rows]
        assert drs[0] == 1
        assert drs[1] == 1
        assert drs[2] == 2  # 220
        assert drs[3] == 3  # 200
        assert drs[4] == 3  # 200
        assert drs[5] == 4  # 190


class TestWindowEdgeCases:
    """窗口函数边界情况"""

    def test_window_without_order_by(self, game_config):
        """无ORDER BY的ROW_NUMBER: 按原始行顺序"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, ROW_NUMBER() OVER () as rn FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 8
        ranks = [r['rn'] for r in rows]
        assert sorted(ranks) == list(range(1, 9))

    def test_window_with_where(self, game_config):
        """窗口函数 + WHERE: 先过滤再计算窗口"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, ROW_NUMBER() OVER (ORDER BY damage DESC) as rn FROM 技能配置 WHERE skill_type = 'mage'"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 只应该有法师的4个技能
        assert len(rows) == 4
        assert all(r['rn'] <= 4 for r in rows)

    def test_window_with_group_by(self, game_config):
        """窗口函数 + GROUP BY: 在聚合结果上计算排名"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_type, AVG(damage) as avg_dmg, RANK() OVER (ORDER BY AVG(damage) DESC) as r FROM 技能配置 GROUP BY skill_type"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 3个职业 + TOTAL行
        assert len(rows) >= 3

        mage_row = [r for r in rows if r['skill_type'] == 'mage']
        assert len(mage_row) == 1
        assert mage_row[0]['r'] == 1  # 法师平均伤害最高

    def test_unsupported_window_function(self, game_config):
        """不支持的窗口函数应该报错"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, NTILE(3) OVER (ORDER BY damage DESC) as bucket FROM 技能配置"
        )
        assert result['success'] is False
        assert '不支持' in result['message']

    def test_window_column_not_exists(self, game_config):
        """窗口函数引用不存在的列"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, ROW_NUMBER() OVER (ORDER BY nonexistent DESC) as rn FROM 技能配置"
        )
        assert result['success'] is False
        assert '不存在' in result['message']

    def test_window_no_partition_no_order(self, game_config):
        """ROW_NUMBER() OVER () — 无分区无排序"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, ROW_NUMBER() OVER () as rn FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        ranks = [r['rn'] for r in rows]
        assert sorted(ranks) == list(range(1, 9))

    def test_rank_dense_rank_comparison(self, game_config_with_dup):
        """RANK vs DENSE_RANK 对比"""
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config_with_dup,
            "SELECT skill_name, damage, RANK() OVER (ORDER BY damage DESC) as r, DENSE_RANK() OVER (ORDER BY damage DESC) as dr FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)

        # 第1-2行: rank=1, dense_rank=1
        assert rows[0]['r'] == 1 and rows[0]['dr'] == 1
        assert rows[1]['r'] == 1 and rows[1]['dr'] == 1
        # 第3行: rank=3 (跳过2), dense_rank=2 (不跳过)
        assert rows[2]['r'] == 3
        assert rows[2]['dr'] == 2
