"""
窗口函数测试 - ROW_NUMBER, RANK, DENSE_RANK, LAG, LEAD
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, ROW_NUMBER() OVER (ORDER BY damage DESC) as rn FROM 技能配置 LIMIT 5"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 5

    def test_row_number_order_by_asc(self, game_config):
        """ROW_NUMBER ORDER BY ASC: 升序编号"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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

    def test_ntile_basic(self, game_config):
        """NTILE: 将分组均匀分为N个桶"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, NTILE(3) OVER (ORDER BY damage DESC) as bucket FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 8
        # 8行分成3个桶，桶大小应该是 3, 3, 2
        bucket_counts = {}
        for r in rows:
            bucket_counts[r['bucket']] = bucket_counts.get(r['bucket'], 0) + 1
        assert bucket_counts[1] == 3  # 桶1有3行
        assert bucket_counts[2] == 3  # 桶2有3行
        assert bucket_counts[3] == 2  # 桶3有2行

    def test_nth_value(self, game_config):
        """NTH_VALUE: 返回窗口中第N行的值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, NTH_VALUE(damage, 2) OVER (ORDER BY damage DESC) as val FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 8

    def test_window_column_not_exists(self, game_config):
        """窗口函数引用不存在的列"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, ROW_NUMBER() OVER (ORDER BY nonexistent DESC) as rn FROM 技能配置"
        )
        assert result['success'] is False
        assert '不存在' in result['message']

    def test_window_no_partition_no_order(self, game_config):
        """ROW_NUMBER() OVER () — 无分区无排序"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
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


class TestLag:
    """LAG() 测试"""

    def test_lag_basic(self, game_config):
        """基本LAG: 获取前一行的伤害值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LAG(damage, 1) OVER (ORDER BY damage DESC) as prev_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 8
        # 第一行LAG应该是NULL（没有前一行）
        assert pd.isna(rows[0]['prev_damage']) or rows[0]['prev_damage'] is None
        # 第二行LAG应该是第一行的伤害(250)
        assert rows[1]['prev_damage'] == 250

    def test_lag_partition(self, game_config):
        """LAG with PARTITION BY: 每个职业内前一行"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, skill_type, damage, LAG(damage, 1) OVER (PARTITION BY skill_type ORDER BY damage DESC) as prev_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)

        # 法师分组的第一行LAG应该是NULL
        mages = [r for r in rows if r['skill_type'] == 'mage']
        mage_damages = sorted([r['damage'] for r in mages], reverse=True)
        assert len(mages) == 4
        # 检查法师分区内LAG正确
        for i, mage in enumerate(sorted(mages, key=lambda x: x['damage'], reverse=True)):
            if i == 0:
                assert pd.isna(mage['prev_damage']) or mage['prev_damage'] is None
            else:
                assert mage['prev_damage'] == mage_damages[i - 1]

    def test_lag_offset_2(self, game_config):
        """LAG with offset 2: 获取前两行的值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LAG(damage, 2) OVER (ORDER BY damage DESC) as prev2_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 前两行LAG应该是NULL
        assert pd.isna(rows[0]['prev2_damage']) or rows[0]['prev2_damage'] is None
        assert pd.isna(rows[1]['prev2_damage']) or rows[1]['prev2_damage'] is None
        # 第三行LAG应该是第一行的值
        assert rows[2]['prev2_damage'] == rows[0]['damage']

    def test_lag_default_value(self, game_config):
        """LAG with default value: 为NULL提供默认值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LAG(damage, 1, 0) OVER (ORDER BY damage DESC) as prev_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 第一行LAG应该使用默认值0
        assert rows[0]['prev_damage'] == 0

    def test_lag_no_order_by_error(self, game_config):
        """LAG without ORDER BY should error"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LAG(damage, 1) OVER () as prev_damage FROM 技能配置"
        )
        assert result['success'] is False
        assert '需要 ORDER BY' in result['message']


class TestLead:
    """LEAD() 测试"""

    def test_lead_basic(self, game_config):
        """基本LEAD: 获取后一行的伤害值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LEAD(damage, 1) OVER (ORDER BY damage DESC) as next_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 8
        # 最后一行LEAD应该是NULL（没有后一行）
        assert pd.isna(rows[-1]['next_damage']) or rows[-1]['next_damage'] is None
        # 第一行LEAD应该是第二行的伤害(220)
        assert rows[0]['next_damage'] == 220

    def test_lead_partition(self, game_config):
        """LEAD with PARTITION BY: 每个职业内后一行"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, skill_type, damage, LEAD(damage, 1) OVER (PARTITION BY skill_type ORDER BY damage DESC) as next_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)

        # 法师分组的最后一行LEAD应该是NULL
        mages = [r for r in rows if r['skill_type'] == 'mage']
        mage_damages = sorted([r['damage'] for r in mages], reverse=True)
        assert len(mages) == 4
        # 检查法师分区内LEAD正确
        sorted_mages = sorted(mages, key=lambda x: x['damage'], reverse=True)
        for i in range(len(sorted_mages)):
            if i == len(sorted_mages) - 1:
                assert pd.isna(sorted_mages[i]['next_damage']) or sorted_mages[i]['next_damage'] is None
            else:
                assert sorted_mages[i]['next_damage'] == mage_damages[i + 1]

    def test_lead_offset_2(self, game_config):
        """LEAD with offset 2: 获取后两行的值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LEAD(damage, 2) OVER (ORDER BY damage DESC) as next2_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 最后两行LEAD应该是NULL
        assert pd.isna(rows[-1]['next2_damage']) or rows[-1]['next2_damage'] is None
        assert pd.isna(rows[-2]['next2_damage']) or rows[-2]['next2_damage'] is None

    def test_lead_default_value(self, game_config):
        """LEAD with default value: 为NULL提供默认值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LEAD(damage, 1, 0) OVER (ORDER BY damage DESC) as next_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 最后一行LEAD应该使用默认值0
        assert rows[-1]['next_damage'] == 0

    def test_lead_no_order_by_error(self, game_config):
        """LEAD without ORDER BY should error"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LEAD(damage, 1) OVER () as next_damage FROM 技能配置"
        )
        assert result['success'] is False
        assert '需要 ORDER BY' in result['message']


class TestFirstValue:
    """FIRST_VALUE() 测试"""

    def test_first_value_basic(self, game_config):
        """基本FIRST_VALUE: 获取按伤害排序后第一行的伤害值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, FIRST_VALUE(damage) OVER (ORDER BY damage DESC) as first_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 8
        # 所有行都应该返回第一行的伤害值(250)
        assert all(r['first_damage'] == 250 for r in rows)

    def test_first_value_partition(self, game_config):
        """FIRST_VALUE with PARTITION BY: 每个职业内第一行的伤害值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, skill_type, damage, FIRST_VALUE(damage) OVER (PARTITION BY skill_type ORDER BY damage DESC) as first_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)

        # 法师分组：所有法师行应该返回法师最高的伤害值(250)
        mages = [r for r in rows if r['skill_type'] == 'mage']
        assert len(mages) == 4
        assert all(r['first_damage'] == 250 for r in mages)

        # 战士分组：所有战士行应该返回战士最高的伤害值(200)
        warriors = [r for r in rows if r['skill_type'] == 'warrior']
        assert len(warriors) == 2
        assert all(r['first_damage'] == 200 for r in warriors)

    def test_first_value_order_by_asc(self, game_config):
        """FIRST_VALUE ORDER BY ASC: 获取排序后第一行（最小值）"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, FIRST_VALUE(damage) OVER (ORDER BY damage ASC) as first_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 所有行都应该返回第一行的伤害值(0)
        assert all(r['first_damage'] == 0 for r in rows)

    def test_first_value_no_order_by_error(self, game_config):
        """FIRST_VALUE without ORDER BY should error"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, FIRST_VALUE(damage) OVER () as first_damage FROM 技能配置"
        )
        assert result['success'] is False
        assert '需要 ORDER BY' in result['message']


class TestLastValue:
    """LAST_VALUE() 测试"""

    def test_last_value_basic(self, game_config):
        """基本LAST_VALUE: 获取按伤害排序后最后一行的伤害值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LAST_VALUE(damage) OVER (ORDER BY damage DESC) as last_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 8
        # 所有行都应该返回最后一行的伤害值(0)
        assert all(r['last_damage'] == 0 for r in rows)

    def test_last_value_partition(self, game_config):
        """LAST_VALUE with PARTITION BY: 每个职业内最后一行的伤害值"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, skill_type, damage, LAST_VALUE(damage) OVER (PARTITION BY skill_type ORDER BY damage DESC) as last_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)

        # 法师分组：所有法师行应该返回法师最低的伤害值(150)
        mages = [r for r in rows if r['skill_type'] == 'mage']
        assert len(mages) == 4
        assert all(r['last_damage'] == 150 for r in mages)

        # 战士分组：所有战士行应该返回战士最低的伤害值(190)
        warriors = [r for r in rows if r['skill_type'] == 'warrior']
        assert len(warriors) == 2
        assert all(r['last_damage'] == 190 for r in warriors)

    def test_last_value_order_by_asc(self, game_config):
        """LAST_VALUE ORDER BY ASC: 获取排序后最后一行（最大值）"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LAST_VALUE(damage) OVER (ORDER BY damage ASC) as last_damage FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 所有行都应该返回最后一行的伤害值(250)
        assert all(r['last_damage'] == 250 for r in rows)

    def test_last_value_no_order_by_error(self, game_config):
        """LAST_VALUE without ORDER BY should error"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, damage, LAST_VALUE(damage) OVER () as last_damage FROM 技能配置"
        )
        assert result['success'] is False
        assert '需要 ORDER BY' in result['message']


class TestFirstLastValueCombo:
    """FIRST_VALUE 和 LAST_VALUE 组合测试"""

    def test_first_last_value_combo(self, game_config):
        """同时使用FIRST_VALUE和LAST_VALUE"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            game_config,
            "SELECT skill_name, skill_type, damage, FIRST_VALUE(damage) OVER (PARTITION BY skill_type ORDER BY damage DESC) as first_dmg, LAST_VALUE(damage) OVER (PARTITION BY skill_type ORDER BY damage DESC) as last_dmg FROM 技能配置"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)

        # 法师分组：first_dmg=250(最高), last_dmg=150(最低)
        mages = [r for r in rows if r['skill_type'] == 'mage']
        assert len(mages) == 4
        assert all(r['first_dmg'] == 250 for r in mages)
        assert all(r['last_dmg'] == 150 for r in mages)
