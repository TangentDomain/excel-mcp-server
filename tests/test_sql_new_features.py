"""
新增SQL功能测试

测试内容:
- 7个新窗口函数: LAG, LEAD, FIRST_VALUE, LAST_VALUE, NTILE, PERCENT_RANK, CUME_DIST
- GROUP_CONCAT聚合函数
- UPDATE中窗口函数支持
"""

import pytest
import tempfile
import pandas as pd
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


@pytest.fixture
def loot_file():
    """掉落表测试文件"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        df = pd.DataFrame({
            'LootID': [92000011]*6 + [92000111]*6 + [92000112]*3,
            'PropID': [10131, 10132, 10133, 20131, 20132, 20133,
                       10131, 10132, 10133, 20131, 20132, 20133,
                       30131, 30132, 30133],
            'PropType': ['[类型]主武器']*3 + ['[类型]护手']*3 +
                        ['[类型]主武器']*3 + ['[类型]护手']*3 +
                        ['[类型]碎片']*3,
            'Quality': ['传奇', '史诗', '稀有'] * 4 + ['普通'] * 3,
            'Level': [50, 45, 40, 50, 45, 40, 55, 50, 45, 55, 50, 45, 30, 30, 30],
        })
        df.to_excel(tmp.name, index=False, sheet_name="LootList")
        yield tmp.name


# ==================== LAG / LEAD ====================

class TestLagLead:
    def test_lag_basic(self, loot_file):
        """LAG取前一行值"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT PropID, LAG(PropID, 1) OVER (ORDER BY PropID) AS prev_id FROM LootList"
        )
        assert result['success'] is True
        data = result['data']
        # 第一行的LAG应该为None
        assert data[1][1] is None or str(data[1][1]) == 'None' or data[1][1] == ''

    def test_lag_partition(self, loot_file):
        """LAG + PARTITION BY"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT LootID, PropID, LAG(PropID, 1) OVER (PARTITION BY LootID ORDER BY PropID) AS prev FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        data = result['data']
        # 分区第一行prev应为None，第二行prev应为10131
        assert data[2][2] == 10131  # 第二行的prev是第一行的PropID

    def test_lead_basic(self, loot_file):
        """LEAD取后一行值"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT LootID, PropID, LEAD(PropID, 1) OVER (PARTITION BY LootID ORDER BY PropID) AS next_id FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        data = result['data']
        # 第一行的next应为10132（同一LootID内按PropID排序）
        assert data[1][2] == 10132


# ==================== FIRST_VALUE / LAST_VALUE ====================

class TestFirstLastValue:
    def test_first_value(self, loot_file):
        """FIRST_VALUE分区第一个值"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT LootID, PropID, FIRST_VALUE(PropID) OVER (PARTITION BY LootID ORDER BY PropID) AS first_id FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        data = result['data']
        # 每行的first_id都应该是10131
        for row in data[1:]:
            assert row[2] == 10131

    def test_last_value(self, loot_file):
        """LAST_VALUE分区最后一个值"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT LootID, PropID, LAST_VALUE(PropID) OVER (PARTITION BY LootID ORDER BY PropID) AS last_id FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        data = result['data']
        # 每行的last_id都应该是20133
        for row in data[1:]:
            assert row[2] == 20133


# ==================== NTILE ====================

class TestNtile:
    def test_ntile_basic(self, loot_file):
        """NTILE分桶"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT PropID, NTILE(4) OVER (ORDER BY PropID) AS bucket FROM LootList"
        )
        assert result['success'] is True
        data = result['data']
        # 15行分为4桶: 4,4,4,3
        buckets = [row[1] for row in data[1:]]
        assert buckets[-1] <= 4
        assert buckets[0] == 1


# ==================== PERCENT_RANK / CUME_DIST ====================

class TestPercentRankCumeDist:
    def test_percent_rank(self, loot_file):
        """PERCENT_RANK百分比排名"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT PropID, Level, PERCENT_RANK() OVER (ORDER BY Level) AS pct FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        data = result['data']
        # PERCENT_RANK值应在0~1之间
        for row in data[1:]:
            val = float(row[2])
            assert 0.0 <= val <= 1.0

    def test_cume_dist(self, loot_file):
        """CUME_DIST累积分布"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT PropID, Level, CUME_DIST() OVER (ORDER BY Level) AS cume FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        data = result['data']
        # CUME_DIST值应在0~1之间，最大值应接近1.0
        for row in data[1:]:
            val = float(row[2])
            assert 0.0 < val <= 1.0


# ==================== GROUP_CONCAT ====================

class TestGroupConcat:
    def test_group_concat(self, loot_file):
        """GROUP_CONCAT分组拼接"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT LootID, GROUP_CONCAT(PropID) AS ids FROM LootList WHERE LootID = 92000011 GROUP BY LootID"
        )
        assert result['success'] is True
        data = result['data']
        # 结果应为逗号分隔的PropID
        ids = str(data[1][1])
        assert '10131' in ids
        assert ',' in ids

    def test_group_concat_multiple_groups(self, loot_file):
        """GROUP_CONCAT多组"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT Quality, GROUP_CONCAT(PropID) AS ids FROM LootList WHERE LootID = 92000011 GROUP BY Quality"
        )
        assert result['success'] is True
        assert len(result['data']) > 2  # 表头 + 至少2个质量分组


# ==================== UPDATE中窗口函数 ====================

class TestUpdateWithWindow:
    def test_update_with_row_number(self, loot_file):
        """UPDATE WHERE中使用ROW_NUMBER()"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_update_query(
            loot_file,
            "UPDATE LootList SET Quality = 'TOP1' WHERE ROW_NUMBER() OVER (ORDER BY Level DESC) = 1",
            dry_run=True
        )
        assert result['success'] is True
        assert result.get('affected_rows', 0) >= 1

    def test_update_with_partition(self, loot_file):
        """UPDATE WHERE中使用PARTITION BY窗口函数"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_update_query(
            loot_file,
            "UPDATE LootList SET Quality = '组内第一' WHERE ROW_NUMBER() OVER (PARTITION BY LootID ORDER BY Level DESC) = 1",
            dry_run=True
        )
        assert result['success'] is True
        # 3个LootID分组，每组1行被标记
        assert result.get('affected_rows', 0) == 3
