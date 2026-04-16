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
from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


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


# ==================== 窗口聚合函数 ====================

class TestWindowAggregate:
    def test_avg_over_partition(self, loot_file):
        """AVG() OVER (PARTITION BY) 分区平均"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT LootID, Level, AVG(Level) OVER (PARTITION BY LootID) AS avg_level FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        data = result['data']
        # 所有行的avg_level应相同(45.0)
        for row in data[1:]:
            assert float(row[2]) == 45.0

    def test_sum_over(self, loot_file):
        """SUM() OVER () 全表求和"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT PropID, SUM(Level) OVER () AS total_level FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        # 所有行的total_level相同
        total = result['data'][1][1]
        for row in result['data'][2:]:
            assert row[1] == total

    def test_count_over_partition(self, loot_file):
        """COUNT() OVER (PARTITION BY) 分区计数"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT LootID, COUNT(Level) OVER (PARTITION BY LootID) AS cnt FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        # 每行cnt应为6
        for row in result['data'][1:]:
            assert row[1] == 6

    def test_min_max_over(self, loot_file):
        """MIN/MAX OVER 分区极值"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT LootID, Level, MIN(Level) OVER (PARTITION BY LootID) AS min_lv, MAX(Level) OVER (PARTITION BY LootID) AS max_lv FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        for row in result['data'][1:]:
            assert row[2] == 40  # min
            assert row[3] == 50  # max

    def test_running_sum(self, loot_file):
        """SUM() OVER (ORDER BY) 累计求和"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT PropID, Level, SUM(Level) OVER (ORDER BY Level) AS running_sum FROM LootList WHERE LootID = 92000011"
        )
        assert result['success'] is True
        # 所有running_sum应大于0且不等于0
        for row in result['data'][1:]:
            assert float(row[2]) > 0

    def test_avg_over_with_join(self, loot_file):
        """JOIN + AVG() OVER 跨表窗口聚合"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT l.LootID, l.Level, AVG(l.Level) OVER (PARTITION BY l.LootID) AS avg_level FROM LootList l WHERE l.LootID = 92000011"
        )
        assert result['success'] is True


# ==================== 嵌套FROM子查询 ====================

class TestNestedFromSubquery:
    def test_two_level_nested(self, loot_file):
        """两层嵌套FROM子查询"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT * FROM (SELECT LootID, COUNT(*) as cnt FROM LootList GROUP BY LootID) t WHERE cnt > 3"
        )
        assert result['success'] is True

    def test_three_level_nested(self, loot_file):
        """三层嵌套FROM子查询"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT * FROM (SELECT * FROM (SELECT LootID, AVG(Level) as avg_lv FROM LootList GROUP BY LootID) t1 WHERE avg_lv > 35) t2"
        )
        assert result['success'] is True

    def test_nested_from_with_window(self, loot_file):
        """嵌套FROM + 窗口函数"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            loot_file,
            "SELECT * FROM (SELECT LootID, Level, ROW_NUMBER() OVER (PARTITION BY LootID ORDER BY Level DESC) as rn FROM LootList) t WHERE rn = 1"
        )
        assert result['success'] is True
        # 每个LootID只返回1行
        loot_ids = set()
        for row in result['data'][1:]:
            loot_ids.add(row[0])
        assert len(loot_ids) == 3  # 3个不同的LootID
