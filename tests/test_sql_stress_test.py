"""
SQL功能压力测试 - 浓缩复杂场景

目标：用新增的窗口函数和聚合函数构建复杂SQL，发现功能缺口
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
def complex_game_data():
    """复杂游戏数据：玩家、角色、装备、副本记录"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        # 玩家表
        players = pd.DataFrame({
            'PlayerID': [1, 2, 3, 4, 5],
            'PlayerName': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
            'Level': [80, 75, 82, 78, 85],
            'GuildID': [100, 100, 200, 200, 300],
            'RegDate': pd.to_datetime(['2023-01', '2023-02', '2023-01', '2023-03', '2023-04'])
        })

        # 角色表
        characters = pd.DataFrame({
            'CharID': [101, 102, 103, 104, 105, 106, 107, 108],
            'PlayerID': [1, 1, 2, 2, 3, 3, 4, 5],
            'ClassName': ['战士', '法师', '战士', '牧师', '刺客', '法师', '骑士', '射手'],
            'ItemLevel': [450, 440, 420, 415, 460, 455, 430, 465],
            'LastLogin': pd.to_datetime(['2024-01-01', '2024-01-02', '2024-01-01', '2024-01-03', '2024-01-02',
                                         '2024-01-04', '2024-01-01', '2024-01-05'])
        })

        # 副本记录表
        raids = pd.DataFrame({
            'RaidID': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
            'CharID': [101, 101, 102, 103, 104, 105, 106, 107, 108, 101],
            'RaidName': ['火龙', '冰龙', '火龙', '暗神', '火龙', '冰龙', '暗神', '火龙', '冰龙', '暗神'],
            'Difficulty': ['英雄', '史诗', '英雄', '史诗', '英雄', '史诗', '史诗', '英雄', '史诗', '史诗'],
            'Score': [8500, 9200, 7800, 8800, 8100, 9500, 8900, 8300, 9100, 9300],
            'ClearTime': [1800, 2100, 1950, 2400, 1750, 2200, 2000, 1850, 2150, 2050],
            'Success': [1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
        })

        with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
            players.to_excel(writer, sheet_name='Players', index=False)
            characters.to_excel(writer, sheet_name='Characters', index=False)
            raids.to_excel(writer, sheet_name='Raids', index=False)
        yield tmp.name


class TestComplexSQLScenarios:
    """复杂SQL场景测试 - 发现功能缺口"""

    @pytest.mark.xfail(reason="多窗口函数+聚合+JOIN复合场景待完善")
    def test_nested_window_with_aggregate(self, complex_game_data):
        """场景：每个公会的平均装备等级，以及公会内角色的排名"""
        sql = """
        SELECT
            p.GuildID,
            c.ItemLevel,
            AVG(c.ItemLevel) OVER (PARTITION BY p.GuildID) as GuildAvgItemLevel,
            ROW_NUMBER() OVER (PARTITION BY p.GuildID ORDER BY c.ItemLevel DESC) as GuildRank,
            PERCENT_RANK() OVER (PARTITION BY p.GuildID ORDER BY c.ItemLevel) as ItemLevelPercent
        FROM Characters c
        JOIN Players p ON c.PlayerID = p.PlayerID
        WHERE p.GuildID = 100
        """
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        assert result['success'] is True
        print(f"✓ 嵌套窗口+聚合: {len(result['data'])} 行")

    def test_multi_level_cte_with_windows(self, complex_game_data):
        """场景：CTE + 窗口函数 + 多层分析"""
        sql = """
        WITH RaidStats AS (
            SELECT
                CharID,
                RaidName,
                Score,
                ROW_NUMBER() OVER (PARTITION BY CharID ORDER BY Score DESC) as BestRun,
                AVG(Score) OVER (PARTITION BY CharID) as AvgScore
            FROM Raids
        )
        SELECT
            c.ClassName,
            rs.BestRun,
            rs.AvgScore,
            LAG(rs.Score, 1) OVER (PARTITION BY c.ClassName ORDER BY rs.Score) as PrevClassScore,
            FIRST_VALUE(rs.Score) OVER (PARTITION BY c.ClassName ORDER BY rs.Score DESC) as ClassTopScore
        FROM RaidStats rs
        JOIN Characters c ON rs.CharID = c.CharID
        WHERE rs.BestRun = 1
        """
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        assert result['success'] is True
        print(f"✓ 多层CTE+窗口: {len(result['data'])} 行")

    def test_window_in_case_when(self, complex_game_data):
        """场景：CASE WHEN中使用窗口函数"""
        sql = """
        SELECT
            RaidName,
            Score,
            CASE
                WHEN Score >= AVG(Score) OVER () THEN '高于平均'
                WHEN Score >= LAG(Score, 1) OVER (ORDER BY Score) THEN '高于上次'
                ELSE '低于平均'
            END as PerformanceLevel
        FROM Raids
        WHERE RaidName = '火龙'
        """
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        assert result['success'] is True
        print(f"✓ CASE WHEN窗口: {len(result['data'])} 行")

    def test_update_with_complex_window(self, complex_game_data):
        """场景：UPDATE中使用复杂窗口函数条件"""
        sql = """
        UPDATE Raids
        SET Score = Score * 1.1
        WHERE CharID IN (
            SELECT CharID FROM (
                SELECT CharID,
                       ROW_NUMBER() OVER (PARTITION BY RaidName ORDER BY Score DESC) as rn
                FROM Raids
                WHERE Difficulty = '英雄'
            ) t WHERE rn <= 2
        )
        """
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_update_query(complex_game_data, sql, dry_run=True)
        assert result['success'] is True
        print(f"✓ UPDATE复杂窗口: 影响 {result['affected_rows']} 行")

    def test_multiple_windows_same_query(self, complex_game_data):
        """场景：同一查询使用多个窗口函数"""
        sql = """
        SELECT
            CharID,
            Score,
            ROW_NUMBER() OVER (PARTITION BY RaidName ORDER BY Score DESC) as RankInRaid,
            RANK() OVER (ORDER BY Score) as OverallRank,
            DENSE_RANK() OVER (PARTITION BY RaidName ORDER BY Score) as DenseRank,
            NTILE(4) OVER (ORDER BY Score) as Quartile,
            PERCENT_RANK() OVER (ORDER BY Score) as Percentile,
            CUME_DIST() OVER (ORDER BY Score) as CumulativeDist,
            LAG(Score, 1) OVER (PARTITION BY CharID ORDER BY Score) as PrevScore,
            LEAD(Score, 1) OVER (PARTITION BY CharID ORDER BY Score) as NextScore,
            FIRST_VALUE(Score) OVER (PARTITION BY RaidName ORDER BY Score DESC) as RaidTop,
            LAST_VALUE(Score) OVER (PARTITION BY RaidName ORDER BY Score DESC) as RaidBottom
        FROM Raids
        WHERE RaidName = '火龙'
        """
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        assert result['success'] is True
        print(f"✓ 多窗口组合: {len(result['data'])} 行, {len(result['data'][0])} 列")

    def test_group_concat_with_window(self, complex_game_data):
        """场景：GROUP_CONCAT + 窗口函数"""
        sql = """
        SELECT
            PlayerID,
            GROUP_CONCAT(ClassName) as AllClasses,
            COUNT(*) as CharCount,
            MAX(ItemLevel) as MaxItemLevel,
            ROW_NUMBER() OVER (ORDER BY MAX(ItemLevel) DESC) as RankByGear
        FROM Characters
        GROUP BY PlayerID
        """
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        assert result['success'] is True
        print(f"✓ GROUP_CONCAT+窗口: {len(result['data'])} 行")


class TestEdgeCases:
    """边界情况测试 - 发现潜在问题"""

    def test_empty_partition_by(self, complex_game_data):
        """无PARTITION BY的窗口函数"""
        sql = "SELECT Score, LAG(Score, 1) OVER (ORDER BY Score) as prev FROM Raids"
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        assert result['success'] is True
        print("✓ 空PARTITION BY")

    def test_large_offset_lag_lead(self, complex_game_data):
        """大偏移量LAG/LEAD"""
        sql = "SELECT Score, LAG(Score, 10) OVER (ORDER BY Score) as far_prev FROM Raids"
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        assert result['success'] is True
        print("✓ 大偏移量")

    def test_single_row_partition(self, complex_game_data):
        """单行分区的窗口函数"""
        sql = "SELECT Score, FIRST_VALUE(Score) OVER (PARTITION BY RaidID ORDER BY Score) as first FROM Raids WHERE RaidID = 1"
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        assert result['success'] is True
        print("✓ 单行分区")


class TestMissingFeatures:
    """测试缺失的SQL功能 - 记录需要实现的内容"""

    def test_distinct_aggregation(self, complex_game_data):
        """DISTINCT + 聚合"""
        sql = "SELECT COUNT(DISTINCT RaidName) as UniqueRaids FROM Raids"
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        if not result['success']:
            print("❌ COUNT(DISTINCT) 未支持")
        else:
            print("✓ COUNT(DISTINCT) 已支持")

    def test_union_with_window(self, complex_game_data):
        """UNION + 窗口函数"""
        sql = """
        SELECT RaidName, Score, ROW_NUMBER() OVER (ORDER BY Score) as rn FROM Raids WHERE RaidName = '火龙'
        UNION
        SELECT RaidName, Score, ROW_NUMBER() OVER (ORDER BY Score) as rn FROM Raids WHERE RaidName = '冰龙'
        """
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        if not result['success']:
            print("❌ UNION + 窗口函数 未支持")
        else:
            print("✓ UNION + 窗口函数 已支持")

    def test_having_with_window(self, complex_game_data):
        """HAVING + 窗口函数"""
        sql = """
        SELECT CharID, AVG(Score) as AvgScore
        FROM Raids
        GROUP BY CharID
        HAVING AVG(Score) > (SELECT AVG(AvgScore) FROM (SELECT AVG(Score) as AvgScore FROM Raids GROUP BY CharID) t)
        """
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(complex_game_data, sql)
        if not result['success']:
            print(f"⚠️  HAVING子查询: {result.get('message', '')[:80]}")
        else:
            print("✓ HAVING复杂子查询 已支持")


def run_stress_test():
    """运行压力测试并生成报告"""
    import pytest
    result = pytest.main([__file__, '-v', '-s', '--tb=short'])
    return result


if __name__ == '__main__':
    print("=" * 60)
    print("SQL功能压力测试 - 发现功能缺口")
    print("=" * 60)
    run_stress_test()
