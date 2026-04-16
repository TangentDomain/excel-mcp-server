"""
INSERT/DELETE SQL + 跨文件JOIN 测试
"""

import pytest
import tempfile
import pandas as pd
import os
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))
from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


@pytest.fixture
def raids_file(tmp_path):
    """基础测试文件"""
    file_path = str(tmp_path / 'raids.xlsx')
    df = pd.DataFrame({
        'RID': [1, 2, 3, 4, 5],
        'CID': [101, 101, 102, 103, 104],
        'Name': ['fire', 'ice', 'fire', 'dark', 'ice'],
        'Score': [85, 92, 78, 88, 95],
    })
    df.to_excel(file_path, index=False, sheet_name='Raids')
    return file_path


# ==================== INSERT ====================

class TestInsert:
    def test_insert_single_row(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_insert_query(
            raids_file,
            "INSERT INTO Raids (RID, CID, Name, Score) VALUES (6, 105, 'fire', 70)"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 1

        # 验证
        verify = engine.execute_sql_query(raids_file, "SELECT * FROM Raids WHERE RID = 6")
        assert verify['success'] is True
        assert len(verify['data']) == 2  # header + 1 row

    def test_insert_multi_row(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_insert_query(
            raids_file,
            "INSERT INTO Raids (RID, Score) VALUES (6, 70), (7, 80)"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 2

    def test_insert_dry_run(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_insert_query(
            raids_file,
            "INSERT INTO Raids (RID, Score) VALUES (6, 70)",
            dry_run=True
        )
        assert result['success'] is True
        assert result['dry_run'] is True
        assert result['affected_rows'] == 1

        # 验证没有实际插入
        verify = engine.execute_sql_query(raids_file, "SELECT COUNT(*) as cnt FROM Raids")
        assert verify['data'][1][0] == 5  # 仍然是5行

    def test_insert_wrong_column(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_insert_query(
            raids_file,
            "INSERT INTO Raids (NonExist, Score) VALUES (1, 70)"
        )
        assert result['success'] is False
        assert '不存在' in result['message']

    def test_insert_wrong_table(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_insert_query(
            raids_file,
            "INSERT INTO NonExist (RID) VALUES (1)"
        )
        assert result['success'] is False
        assert '不存在' in result['message']

    def test_insert_null_value(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_insert_query(
            raids_file,
            "INSERT INTO Raids (RID, Name, Score) VALUES (6, NULL, 70)"
        )
        assert result['success'] is True

    def test_insert_negative_value(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_insert_query(
            raids_file,
            "INSERT INTO Raids (RID, Score) VALUES (6, -5)"
        )
        assert result['success'] is True


# ==================== DELETE ====================

class TestDelete:
    def test_delete_with_where(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_delete_query(
            raids_file,
            "DELETE FROM Raids WHERE Score < 80"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 1  # RID=3, Score=78

    def test_delete_with_row_number(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_delete_query(
            raids_file,
            "DELETE FROM Raids WHERE _ROW_NUMBER_ IN (1, 3)"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 2

    def test_delete_no_match(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_delete_query(
            raids_file,
            "DELETE FROM Raids WHERE Score > 200"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 0

    def test_delete_dry_run(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_delete_query(
            raids_file,
            "DELETE FROM Raids WHERE Score < 80",
            dry_run=True
        )
        assert result['success'] is True
        assert result['dry_run'] is True
        assert result['affected_rows'] == 1

        # 验证没有实际删除
        verify = engine.execute_sql_query(raids_file, "SELECT COUNT(*) as cnt FROM Raids")
        assert verify['data'][1][0] == 5

    def test_delete_without_where(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_delete_query(
            raids_file,
            "DELETE FROM Raids"
        )
        assert result['success'] is False
        assert 'WHERE' in result['message']

    def test_delete_wrong_table(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_delete_query(
            raids_file,
            "DELETE FROM NonExist WHERE Score > 0"
        )
        assert result['success'] is False

    def test_delete_with_complex_condition(self, raids_file):
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_delete_query(
            raids_file,
            "DELETE FROM Raids WHERE Name = 'ice' AND Score > 90"
        )
        assert result['success'] is True
        assert result['affected_rows'] == 2  # RID=2(ice,92), RID=5(ice,95)


# ==================== 跨文件JOIN ====================

class TestCrossFileJoin:
    def test_cross_file_join(self, tmp_path):
        """测试跨文件JOIN (@'path'语法)"""
        # 创建两个文件
        file1 = str(tmp_path / 'players.xlsx')
        file2 = str(tmp_path / 'chars.xlsx')

        pd.DataFrame({
            'PID': [1, 2, 3],
            'Guild': [100, 100, 200]
        }).to_excel(file1, index=False, sheet_name='Players')

        pd.DataFrame({
            'CID': [101, 102, 103],
            'PID': [1, 1, 2],
            'ILv': [450, 440, 420]
        }).to_excel(file2, index=False, sheet_name='Chars')

        engine = AdvancedSQLQueryEngine()

        # 使用相对路径引用file2
        rel_path = './chars.xlsx'
        sql = f"SELECT c.CID, p.Guild FROM Chars@'{rel_path}' c JOIN Players c2 ON c.PID = c2.PID"

        result = engine.execute_sql_query(file1, sql)
        assert result['success'] is True
        assert len(result['data']) > 1  # header + data rows

    def test_cross_file_file_not_found(self, tmp_path):
        """跨文件引用不存在的文件"""
        file1 = str(tmp_path / 'players.xlsx')
        pd.DataFrame({'PID': [1]}).to_excel(file1, index=False, sheet_name='Players')

        engine = AdvancedSQLQueryEngine()
        sql = "SELECT * FROM Chars@'./nonexistent.xlsx'"
        result = engine.execute_sql_query(file1, sql)
        assert result['success'] is False
        assert '不存在' in result['message']
