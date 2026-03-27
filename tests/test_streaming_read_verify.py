"""
REQ-015 流式写入后读取工具验证测试

验证所有6种流式写入操作（StreamingWriter）执行后，
所有读取工具是否正常工作。

写入操作: batch_insert_rows, delete_rows, update_range, insert_rows, delete_columns, upsert_row
读取操作: find_last_row, get_headers, get_range, query(SQL), list_sheets, search

Bug背景: 之前insert_rows使用StreamingWriter.update_range（覆盖模式）而非
insert_rows_streaming（插入模式），导致流式插入空行失败。
"""

import os
import sys
import tempfile
import shutil

import openpyxl
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations
from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


@pytest.fixture
def test_file():
    """创建测试Excel文件：角色表（1表头+10数据行，5列）"""
    tmpdir = tempfile.mkdtemp()
    fp = os.path.join(tmpdir, 'test_streaming.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '角色'
    ws.append(['角色ID', '名称', '职业', '等级', '攻击力'])
    for i in range(1, 11):
        ws.append([1000+i, f'角色{i}', '战士' if i%3==0 else '法师' if i%3==1 else '弓手', 10+i*2, 50+i*5])
    wb.save(fp)
    wb.close()
    yield fp
    shutil.rmtree(tmpdir)


@pytest.fixture
def sql_engine():
    return AdvancedSQLQueryEngine()


def q(engine, fp, sql, sheet='角色'):
    return engine.execute_sql_query(fp, sql, sheet_name=sheet, include_headers=False)


def count_rows(engine, fp, sheet='角色'):
    r = q(engine, fp, 'SELECT COUNT(*) as total FROM ' + sheet, sheet)
    if r['success'] and r['data']:
        for row in r['data']:
            for v in row:
                if isinstance(v, (int, float)):
                    return int(v)
    return -1


class TestStreamingBatchInsert:
    """batch_insert_rows 流式写入后读取验证"""

    def test_find_last_row(self, test_file):
        ExcelOperations.batch_insert_rows(test_file, '角色', [
            {'角色ID': 9001, '名称': '刺客A', '职业': '刺客', '等级': 50, '攻击力': 300},
        ])
        r = ExcelOperations.find_last_row(test_file, '角色')
        assert r['data']['last_row'] == 12

    def test_get_headers(self, test_file):
        ExcelOperations.batch_insert_rows(test_file, '角色', [
            {'角色ID': 9001, '名称': '刺客A', '职业': '刺客', '等级': 50, '攻击力': 300},
        ])
        r = ExcelOperations.get_headers(test_file, '角色')
        assert len(r['field_names']) == 5

    def test_get_range_new_rows(self, test_file):
        ExcelOperations.batch_insert_rows(test_file, '角色', [
            {'角色ID': 9001, '名称': '刺客A', '职业': '刺客', '等级': 50, '攻击力': 300},
            {'角色ID': 9002, '名称': '刺客B', '职业': '刺客', '等级': 55, '攻击力': 350},
        ])
        r = ExcelOperations.get_range(test_file, '角色!A12:E13')
        vals = [row[1]['value'] for row in r['data']]
        assert vals == ['刺客A', '刺客B']

    def test_query_count(self, test_file, sql_engine):
        ExcelOperations.batch_insert_rows(test_file, '角色', [
            {'角色ID': 9001, '名称': '刺客A', '职业': '刺客', '等级': 50, '攻击力': 300},
        ])
        assert count_rows(sql_engine, test_file) == 11

    def test_query_where(self, test_file, sql_engine):
        ExcelOperations.batch_insert_rows(test_file, '角色', [
            {'角色ID': 9001, '名称': '刺客A', '职业': '刺客', '等级': 50, '攻击力': 300},
        ])
        r = q(sql_engine, test_file, 'SELECT * FROM 角色 WHERE 职业 = "刺客"')
        assert len(r['data']) == 1


class TestStreamingDeleteRows:
    """delete_rows 流式写入后读取验证"""

    def test_find_last_row(self, test_file):
        ExcelOperations.delete_rows(test_file, '角色', 10, 2)
        r = ExcelOperations.find_last_row(test_file, '角色')
        assert r['data']['last_row'] == 9

    def test_query_count(self, test_file, sql_engine):
        ExcelOperations.delete_rows(test_file, '角色', 10, 2)
        assert count_rows(sql_engine, test_file) == 8


class TestStreamingUpdateRange:
    """update_range 流式覆盖写入后读取验证"""

    def test_get_range_after_update(self, test_file):
        ExcelOperations.update_range(test_file, '角色!E2:E4', [[999], [888], [777]], insert_mode=False)
        r = ExcelOperations.get_range(test_file, '角色!E2:E4')
        vals = [row[0]['value'] for row in r['data']]
        assert vals == [999, 888, 777]

    def test_query_after_update(self, test_file, sql_engine):
        ExcelOperations.update_range(test_file, '角色!E2:E4', [[999], [888], [777]], insert_mode=False)
        r = q(sql_engine, test_file, 'SELECT 攻击力 FROM 角色 WHERE 角色ID = 1001')
        assert r['data'][0][0] == 999


class TestStreamingInsertRows:
    """insert_rows 流式插入空行后读取验证"""

    def test_find_last_row(self, test_file):
        ExcelOperations.insert_rows(test_file, '角色', 2, 2)
        r = ExcelOperations.find_last_row(test_file, '角色')
        assert r['data']['last_row'] == 13  # 11 + 2 inserted

    def test_query_count(self, test_file, sql_engine):
        ExcelOperations.insert_rows(test_file, '角色', 2, 2)
        assert count_rows(sql_engine, test_file) == 12  # 10 + 2 inserted

    def test_get_range_inserted_empty(self, test_file):
        ExcelOperations.insert_rows(test_file, '角色', 2, 1)
        r = ExcelOperations.get_range(test_file, '角色!A2:E2')
        # Inserted row should be empty
        vals = [cell['value'] for cell in r['data'][0]]
        assert all(v in (None, '') for v in vals)


class TestStreamingDeleteColumns:
    """delete_columns 流式写入后读取验证"""

    def test_get_headers_after_delete(self, test_file):
        ExcelOperations.delete_columns(test_file, '角色', 5, 1)
        r = ExcelOperations.get_headers(test_file, '角色')
        assert len(r['field_names']) == 4
        assert '攻击力' not in r['field_names']

    def test_query_after_delete(self, test_file, sql_engine):
        ExcelOperations.delete_columns(test_file, '角色', 5, 1)
        r = q(sql_engine, test_file, 'SELECT * FROM 角色 WHERE 角色ID = 1001')
        # Should only have 4 columns
        assert len(r['data'][0]) == 4


class TestStreamingUpsertRow:
    """upsert_row 流式写入后读取验证"""

    def test_insert_find_last_row(self, test_file):
        ExcelOperations.upsert_row(test_file, '角色', '角色ID', 7777, {'名称': '新角色', '职业': '刺客', '等级': 1})
        r = ExcelOperations.find_last_row(test_file, '角色')
        assert r['data']['last_row'] == 12

    def test_insert_query(self, test_file, sql_engine):
        ExcelOperations.upsert_row(test_file, '角色', '角色ID', 7777, {'名称': '新角色', '职业': '刺客', '等级': 1})
        r = q(sql_engine, test_file, 'SELECT * FROM 角色 WHERE 角色ID = 7777')
        assert len(r['data']) == 1

    def test_update_query(self, test_file, sql_engine):
        ExcelOperations.upsert_row(test_file, '角色', '角色ID', 1001, {'等级': 888})
        r = q(sql_engine, test_file, 'SELECT 等级 FROM 角色 WHERE 角色ID = 1001')
        assert r['data'][0][0] == 888


class TestStreamingCrossTool:
    """流式写入后跨工具一致性验证"""

    def test_batch_insert_then_delete_consistency(self, test_file, sql_engine):
        """插入后删除，数据应恢复原状"""
        ExcelOperations.batch_insert_rows(test_file, '角色', [
            {'角色ID': 9001, '名称': '刺客A', '职业': '刺客', '等级': 50, '攻击力': 300},
        ])
        ExcelOperations.delete_rows(test_file, '角色', 12, 1)
        assert count_rows(sql_engine, test_file) == 10

    def test_upsert_update_then_query_consistency(self, test_file, sql_engine):
        """upsert更新后，query应看到新值"""
        ExcelOperations.upsert_row(test_file, '角色', '角色ID', 1001, {'名称': '改名', '等级': 777})
        r = q(sql_engine, test_file, 'SELECT 名称, 等级 FROM 角色 WHERE 角色ID = 1001')
        assert r['data'][0][0] == '改名'
        assert r['data'][0][1] == 777

    def test_update_range_then_search_consistency(self, test_file):
        """update_range后，search应能找到新值"""
        ExcelOperations.update_range(test_file, '角色!B2', [['改名称']], insert_mode=False)
        r = ExcelOperations.search(test_file, '改名称', sheet_name='角色')
        assert r['success']
        assert len(r['data']) > 0

    def test_multiple_operations_sequential(self, test_file, sql_engine):
        """连续多种流式操作后，所有读取工具正常"""
        # batch_insert
        ExcelOperations.batch_insert_rows(test_file, '角色', [
            {'角色ID': 9001, '名称': 'A', '职业': '刺客', '等级': 50, '攻击力': 300},
        ])
        # delete_rows
        ExcelOperations.delete_rows(test_file, '角色', 3, 1)
        # update_range
        ExcelOperations.update_range(test_file, '角色!E2', [[999]], insert_mode=False)
        # upsert
        ExcelOperations.upsert_row(test_file, '角色', '角色ID', 1001, {'等级': 888})

        # Verify
        assert ExcelOperations.find_last_row(test_file, '角色')['data']['last_row'] == 11
        assert len(ExcelOperations.get_headers(test_file, '角色')['field_names']) == 5
        assert count_rows(sql_engine, test_file) == 10
        r = q(sql_engine, test_file, 'SELECT 等级, 攻击力 FROM 角色 WHERE 角色ID = 1001')
        assert r['data'][0][0] == 888
        assert r['data'][0][1] == 999
