"""REQ-015: 流式写入后读取工具验证

验证StreamingWriter使用write_only模式写入后，所有读取工具仍能正常工作。
核心问题：write_only模式不写<dimension>元数据，导致read_only模式下max_row/max_column返回None。
"""
import pytest
import tempfile
import os
from openpyxl import Workbook


@pytest.fixture
def streamed_file():
    """创建经过流式写入的测试文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "技能表"
    ws.append(["技能ID", "名称", "伤害", "类型"])
    for i in range(1, 6):
        ws.append([i, f"技能_{i}", i * 100, "fire" if i % 2 == 0 else "ice"])

    ws2 = wb.create_sheet("装备表")
    ws2.append(["装备ID", "名称", "品质"])
    ws2.append([1, "木剑", "common"])
    ws2.append([2, "铁剑", "rare"])

    tmp = tempfile.mktemp(suffix=".xlsx")
    wb.save(tmp)
    yield tmp
    if os.path.exists(tmp):
        os.unlink(tmp)


class TestStreamingWriteReadCompat:
    """流式写入后读取工具兼容性测试"""

    def test_describe_table_after_streaming(self, streamed_file):
        """describe_table在流式写入后应正常工作"""
        from src.excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
        from src.excel_mcp_server_fastmcp.server import excel_describe_table

        StreamingWriter.batch_insert_rows(streamed_file, "技能表", [
            {"技能ID": 6, "名称": "streamed", "伤害": 600, "类型": "fire"},
        ])

        result = excel_describe_table(streamed_file, "技能表")
        assert result['success'], result.get('message', '')
        assert result['data']['row_count'] == 6
        assert result['data']['column_count'] == 4

    def test_get_headers_after_streaming(self, streamed_file):
        """get_headers在流式写入后应正常工作"""
        from src.excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
        from src.excel_mcp_server_fastmcp.server import excel_get_headers

        StreamingWriter.batch_insert_rows(streamed_file, "技能表", [
            {"技能ID": 6, "名称": "streamed", "伤害": 600, "类型": "fire"},
        ])

        result = excel_get_headers(streamed_file, "技能表")
        assert result['success'], result.get('message', '')

    def test_get_range_after_streaming(self, streamed_file):
        """get_range在流式写入后应正常工作"""
        from src.excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
        from src.excel_mcp_server_fastmcp.server import excel_get_range

        StreamingWriter.batch_insert_rows(streamed_file, "技能表", [
            {"技能ID": 6, "名称": "streamed", "伤害": 600, "类型": "fire"},
        ])

        result = excel_get_range(streamed_file, "技能表!A1:D3")
        assert result['success'], result.get('message', '')

    def test_find_last_row_after_streaming(self, streamed_file):
        """find_last_row在流式写入后应正确返回新行数"""
        from src.excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
        from src.excel_mcp_server_fastmcp.server import excel_find_last_row

        StreamingWriter.batch_insert_rows(streamed_file, "技能表", [
            {"技能ID": 6, "名称": "streamed", "伤害": 600, "类型": "fire"},
        ])

        result = excel_find_last_row(streamed_file, "技能表")
        assert result['success'], result.get('message', '')
        assert result['data']['last_row'] == 7  # 1 header + 5 original + 1 streamed

    def test_find_last_row_with_column_after_streaming(self, streamed_file):
        """find_last_row(带列参数)在流式写入后应正常工作"""
        from src.excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
        from src.excel_mcp_server_fastmcp.server import excel_find_last_row

        StreamingWriter.batch_insert_rows(streamed_file, "技能表", [
            {"技能ID": 6, "名称": "streamed", "伤害": 600, "类型": "fire"},
        ])

        result = excel_find_last_row(streamed_file, "技能表", "A")
        assert result['success'], result.get('message', '')
        assert result['data']['last_row'] == 7

    def test_check_duplicate_ids_after_streaming(self, streamed_file):
        """check_duplicate_ids在流式写入后应正常工作（max_row/max_column=None降级）"""
        from src.excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
        from src.excel_mcp_server_fastmcp.server import excel_check_duplicate_ids

        StreamingWriter.batch_insert_rows(streamed_file, "技能表", [
            {"技能ID": 1, "名称": "dup", "伤害": 999, "类型": "fire"},  # duplicate ID
        ])

        result = excel_check_duplicate_ids(streamed_file, "技能表", 1)
        assert result['success'], result.get('message', '')
        assert result['has_duplicates'] is True
        assert result['duplicate_count'] == 1

    def test_sql_query_after_streaming(self, streamed_file):
        """SQL查询在流式写入后应正常工作"""
        from src.excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

        StreamingWriter.batch_insert_rows(streamed_file, "技能表", [
            {"技能ID": 6, "名称": "streamed", "伤害": 600, "类型": "fire"},
        ])

        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(streamed_file, "SELECT * FROM 技能表 WHERE 伤害 >= 500")
        assert result['success'], result.get('message', '')
        # Should find rows with 伤害=500 and 伤害=600
        data_rows = len(result['data']) - 1  # exclude header
        assert data_rows == 2

    def test_sql_join_after_streaming(self, streamed_file):
        """SQL JOIN在流式写入后应正常工作"""
        from src.excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
        from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

        StreamingWriter.batch_insert_rows(streamed_file, "技能表", [
            {"技能ID": 6, "名称": "streamed", "伤害": 600, "类型": "fire"},
        ])

        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            streamed_file,
            "SELECT s.名称, d.名称 AS 装备名称 FROM 技能表 s JOIN 装备表 d ON s.技能ID = d.装备ID"
        )
        assert result['success'], result.get('message', '')

    def test_search_after_streaming(self, streamed_file):
        """search在流式写入后应正常工作"""
        from src.excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
        from src.excel_mcp_server_fastmcp.server import excel_search

        StreamingWriter.batch_insert_rows(streamed_file, "技能表", [
            {"技能ID": 6, "名称": "streamed_unique", "伤害": 600, "类型": "fire"},
        ])

        result = excel_search(streamed_file, "streamed_unique", "技能表")
        assert result['success'], result.get('message', '')
