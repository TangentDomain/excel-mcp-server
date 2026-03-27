"""
测试统一返回值结构。

所有工具成功时必须返回:
  {success: True, message: "...", data: {...}, meta?: {...}}

data 字段包含实际数据载荷，AI 客户端只需检查 result.data。
"""
import pytest
from src.excel_mcp_server_fastmcp.server import (
    excel_list_sheets, excel_get_headers, excel_find_last_row,
    excel_query, excel_get_range, excel_search
)


class TestUnifiedReturnStructure:
    """验证所有工具返回统一的 {success, message, data} 结构"""

    @pytest.fixture
    def game_config(self):
        return "tests/test_data/game_config.xlsx"

    def test_list_sheets_has_data_dict(self, game_config):
        """list_sheets 成功时 data 必须是 dict"""
        result = excel_list_sheets(game_config)
        assert result['success'] is True
        assert isinstance(result['data'], dict)
        assert 'sheets' in result['data']
        assert isinstance(result['data']['sheets'], list)
        assert 'message' in result

    def test_get_headers_has_data_dict(self, game_config):
        """get_headers 成功时 data 必须是 dict"""
        result = excel_get_headers(game_config, "技能配置")
        assert result['success'] is True
        assert isinstance(result['data'], dict)
        assert 'field_names' in result['data']
        assert 'headers' in result['data']
        assert 'descriptions' in result['data']
        assert 'header_count' in result['data']
        assert 'message' in result

    def test_find_last_row_has_data_dict(self, game_config):
        """find_last_row 成功时 data 必须是 dict"""
        result = excel_find_last_row(game_config, "技能配置")
        assert result['success'] is True
        assert isinstance(result['data'], dict)
        assert 'last_row' in result['data']
        assert 'sheet_name' in result['data']
        assert 'message' in result

    def test_query_data_is_list(self, game_config):
        """SQL query 的 data 仍然是 list（标准格式不受影响）"""
        result = excel_query(game_config, "SELECT * FROM 技能配置 LIMIT 2")
        assert result['success'] is True
        assert isinstance(result['data'], list)
        assert 'query_info' in result
        assert 'message' in result

    def test_query_json_keeps_formatted_output(self, game_config):
        """SQL query JSON 格式的 formatted_output 保持在顶层"""
        result = excel_query(
            game_config,
            "SELECT skill_name FROM 技能配置 LIMIT 1",
            output_format='json'
        )
        assert result['success'] is True
        assert 'formatted_output' in result
        assert isinstance(result['data'], list)

    def test_error_response_has_message(self, game_config):
        """错误响应必须有 message 字段"""
        result = excel_list_sheets("/nonexistent/file.xlsx")
        assert result['success'] is False
        assert 'message' in result

    def test_no_leaking_data_keys(self, game_config):
        """成功响应的顶层不应有数据载荷键（应全在 data 内）"""
        result = excel_list_sheets(game_config)
        assert result['success'] is True
        top_keys = set(result.keys()) - {'success', 'message', 'data', 'meta'}
        assert top_keys == set(), f"Found non-standard top-level keys: {top_keys}"

    def test_get_all_headers_has_data_dict(self, game_config):
        """get_all_headers（不传 sheet_name）data 必须是 dict"""
        result = excel_get_headers(game_config)
        assert result['success'] is True
        assert isinstance(result['data'], dict)
        assert 'sheets_with_headers' in result['data']
        assert 'total_sheets' in result['data']
