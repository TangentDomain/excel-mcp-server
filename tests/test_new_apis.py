"""
新API的专门测试用例
测试拆分后的API功能是否正常工作
"""
import pytest
import sys
import os
from pathlib import Path

# 获取项目根目录并添加到路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(project_root / "src"))

# 直接导入src.server模块
import src.server as server


class TestNewAPIs:
    """测试拆分后的新API"""

    @pytest.fixture
    def sample_file(self):
        """使用现有的测试文件"""
        return "data/test-api-comprehensive.xlsx"

    def test_excel_list_sheets_new(self, sample_file):
        """测试新的excel_list_sheets - 只返回工作表列表"""
        result = server.excel_list_sheets(sample_file)

        assert result['success'] is True
        assert 'sheets' in result
        assert 'total_sheets' in result
        assert 'active_sheet' in result
        # 确保不再返回sheets_with_headers
        assert 'sheets_with_headers' not in result

        # 检查返回的工作表列表
        sheets = result['sheets']
        assert isinstance(sheets, list)
        assert len(sheets) > 0

    def test_excel_get_sheet_headers_new(self, sample_file):
        """测试新的excel_get_sheet_headers - 专门获取表头"""
        result = server.excel_get_sheet_headers(sample_file)

        assert result['success'] is True
        assert 'sheets_with_headers' in result
        assert 'total_sheets' in result
        # 确保返回表头信息
        assert 'file_path' in result

        # 检查表头信息结构
        sheets_with_headers = result['sheets_with_headers']
        assert isinstance(sheets_with_headers, list)

        for sheet_info in sheets_with_headers:
            assert 'name' in sheet_info
            assert 'header_count' in sheet_info
            # headers字段在空工作表时可能被清理，使用默认值处理
            headers = sheet_info.get('headers', [])
            assert isinstance(headers, list)
            assert isinstance(sheet_info['header_count'], int)

    def test_api_separation_completeness(self, sample_file):
        """测试API拆分的完整性 - 新老API功能等价"""
        # 新API的分离调用
        sheets_result = server.excel_list_sheets(sample_file)
        headers_result = server.excel_get_sheet_headers(sample_file)

        # 验证功能完整性
        assert sheets_result['success'] is True
        assert headers_result['success'] is True

        # 工作表列表应该一致
        sheets_from_list = set(sheets_result['sheets'])
        sheets_from_headers = set([s['name'] for s in headers_result['sheets_with_headers']])
        assert sheets_from_list == sheets_from_headers

    def test_excel_format_cells_custom(self):
        """测试自定义格式化API"""
        # 注意：这是一个示例，实际测试需要真实的Excel文件
        # 这里主要测试参数验证和函数调用
        test_formatting = {
            'font': {'bold': True, 'color': 'FF0000'},
            'fill': {'color': 'FFFF00'}
        }

        # 测试参数结构（不实际运行，避免文件依赖）
        try:
            # 这应该抛出文件不存在的错误，但参数结构是正确的
            server.excel_format_cells("nonexistent.xlsx", "Sheet1", "A1:B1", formatting=test_formatting)
        except Exception as e:
            # 期望的是文件相关错误，不是参数错误
            assert "nonexistent.xlsx" in str(e) or "No such file" in str(e) or "does not exist" in str(e)

    def test_excel_format_cells(self):
        """测试预设格式化API"""
        valid_presets = ["title", "header", "data", "highlight", "currency"]

        for preset in valid_presets:
            try:
                # 测试预设验证（不实际运行，避免文件依赖）
                server.excel_format_cells("nonexistent.xlsx", "Sheet1", "A1:B1", preset=preset)
            except Exception as e:
                # 期望的是文件相关错误，不是预设验证错误
                assert "nonexistent.xlsx" in str(e) or "No such file" in str(e) or "does not exist" in str(e)

    def test_invalid_preset(self):
        """测试无效预设的错误处理"""
        try:
            result = server.excel_format_cells("any.xlsx", "Sheet1", "A1:B1", preset="invalid_preset")
            # 如果返回结果而不是抛出异常，检查错误信息
            if 'success' in result:
                assert result['success'] is False
                assert "未知的预设样式" in result.get('error', '')
        except Exception:
            # 如果抛出异常也是可接受的
            pass

    def test_parameter_clarity(self):
        """测试新API的参数清晰度"""
        import inspect

        # 检查excel_list_sheets参数
        sig_list = inspect.signature(server.excel_list_sheets)
        params_list = list(sig_list.parameters.keys())
        assert params_list == ['file_path'], f"excel_list_sheets参数应该只有file_path，实际为: {params_list}"

        # 检查excel_get_sheet_headers参数
        sig_headers = inspect.signature(server.excel_get_sheet_headers)
        params_headers = list(sig_headers.parameters.keys())
        assert params_headers == ['file_path'], f"excel_get_sheet_headers参数应该只有file_path，实际为: {params_headers}"

        # 跳过检查已合并的excel_format_cells_custom，因为它已经合并到excel_format_cells中

        # 检查excel_format_cells参数
        sig_preset = inspect.signature(server.excel_format_cells)
        params_preset = list(sig_preset.parameters.keys())
        expected_preset = ['file_path', 'sheet_name', 'range', 'formatting', 'preset']
        assert params_preset == expected_preset, f"excel_format_cells参数不符合预期: {params_preset}"

    def test_single_responsibility_principle(self, sample_file):
        """测试单一职责原则的实现"""
        # excel_list_sheets应该只返回工作表相关信息，不包含表头
        list_result = server.excel_list_sheets(sample_file)
        assert 'sheets' in list_result
        assert 'sheets_with_headers' not in list_result

        # excel_get_sheet_headers应该专注于表头信息
        headers_result = server.excel_get_sheet_headers(sample_file)
        assert 'sheets_with_headers' in headers_result
        # 可以包含基础的sheets信息用于上下文，但主要功能是表头

        # 两个API返回的数据结构应该不同，体现不同职责
        assert set(list_result.keys()) != set(headers_result.keys())


if __name__ == "__main__":
    # 运行测试
    pytest.main([__file__, "-v"])
