"""
Excel MCP Server - Excel操作API扩展测试

测试新增的4个核心API方法：update_range, list_sheets, get_headers, create_file

@intention: 验证API扩展的正确性和完整性
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
import tempfile
import os

from src.api.excel_operations import ExcelOperations


class TestExcelOperationsExtended:
    """测试扩展后的ExcelOperations API"""

    def test_update_range_success(self):
        """测试update_range成功场景"""
        with patch('src.api.excel_operations.ExcelWriter') as mock_writer_class:
            mock_writer = MagicMock()
            mock_writer_class.return_value = mock_writer

            # 模拟成功的写入结果
            mock_result = Mock()
            mock_result.success = True
            mock_result.data = {'updated_cells': 4}
            mock_result.message = "更新成功"
            mock_writer.update_range.return_value = mock_result

            # 调用方法
            result = ExcelOperations.update_range(
                "test.xlsx",
                "Sheet1!A1:B2",
                [["姓名", "年龄"], ["张三", 25]]
            )

            # 验证结果
            assert result['success'] is True
            assert 'data' in result
            mock_writer.update_range.assert_called_once()

    def test_update_range_invalid_format(self):
        """测试update_range无效格式"""
        result = ExcelOperations.update_range(
            "test.xlsx",
            "A1:B2",  # 缺少工作表名
            [["data"]]
        )

        assert result['success'] is False
        assert 'range必须包含工作表名' in result['error']

    def test_list_sheets_success(self):
        """测试list_sheets成功场景"""
        with patch('src.api.excel_operations.ExcelReader') as mock_reader_class:
            mock_reader = MagicMock()
            mock_reader_class.return_value = mock_reader

            # 模拟工作表信息
            mock_sheet1 = Mock()
            mock_sheet1.name = "Sheet1"
            mock_sheet2 = Mock()
            mock_sheet2.name = "数据表"

            mock_result = Mock()
            mock_result.data = [mock_sheet1, mock_sheet2]
            mock_result.metadata = {
                'total_sheets': 2,
                'active_sheet': 'Sheet1'
            }
            mock_reader.list_sheets.return_value = mock_result

            # 调用方法
            result = ExcelOperations.list_sheets("test.xlsx")

            # 验证结果
            assert result['success'] is True
            assert result['sheets'] == ["Sheet1", "数据表"]
            assert result['total_sheets'] == 2
            # active_sheet概念已被移除
            assert 'active_sheet' not in result
            mock_reader.close.assert_called_once()

    def test_get_headers_success(self):
        """测试get_headers成功场景"""
        with patch('src.api.excel_operations.ExcelReader') as mock_reader_class:
            mock_reader = MagicMock()
            mock_reader_class.return_value = mock_reader

            # 模拟双行表头数据（第1行：描述，第2行：字段名）
            mock_result = Mock()
            mock_result.success = True
            mock_result.data = [
                ["ID描述", "姓名描述", "年龄描述", "部门描述"],  # 第1行：字段描述
                ["ID", "姓名", "年龄", "部门"]              # 第2行：字段名
            ]
            mock_reader.get_range.return_value = mock_result

            # 调用方法
            result = ExcelOperations.get_headers("test.xlsx", "Sheet1")

            # 验证结果 - 现在返回双行表头信息
            assert result['success'] is True
            assert result['field_names'] == ["ID", "姓名", "年龄", "部门"]
            assert result['descriptions'] == ["ID描述", "姓名描述", "年龄描述", "部门描述"]
            assert result['headers'] == ["ID", "姓名", "年龄", "部门"]  # 兼容性字段
            assert result['header_count'] == 4
            assert result['sheet_name'] == "Sheet1"
            mock_reader.close.assert_called_once()

    def test_get_headers_with_max_columns(self):
        """测试get_headers带列数限制"""
        with patch('src.api.excel_operations.ExcelReader') as mock_reader_class:
            mock_reader = MagicMock()
            mock_reader_class.return_value = mock_reader

            # 模拟双行表头数据
            mock_result = Mock()
            mock_result.success = True
            mock_result.data = [
                ["ID描述", "姓名描述", "年龄描述", "部门描述"],  # 第1行：字段描述
                ["ID", "姓名", "年龄", "部门"]              # 第2行：字段名
            ]
            mock_reader.get_range.return_value = mock_result

            # 调用方法
            result = ExcelOperations.get_headers("test.xlsx", "Sheet1", max_columns=3)

            # 验证结果 - 应该只返回前3列
            assert result['success'] is True
            assert len(result['field_names']) == 3  # 只返回前3列
            assert result['field_names'] == ["ID", "姓名", "年龄"]
            assert result['descriptions'] == ["ID描述", "姓名描述", "年龄描述"]
            assert result['headers'] == ["ID", "姓名", "年龄"]  # 兼容性字段

    def test_create_file_success(self):
        """测试create_file成功场景"""
        with patch('src.api.excel_operations.ExcelManager') as mock_manager:
            # 模拟创建结果
            mock_result = Mock()
            mock_result.success = True
            mock_result.data = {'file_path': 'test.xlsx', 'sheets': ['Sheet1']}
            mock_manager.create_file.return_value = mock_result

            # 调用方法
            result = ExcelOperations.create_file("test.xlsx")

            # 验证结果
            assert result['success'] is True
            assert 'data' in result
            mock_manager.create_file.assert_called_once_with("test.xlsx", None)

    def test_create_file_with_custom_sheets(self):
        """测试create_file带自定义工作表"""
        with patch('src.api.excel_operations.ExcelManager') as mock_manager:
            # 模拟创建结果
            mock_result = Mock()
            mock_result.success = True
            mock_result.data = {'file_path': 'test.xlsx', 'sheets': ['数据', '分析']}
            mock_manager.create_file.return_value = mock_result

            # 调用方法
            sheet_names = ['数据', '分析']
            result = ExcelOperations.create_file("test.xlsx", sheet_names)

            # 验证结果
            assert result['success'] is True
            mock_manager.create_file.assert_called_once_with("test.xlsx", sheet_names)

    def test_build_header_range_with_max_columns(self):
        """测试_build_header_range带列数限制"""
        range_expr = ExcelOperations._build_header_range("Sheet1", 1, 5)
        assert range_expr == "Sheet1!A1:E1"

    def test_build_header_range_default(self):
        """测试_build_header_range默认范围"""
        range_expr = ExcelOperations._build_header_range("Sheet1", 2, None)
        assert range_expr == "Sheet1!A2:CV2"  # CV = 第100列

    def test_parse_header_data_with_cell_info_objects(self):
        """测试_parse_header_data处理CellInfo对象"""
        # 模拟CellInfo对象
        cell1 = Mock()
        cell1.value = "ID"
        cell2 = Mock()
        cell2.value = "姓名"
        cell3 = "年龄"  # 普通字符串
        cell4 = None   # None值

        data = [[cell1, cell2, cell3, cell4]]
        headers = ExcelOperations._parse_header_data(data, None)

        assert headers == ["ID", "姓名", "年龄"]  # None值处停止

    def test_parse_header_data_with_max_columns_preserve_empty(self):
        """测试_parse_header_data在max_columns模式下保留空值"""
        data = [["ID", "", "年龄", None]]
        headers = ExcelOperations._parse_header_data(data, 4)

        assert headers == ["ID", "", "年龄", ""]  # 保留空字符串和None转换

    def test_error_handling_in_all_methods(self):
        """测试所有方法的错误处理"""
        # 测试update_range错误处理
        with patch('src.api.excel_operations.ExcelWriter', side_effect=Exception("写入错误")):
            result = ExcelOperations.update_range("test.xlsx", "Sheet1!A1:B1", [["data"]])
            assert result['success'] is False
            assert "更新范围数据失败" in result['error']

        # 测试list_sheets错误处理
        with patch('src.api.excel_operations.ExcelReader', side_effect=Exception("读取错误")):
            result = ExcelOperations.list_sheets("test.xlsx")
            assert result['success'] is False
            assert "获取工作表列表失败" in result['error']

        # 测试get_headers错误处理
        with patch('src.api.excel_operations.ExcelReader', side_effect=Exception("表头错误")):
            result = ExcelOperations.get_headers("test.xlsx", "Sheet1")
            assert result['success'] is False
            assert "获取表头失败" in result['error']

        # 测试create_file错误处理
        with patch('src.api.excel_operations.ExcelManager.create_file', side_effect=Exception("创建错误")):
            result = ExcelOperations.create_file("test.xlsx")
            assert result['success'] is False
            assert "创建文件失败" in result['error']

    def test_logging_integration(self):
        """测试日志集成"""
        with patch('src.api.excel_operations.logger') as mock_logger:
            # 启用调试日志
            ExcelOperations.DEBUG_LOG_ENABLED = True

            try:
                with patch('src.api.excel_operations.ExcelReader') as mock_reader_class:
                    mock_reader = MagicMock()
                    mock_reader_class.return_value = mock_reader

                    mock_result = Mock()
                    mock_result.data = []
                    mock_result.metadata = {}
                    mock_reader.list_sheets.return_value = mock_result

                    ExcelOperations.list_sheets("test.xlsx")

                    # 验证日志调用
                    assert mock_logger.info.called

            finally:
                # 重置日志状态
                ExcelOperations.DEBUG_LOG_ENABLED = False
