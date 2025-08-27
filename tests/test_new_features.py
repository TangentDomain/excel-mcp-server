"""
Excel MCP Server - 新功能单元测试

测试新添加的格式化和文件操作功能
"""
import os
import tempfile
import pytest
from pathlib import Path
from unittest.mock import Mock, patch

# 导入要测试的模块
from src.server import (
    excel_merge_cells,
    excel_unmerge_cells,
    excel_set_borders,
    excel_set_row_height,
    excel_set_column_width,
    excel_get_file_info,
    excel_create_file,
    excel_update_range
)


class TestNewFeatures:
    """测试Excel MCP Server新功能"""

    @pytest.fixture
    def sample_excel_file(self):
        """创建临时测试Excel文件"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
            temp_path = temp_file.name

        # 创建文件并添加测试数据
        result = excel_create_file(temp_path, ["测试表"])
        assert result['success'] is True

        # 添加测试数据
        test_data = [
            ["姓名", "年龄", "部门"],
            ["张三", 25, "技术部"],
            ["李四", 30, "销售部"]
        ]
        result = excel_update_range(temp_path, "测试表!A1:C3", test_data)
        assert result['success'] is True

        yield temp_path

        # 清理临时文件
        try:
            os.unlink(temp_path)
        except FileNotFoundError:
            pass

    def test_excel_merge_cells_success(self, sample_excel_file):
        """测试合并单元格成功"""
        result = excel_merge_cells(sample_excel_file, "测试表", "A1:C1")

        assert result['success'] is True
        assert 'merged_range' in result['data']
        assert result['data']['merged_range'] == "A1:C1"
        assert result['data']['sheet_name'] == "测试表"

    def test_excel_unmerge_cells_success(self, sample_excel_file):
        """测试取消合并单元格成功"""
        # 先合并然后取消合并
        merge_result = excel_merge_cells(sample_excel_file, "测试表", "A1:C1")
        assert merge_result['success'] is True

        result = excel_unmerge_cells(sample_excel_file, "测试表", "A1:C1")

        assert result['success'] is True
        assert 'unmerged_range' in result['data']
        assert result['data']['unmerged_range'] == "A1:C1"
        assert result['data']['sheet_name'] == "测试表"

    def test_excel_set_borders_success(self, sample_excel_file):
        """测试设置边框成功"""
        result = excel_set_borders(sample_excel_file, "测试表", "A1:C3", "thick")

        assert result['success'] is True
        assert 'border_style' in result['data']
        assert result['data']['border_style'] == "thick"
        assert result['data']['range'] == "A1:C3"
        assert result['data']['cell_count'] == 9  # 3x3 区域

    def test_excel_set_row_height_success(self, sample_excel_file):
        """测试设置行高成功"""
        result = excel_set_row_height(sample_excel_file, "测试表", 1, 25.0)

        assert result['success'] is True
        assert 'row_number' in result['data']
        assert result['data']['row_number'] == 1
        assert result['data']['height'] == 25.0
        assert result['data']['sheet_name'] == "测试表"

    def test_excel_set_column_width_success(self, sample_excel_file):
        """测试设置列宽成功"""
        result = excel_set_column_width(sample_excel_file, "测试表", 1, 15.0)

        assert result['success'] is True
        assert 'column' in result['data']
        assert result['data']['column'] == "A"  # 列索引1对应A列
        assert result['data']['width'] == 15.0
        assert result['data']['sheet_name'] == "测试表"

    def test_excel_get_file_info_success(self, sample_excel_file):
        """测试获取文件信息成功"""
        result = excel_get_file_info(sample_excel_file)

        assert result['success'] is True
        assert 'file_size' in result['data']
        assert 'sheet_count' in result['data']
        assert 'format' in result['data']
        assert result['data']['format'] == 'xlsx'
        assert result['data']['sheet_count'] == 1
        assert "测试表" in result['data']['sheet_names']

    def test_excel_merge_cells_invalid_sheet(self, sample_excel_file):
        """测试合并单元格 - 工作表不存在"""
        result = excel_merge_cells(sample_excel_file, "不存在的表", "A1:C1")

        assert result['success'] is False
        assert 'error' in result

    def test_excel_set_borders_invalid_range(self, sample_excel_file):
        """测试设置边框 - 无效范围"""
        result = excel_set_borders(sample_excel_file, "测试表", "invalid_range", "thin")

        assert result['success'] is False
        assert 'error' in result

    def test_excel_set_row_height_invalid_sheet(self, sample_excel_file):
        """测试设置行高 - 工作表不存在"""
        result = excel_set_row_height(sample_excel_file, "不存在的表", 1, 20.0)

        assert result['success'] is False
        assert 'error' in result

    def test_excel_set_column_width_invalid_sheet(self, sample_excel_file):
        """测试设置列宽 - 工作表不存在"""
        result = excel_set_column_width(sample_excel_file, "不存在的表", 1, 10.0)

        assert result['success'] is False
        assert 'error' in result

    def test_excel_get_file_info_nonexistent_file(self):
        """测试获取文件信息 - 文件不存在"""
        result = excel_get_file_info("nonexistent_file.xlsx")

        assert result['success'] is False
        assert 'error' in result

    def test_excel_merge_cells_with_sheet_in_range(self, sample_excel_file):
        """测试合并单元格 - 范围表达式包含工作表名"""
        result = excel_merge_cells(sample_excel_file, None, "测试表!A1:B1")

        assert result['success'] is True
        assert result['data']['sheet_name'] == "测试表"

    def test_excel_borders_different_styles(self, sample_excel_file):
        """测试不同边框样式"""
        styles = ["thin", "thick", "double", "dashed", "dotted"]

        for style in styles:
            result = excel_set_borders(sample_excel_file, "测试表", "A1:A1", style)
            assert result['success'] is True
            assert result['data']['border_style'] == style

    def test_multiple_operations_sequence(self, sample_excel_file):
        """测试多个操作的序列组合"""
        # 1. 设置行高
        result1 = excel_set_row_height(sample_excel_file, "测试表", 1, 30.0)
        assert result1['success'] is True

        # 2. 设置列宽
        result2 = excel_set_column_width(sample_excel_file, "测试表", 1, 20.0)
        assert result2['success'] is True

        # 3. 合并单元格
        result3 = excel_merge_cells(sample_excel_file, "测试表", "A1:C1")
        assert result3['success'] is True

        # 4. 设置边框
        result4 = excel_set_borders(sample_excel_file, "测试表", "A1:C3", "thick")
        assert result4['success'] is True

        # 5. 取消合并
        result5 = excel_unmerge_cells(sample_excel_file, "测试表", "A1:C1")
        assert result5['success'] is True

        # 验证文件状态
        file_info = excel_get_file_info(sample_excel_file)
        assert file_info['success'] is True
