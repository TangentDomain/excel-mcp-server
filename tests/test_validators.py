"""
Excel MCP Server - 验证器测试

测试utils.validators模块的所有功能
"""

import pytest
from pathlib import Path

from excel_mcp.utils.validators import ExcelValidator
from excel_mcp.utils.exceptions import (
    FileNotFoundError, InvalidFormatError, DataValidationError
)


class TestExcelValidator:
    """ExcelValidator类的测试"""

    def test_validate_file_path_success(self, sample_xlsx_file):
        """测试有效文件路径验证"""
        result = ExcelValidator.validate_file_path(sample_xlsx_file)
        assert result == str(Path(sample_xlsx_file).absolute())

    def test_validate_file_path_not_found(self, nonexistent_file_path):
        """测试文件不存在的情况"""
        with pytest.raises(FileNotFoundError) as exc_info:
            ExcelValidator.validate_file_path(nonexistent_file_path)
        assert "Excel文件不存在" in str(exc_info.value)

    def test_validate_file_path_invalid_format(self, invalid_format_file):
        """测试无效文件格式"""
        with pytest.raises(InvalidFormatError) as exc_info:
            ExcelValidator.validate_file_path(invalid_format_file)
        assert "不支持的文件格式" in str(exc_info.value)

    def test_validate_sheet_name_valid(self):
        """测试有效工作表名称"""
        ExcelValidator.validate_sheet_name("Sheet1")  # 不应抛出异常
        ExcelValidator.validate_sheet_name("  Valid Name  ")  # 不应抛出异常
        ExcelValidator.validate_sheet_name(None)  # None是允许的

    def test_validate_sheet_name_invalid(self):
        """测试无效工作表名称"""
        with pytest.raises(DataValidationError):
            ExcelValidator.validate_sheet_name("")

        with pytest.raises(DataValidationError):
            ExcelValidator.validate_sheet_name("   ")

    def test_validate_row_operations_success(self):
        """测试有效行操作参数"""
        ExcelValidator.validate_row_operations(1, 1)  # 最小值
        ExcelValidator.validate_row_operations(100, 50)  # 正常值
        ExcelValidator.validate_row_operations(1, 1000)  # 最大值

    def test_validate_row_operations_invalid_row_index(self):
        """测试无效行索引"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_row_operations(0, 1)
        assert "行索引必须大于等于1" in str(exc_info.value)

        with pytest.raises(DataValidationError):
            ExcelValidator.validate_row_operations(-1, 1)

    def test_validate_row_operations_invalid_count(self):
        """测试无效行数"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_row_operations(1, 0)
        assert "操作行数必须大于等于1" in str(exc_info.value)

        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_row_operations(1, 1001)
        assert "一次最多操作1000行" in str(exc_info.value)

    def test_validate_column_operations_success(self):
        """测试有效列操作参数"""
        ExcelValidator.validate_column_operations(1, 1)  # 最小值
        ExcelValidator.validate_column_operations(10, 10)  # 正常值
        ExcelValidator.validate_column_operations(1, 100)  # 最大值

    def test_validate_column_operations_invalid_column_index(self):
        """测试无效列索引"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_column_operations(0, 1)
        assert "列索引必须大于等于1" in str(exc_info.value)

        with pytest.raises(DataValidationError):
            ExcelValidator.validate_column_operations(-1, 1)

    def test_validate_column_operations_invalid_count(self):
        """测试无效列数"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_column_operations(1, 0)
        assert "操作列数必须大于等于1" in str(exc_info.value)

        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_column_operations(1, 101)
        assert "一次最多操作100列" in str(exc_info.value)

    def test_validate_range_data_success(self):
        """测试有效范围数据"""
        data = [
            ["A", "B"],
            ["C", "D"]
        ]
        ExcelValidator.validate_range_data(data, 2, 2)  # 完全匹配
        ExcelValidator.validate_range_data(data, 3, 3)  # 范围更大
        ExcelValidator.validate_range_data([["A"]], 2, 2)  # 数据更小

    def test_validate_range_data_too_many_rows(self):
        """测试数据行数超出范围"""
        data = [
            ["A", "B"],
            ["C", "D"],
            ["E", "F"]
        ]
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_data(data, 2, 2)
        assert "数据行数(3)超过范围行数(2)" in str(exc_info.value)

    def test_validate_range_data_too_many_columns(self):
        """测试数据列数超出范围"""
        data = [
            ["A", "B", "C"]
        ]
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_data(data, 2, 2)
        assert "第1行数据列数(3)超过范围列数(2)" in str(exc_info.value)

    def test_validate_file_for_creation_success(self, temp_dir):
        """测试创建文件路径验证成功"""
        new_file_path = temp_dir / "new_file.xlsx"
        result = ExcelValidator.validate_file_for_creation(str(new_file_path))
        assert result == str(new_file_path.absolute())

        # 测试xlsm格式
        new_file_path_xlsm = temp_dir / "new_file.xlsm"
        result = ExcelValidator.validate_file_for_creation(str(new_file_path_xlsm))
        assert result == str(new_file_path_xlsm.absolute())

    def test_validate_file_for_creation_file_exists(self, sample_xlsx_file):
        """测试文件已存在的情况"""
        with pytest.raises(FileExistsError) as exc_info:
            ExcelValidator.validate_file_for_creation(sample_xlsx_file)
        assert "文件已存在" in str(exc_info.value)

    def test_validate_file_for_creation_invalid_format(self, temp_dir):
        """测试创建时使用无效格式"""
        invalid_path = temp_dir / "new_file.txt"
        with pytest.raises(InvalidFormatError) as exc_info:
            ExcelValidator.validate_file_for_creation(str(invalid_path))
        assert "不支持的文件格式" in str(exc_info.value)
        assert "请使用 .xlsx 或 .xlsm" in str(exc_info.value)
