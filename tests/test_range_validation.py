"""
测试范围表达式验证功能

验证新增的严格范围格式验证功能
"""

import pytest
from src.utils.validators import ExcelValidator, DataValidationError


class TestRangeValidation:
    """测试范围表达式验证功能"""

    def test_valid_cell_range(self):
        """测试有效的单元格范围"""
        # 标准范围
        result = ExcelValidator.validate_range_expression("Sheet1!A1:C10")
        assert result['success'] is True
        assert result['sheet_name'] == "Sheet1"
        assert result['range_part'] == "A1:C10"
        assert result['range_info']['type'] == 'cell_range'
        assert result['range_info']['start_cell'] == 'A1'
        assert result['range_info']['end_cell'] == 'C10'

    def test_valid_row_range(self):
        """测试有效的行范围"""
        result = ExcelValidator.validate_range_expression("数据!1:10")
        assert result['success'] is True
        assert result['sheet_name'] == "数据"
        assert result['range_info']['type'] == 'row_range'
        assert result['range_info']['start_row'] == 1
        assert result['range_info']['end_row'] == 10

    def test_valid_column_range(self):
        """测试有效的列范围"""
        result = ExcelValidator.validate_range_expression("Test!A:C")
        assert result['success'] is True
        assert result['sheet_name'] == "Test"
        assert result['range_info']['type'] == 'column_range'
        assert result['range_info']['start_col'] == 1
        assert result['range_info']['end_col'] == 3

    def test_valid_single_cell(self):
        """测试单个单元格"""
        result = ExcelValidator.validate_range_expression("Sheet1!A1")
        assert result['success'] is True
        assert result['range_info']['type'] == 'single_cell'
        assert result['range_info']['column'] == 1
        assert result['range_info']['row'] == 1

    def test_valid_single_row(self):
        """测试单行"""
        result = ExcelValidator.validate_range_expression("Sheet1!5")
        assert result['success'] is True
        assert result['range_info']['type'] == 'single_row'
        assert result['range_info']['row'] == 5

    def test_valid_single_column(self):
        """测试单列"""
        result = ExcelValidator.validate_range_expression("Sheet1!A")
        assert result['success'] is True
        assert result['range_info']['type'] == 'single_column'
        assert result['range_info']['column'] == 1

    def test_chinese_sheet_name(self):
        """测试中文工作表名"""
        result = ExcelValidator.validate_range_expression("技能配置表!A1:Z100")
        assert result['success'] is True
        assert result['sheet_name'] == "技能配置表"
        assert result['normalized_range'] == "技能配置表!A1:Z100"

    def test_invalid_range_missing_sheet(self):
        """测试缺少工作表名的无效范围"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_expression("A1:C10")

        assert "必须包含工作表名和感叹号" in str(exc_info.value)

    def test_invalid_range_multiple_exclamation(self):
        """测试多个感叹号的无效范围"""
        # 当有多个感叹号时，会被解析为工作表名包含感叹号，然后范围部分会无效
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_expression("Sheet1!!A1:C10")

        # 这种情况会变成工作表名为"Sheet1!"，范围部分为"A1:C10"，实际会通过验证
        # 但我们测试的是真正无效的范围格式
        assert "无法识别的范围格式" in str(exc_info.value) or "工作表名包含无效字符" in str(exc_info.value)

    def test_invalid_empty_sheet_name(self):
        """测试空工作表名"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_expression("!A1:C10")

        assert "工作表名不能为空" in str(exc_info.value)

    def test_invalid_sheet_name_with_special_chars(self):
        """测试包含特殊字符的工作表名"""
        invalid_chars = ['[', ']', '*', ':', '?', '/', '\\']
        for char in invalid_chars:
            with pytest.raises(DataValidationError) as exc_info:
                ExcelValidator.validate_range_expression(f"Sheet{char}1!A1:C10")

            assert "工作表名包含无效字符" in str(exc_info.value)

    def test_invalid_sheet_name_too_long(self):
        """测试过长的工作表名"""
        long_name = "A" * 32  # 超过31个字符
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_expression(f"{long_name}!A1:C10")

        assert "工作表名长度不能超过31个字符" in str(exc_info.value)

    def test_invalid_empty_range_part(self):
        """测试空的范围部分"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_expression("Sheet1!")

        assert "范围部分不能为空" in str(exc_info.value)

    def test_invalid_range_format(self):
        """测试无效的范围格式"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_expression("Sheet1!INVALID_RANGE")

        assert "无法识别的范围格式" in str(exc_info.value)

    def test_invalid_null_or_empty(self):
        """测试空值或None"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_expression("")

        assert "范围表达式不能为空" in str(exc_info.value)

        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_expression(None)

        assert "范围表达式不能为空" in str(exc_info.value)

    def test_invalid_non_string(self):
        """测试非字符串类型"""
        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_range_expression(123)

        assert "范围表达式不能为空且必须是字符串" in str(exc_info.value)

    def test_case_insensitive_columns(self):
        """测试列名大小写不敏感"""
        result1 = ExcelValidator.validate_range_expression("Sheet1!a1:c10")
        result2 = ExcelValidator.validate_range_expression("Sheet1!A1:C10")

        assert result1['range_info']['start_col'] == result2['range_info']['start_col']
        assert result1['range_info']['end_col'] == result2['range_info']['end_col']

    def test_column_conversion(self):
        """测试列字母转换"""
        # 测试简单列
        assert ExcelValidator._col_to_num("A") == 1
        assert ExcelValidator._col_to_num("B") == 2
        assert ExcelValidator._col_to_num("Z") == 26

        # 测试多列
        assert ExcelValidator._col_to_num("AA") == 27
        assert ExcelValidator._col_to_num("AB") == 28
        assert ExcelValidator._col_to_num("AZ") == 52

        # 测试三列
        assert ExcelValidator._col_to_num("AAA") == 703

    def test_operation_scale_validation_low_risk(self):
        """测试低风险操作规模验证"""
        range_info = {
            'type': 'cell_range',
            'start_col': 1,
            'end_col': 10,
            'start_row': 1,
            'end_row': 10
        }

        result = ExcelValidator.validate_operation_scale(range_info)
        assert result['rows'] == 10
        assert result['columns'] == 10
        assert result['total_cells'] == 100
        assert result['risk_level'] == "LOW"
        assert result['warning'] is None
        assert result['within_limits'] is True

    def test_operation_scale_validation_medium_risk(self):
        """测试中等风险操作规模验证"""
        range_info = {
            'type': 'cell_range',
            'start_col': 1,
            'end_col': 20,
            'start_row': 1,
            'end_row': 100
        }

        result = ExcelValidator.validate_operation_scale(range_info)
        assert result['total_cells'] == 2000
        assert result['risk_level'] == "MEDIUM"
        assert "中等风险操作" in result['warning']
        assert result['within_limits'] is True

    def test_operation_scale_validation_high_risk(self):
        """测试高风险操作规模验证"""
        range_info = {
            'type': 'cell_range',
            'start_col': 1,
            'end_col': 50,
            'start_row': 1,
            'end_row': 300
        }

        result = ExcelValidator.validate_operation_scale(range_info)
        assert result['total_cells'] == 15000
        assert result['risk_level'] == "HIGH"
        assert "高风险操作" in result['warning']
        assert result['within_limits'] is True

    def test_operation_scale_validation_rows_exceeded(self):
        """测试行数超限"""
        range_info = {
            'type': 'cell_range',
            'start_col': 1,
            'end_col': 10,
            'start_row': 1,
            'end_row': 2000  # 超过MAX_ROWS_OPERATION
        }

        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_operation_scale(range_info)

        assert "超过限制" in str(exc_info.value)

    def test_operation_scale_validation_columns_exceeded(self):
        """测试列数超限"""
        range_info = {
            'type': 'cell_range',
            'start_col': 1,
            'end_col': 200,  # 超过MAX_COLUMNS_OPERATION
            'start_row': 1,
            'end_row': 10
        }

        with pytest.raises(DataValidationError) as exc_info:
            ExcelValidator.validate_operation_scale(range_info)

        assert "超过限制" in str(exc_info.value)

    def test_single_row_scale_validation(self):
        """测试单行操作规模验证"""
        range_info = {
            'type': 'single_row',
            'row': 5
        }

        result = ExcelValidator.validate_operation_scale(range_info)
        assert result['rows'] == 1
        assert result['columns'] == 1
        assert result['total_cells'] == 1
        assert result['risk_level'] == "LOW"

    def test_single_column_scale_validation(self):
        """测试单列操作规模验证"""
        range_info = {
            'type': 'single_column',
            'column': 5
        }

        result = ExcelValidator.validate_operation_scale(range_info)
        assert result['rows'] == 1
        assert result['columns'] == 1
        assert result['total_cells'] == 1
        assert result['risk_level'] == "LOW"