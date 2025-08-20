"""
Excel MCP Server - 解析器测试

测试utils.parsers模块的所有功能
"""

import pytest
from excel_mcp.utils.parsers import RangeParser
from excel_mcp.models.types import RangeType
from excel_mcp.utils.exceptions import InvalidRangeError


class TestRangeParser:
    """RangeParser类的测试"""

    def test_parse_cell_range(self):
        """测试单元格范围解析"""
        result = RangeParser.parse_range_expression("A1:C10")
        assert result.sheet_name is None
        assert result.cell_range == "A1:C10"
        assert result.range_type == RangeType.CELL_RANGE

    def test_parse_with_sheet_name(self):
        """测试带工作表名的范围解析"""
        result = RangeParser.parse_range_expression("Sheet1!A1:C10")
        assert result.sheet_name == "Sheet1"
        assert result.cell_range == "A1:C10"
        assert result.range_type == RangeType.CELL_RANGE

    def test_parse_row_range(self):
        """测试行范围解析"""
        result = RangeParser.parse_range_expression("1:5")
        assert result.range_type == RangeType.ROW_RANGE
        assert result.cell_range == "1:5"

    def test_parse_single_row(self):
        """测试单行解析"""
        result = RangeParser.parse_range_expression("3")
        assert result.range_type == RangeType.SINGLE_ROW
        assert result.cell_range == "3:3"

    def test_parse_column_range(self):
        """测试列范围解析"""
        result = RangeParser.parse_range_expression("A:C")
        assert result.range_type == RangeType.COLUMN_RANGE
        assert result.cell_range == "A:C"

    def test_parse_single_column(self):
        """测试单列解析"""
        result = RangeParser.parse_range_expression("B")
        assert result.range_type == RangeType.SINGLE_COLUMN
        assert result.cell_range == "B:B"

    def test_parse_empty_expression(self):
        """测试空表达式"""
        with pytest.raises(InvalidRangeError):
            RangeParser.parse_range_expression("")

        with pytest.raises(InvalidRangeError):
            RangeParser.parse_range_expression("   ")

    def test_validate_range_syntax(self):
        """测试范围语法验证"""
        assert RangeParser.validate_range_syntax("A1:C10")
        assert RangeParser.validate_range_syntax("Sheet1!A1:C10")
        assert RangeParser.validate_range_syntax("1:5")
        assert RangeParser.validate_range_syntax("A:C")

        assert not RangeParser.validate_range_syntax("")
        assert not RangeParser.validate_range_syntax("invalid")
