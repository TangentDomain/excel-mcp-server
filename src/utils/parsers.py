"""
Excel MCP Server - 解析工具

提供范围表达式等解析功能
"""

import re
from typing import Dict, Any

from ..models.types import RangeInfo, RangeType
from .exceptions import InvalidRangeError


class RangeParser:
    """范围解析器"""

    @classmethod
    def parse_range_expression(cls, range_expr: str) -> RangeInfo:
        """
        解析范围表达式 (如 'Sheet1!A1:C10' 或 'A1:C10' 或 '1:1' 或 'A:A')

        Args:
            range_expr: 范围表达式

        Returns:
            RangeInfo对象

        Raises:
            InvalidRangeError: 无效的范围表达式
        """
        if not range_expr or not range_expr.strip():
            raise InvalidRangeError("范围表达式不能为空")

        range_expr = range_expr.strip()

        # 分离工作表名和单元格范围
        if '!' in range_expr:
            sheet_name, cell_range = range_expr.split('!', 1)
        else:
            sheet_name = None
            cell_range = range_expr

        # 检测范围类型
        range_type = cls._detect_range_type(cell_range)

        # 规范化范围表达式
        normalized_range = cls._normalize_range(cell_range, range_type)

        return RangeInfo(
            sheet_name=sheet_name,
            cell_range=normalized_range,
            range_type=range_type
        )

    @classmethod
    def _detect_range_type(cls, cell_range: str) -> RangeType:
        """
        检测范围类型

        Args:
            cell_range: 单元格范围字符串

        Returns:
            范围类型

        Raises:
            InvalidRangeError: 无效的范围表达式
        """
        # 检测整行模式 (如 "1:1", "3:5")
        if re.match(r'^\d+:\d+$', cell_range):
            return RangeType.ROW_RANGE

        # 检测整列模式 (如 "A:A", "B:D")
        elif re.match(r'^[A-Z]+:[A-Z]+$', cell_range):
            return RangeType.COLUMN_RANGE

        # 检测单行模式 (如 "1", 只读取第1行)
        elif re.match(r'^\d+$', cell_range):
            return RangeType.SINGLE_ROW

        # 检测单列模式 (如 "A", 只读取A列)
        elif re.match(r'^[A-Z]+$', cell_range):
            return RangeType.SINGLE_COLUMN

        # 检测单元格范围模式 (如 "A1:B2", "A1")
        elif re.match(r'^[A-Z]+\d+$|^[A-Z]+\d+:[A-Z]+\d+$', cell_range):
            return RangeType.CELL_RANGE

        # 如果都不匹配，抛出异常
        else:
            raise InvalidRangeError(f"无效的范围表达式: {cell_range}")

    @classmethod
    def _normalize_range(cls, cell_range: str, range_type: RangeType) -> str:
        """
        规范化范围表达式

        Args:
            cell_range: 原始范围表达式
            range_type: 范围类型

        Returns:
            规范化的范围表达式
        """
        if range_type == RangeType.SINGLE_ROW:
            return f"{cell_range}:{cell_range}"
        elif range_type == RangeType.SINGLE_COLUMN:
            return f"{cell_range}:{cell_range}"
        else:
            return cell_range

    @classmethod
    def validate_range_syntax(cls, range_expr: str) -> bool:
        """
        验证范围表达式语法

        Args:
            range_expr: 范围表达式

        Returns:
            是否为有效语法
        """
        try:
            cls.parse_range_expression(range_expr)
            return True
        except InvalidRangeError:
            return False
