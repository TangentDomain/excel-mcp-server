"""
Excel MCP Server - 数据类型定义

定义了项目中使用的所有数据类型和模型
"""

from typing import Any, Dict, List, Optional, Union
from dataclasses import dataclass
from enum import Enum


class RangeType(Enum):
    """范围类型枚举"""
    CELL_RANGE = "cell_range"
    ROW_RANGE = "row_range"
    COLUMN_RANGE = "column_range"
    SINGLE_ROW = "single_row"
    SINGLE_COLUMN = "single_column"


class MatchType(Enum):
    """搜索匹配类型枚举"""
    VALUE = "value"
    FORMULA = "formula"


@dataclass
class RangeInfo:
    """范围信息"""
    sheet_name: Optional[str]
    cell_range: str
    range_type: RangeType


@dataclass
class ExcelDimensions:
    """Excel范围维度信息"""
    rows: int
    columns: int
    start_row: int
    start_column: int


@dataclass
class CellInfo:
    """单元格信息"""
    coordinate: str
    value: Any
    data_type: Optional[str] = None
    number_format: Optional[str] = None
    font: Optional[str] = None
    fill: Optional[str] = None


@dataclass(frozen=True)
class SheetInfo:
    """工作表信息"""
    index: int
    name: str
    is_active: bool
    max_row: int
    max_column: int
    max_column_letter: str


@dataclass
class SearchMatch:
    """搜索匹配结果"""
    sheet: str
    cell: str
    value: Optional[str] = None
    formula: Optional[str] = None
    match: str = ""
    match_start: int = 0
    match_end: int = 0
    match_type: MatchType = MatchType.VALUE


@dataclass
class OperationResult:
    """操作结果"""
    success: bool
    message: Optional[str] = None
    error: Optional[str] = None
    data: Optional[Any] = None
    metadata: Optional[Dict[str, Any]] = None


@dataclass
class ModifiedCell:
    """修改的单元格信息"""
    coordinate: str
    old_value: Any
    new_value: Any


# 类型别名
ExcelData = List[List[CellInfo]]
RawExcelData = List[List[Any]]
