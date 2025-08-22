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


class DifferenceType(Enum):
    """比较差异类型枚举"""
    VALUE_CHANGED = "value_changed"      # 单元格值变化
    FORMAT_CHANGED = "format_changed"    # 单元格格式变化
    CELL_ADDED = "cell_added"           # 新增单元格
    CELL_REMOVED = "cell_removed"       # 删除单元格
    STRUCTURE_CHANGED = "structure_changed"  # 结构变化（行列数）
    SHEET_ADDED = "sheet_added"         # 新增工作表
    SHEET_REMOVED = "sheet_removed"     # 删除工作表
    SHEET_RENAMED = "sheet_renamed"     # 工作表重命名
    # 结构化数据差异类型
    ROW_ADDED = "row_added"             # 新增行
    ROW_REMOVED = "row_removed"         # 删除行
    ROW_MODIFIED = "row_modified"       # 行数据修改
    HEADER_CHANGED = "header_changed"   # 表头变化


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


@dataclass
class CellDifference:
    """单元格差异信息"""
    coordinate: str
    difference_type: DifferenceType
    old_value: Optional[Any] = None
    new_value: Optional[Any] = None
    old_format: Optional[str] = None
    new_format: Optional[str] = None
    sheet_name: Optional[str] = None


@dataclass
class SheetComparison:
    """工作表比较结果"""
    sheet_name: str
    exists_in_file1: bool
    exists_in_file2: bool
    differences: List[CellDifference]
    total_differences: int
    structural_changes: Dict[str, Any]  # 结构变化信息


@dataclass
class ComparisonResult:
    """Excel比较结果"""
    file1_path: str
    file2_path: str
    identical: bool
    total_differences: int
    sheet_comparisons: List[SheetComparison]
    structural_differences: Dict[str, Any]  # 文件级别的结构差异
    summary: str  # 比较结果摘要


@dataclass
class ComparisonOptions:
    """比较选项配置"""
    compare_values: bool = True      # 比较单元格值
    compare_formulas: bool = False   # 比较公式
    compare_formats: bool = False    # 比较格式
    ignore_empty_cells: bool = True  # 忽略空单元格
    case_sensitive: bool = True      # 大小写敏感
    # 表格化数据比较选项（游戏开发友好）
    header_row: Optional[int] = 1         # 表头行号（1-based），默认第一行
    id_column: Optional[Union[int, str]] = 1  # ID列位置（1-based或列名），默认第一列
    structured_comparison: bool = True    # 默认启用结构化数据比较
    show_numeric_changes: bool = True     # 显示数值变化量和百分比
    game_friendly_format: bool = True     # 游戏开发友好的输出格式


@dataclass 
class RowDifference:
    """行级差异信息"""
    row_id: Any                     # 行的唯一标识
    difference_type: DifferenceType # 差异类型：行增加、删除、修改
    row_data1: Optional[Dict[str, Any]] = None  # 第一个文件中的行数据
    row_data2: Optional[Dict[str, Any]] = None  # 第二个文件中的行数据
    field_differences: Optional[List[str]] = None  # 字段级差异列表
    row_index1: Optional[int] = None # 在第一个文件中的行号
    row_index2: Optional[int] = None # 在第二个文件中的行号


@dataclass
class StructuredSheetComparison:
    """结构化工作表比较结果"""
    sheet_name: str
    exists_in_file1: bool
    exists_in_file2: bool
    headers1: Optional[List[str]] = None  # 第一个文件的表头
    headers2: Optional[List[str]] = None  # 第二个文件的表头
    header_differences: Optional[List[str]] = None  # 表头差异
    row_differences: List[RowDifference] = None  # 行级差异
    total_differences: int = 0
    identical_rows: int = 0         # 完全相同的行数
    modified_rows: int = 0          # 修改的行数
    added_rows: int = 0             # 新增的行数
    removed_rows: int = 0           # 删除的行数


# 类型别名
ExcelData = List[List[CellInfo]]
RawExcelData = List[List[Any]]
ComparisonData = Dict[str, Any]  # 比较数据类型别名
