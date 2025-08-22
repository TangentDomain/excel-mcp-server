#!/usr/bin/env python3
"""
调试紧凑数组格式转换问题
"""

import sys
from pathlib import Path
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src"))

from utils.formatter import format_operation_result

@dataclass
class OperationResult:
    success: bool
    message: Optional[str] = None
    error: Optional[str] = None
    data: Optional[Any] = None
    metadata: Optional[Dict[str, Any]] = None

@dataclass  
class MockRowDifference:
    row_id: Any
    difference_type: str
    row_index1: int
    row_index2: int
    sheet_name: str
    detailed_field_differences: Optional[List] = None

@dataclass
class MockStructuredDataComparison:
    sheet_name: str
    exists_in_file1: bool
    exists_in_file2: bool
    total_differences: int
    row_differences: List[MockRowDifference]

# 测试无字段差异的行处理
row_diff2 = MockRowDifference(
    row_id="1002",
    difference_type="row_added", 
    row_index1=0,
    row_index2=8,
    sheet_name="TrSkill",
    detailed_field_differences=None  # 显式设为None
)

structured_data = MockStructuredDataComparison(
    sheet_name="TrSkill比较",
    exists_in_file1=True,
    exists_in_file2=True,
    total_differences=1,
    row_differences=[row_diff2]
)

result = OperationResult(success=True, data=structured_data)
formatted = format_operation_result(result)

print("格式化后的数据:")
print(formatted["data"]["row_differences"])

print("\n第二行数据:")
second_row = formatted["data"]["row_differences"][1]
print(f"行数据: {second_row}")
print(f"行长度: {len(second_row)}")
print(f"各字段: {[f'{i}: {val}' for i, val in enumerate(second_row)]}")
