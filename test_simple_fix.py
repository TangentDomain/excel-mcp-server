#!/usr/bin/env python3
"""
简化的根源性测试
"""

import json
from dataclasses import dataclass
from typing import Optional
from enum import Enum

# 复制核心数据类型定义
class DifferenceType(Enum):
    CELL_CHANGED = "cell_changed"
    SHEET_ADDED = "sheet_added"
    SHEET_REMOVED = "sheet_removed"

@dataclass
class CellDifference:
    coordinate: str
    difference_type: DifferenceType
    old_value: Optional[str] = None
    new_value: Optional[str] = None
    old_format: Optional[dict] = None
    new_format: Optional[dict] = None
    sheet_name: Optional[str] = None

@dataclass
class SheetComparison:
    sheet_name: str
    exists_in_file1: bool = True
    exists_in_file2: bool = True
    differences: list[CellDifference] = None
    total_differences: int = 0
    structural_changes: Optional[dict] = None

    def __post_init__(self):
        if self.differences is None:
            self.differences = []

@dataclass
class OperationResult:
    success: bool
    message: Optional[str] = None
    error: Optional[str] = None
    data: Optional[any] = None
    metadata: Optional[dict] = None

def _deep_clean_nulls(obj):
    """递归深度清理对象中的null/None值"""
    if isinstance(obj, dict):
        # 过滤字典中的None值，并递归清理剩余值
        cleaned = {}
        for key, value in obj.items():
            if value is not None:
                cleaned_value = _deep_clean_nulls(value)
                if cleaned_value is not None and cleaned_value != {} and cleaned_value != []:
                    cleaned[key] = cleaned_value
        return cleaned
    elif isinstance(obj, list):
        # 清理列表中的None值，并递归清理剩余元素
        cleaned = []
        for item in obj:
            if item is not None:
                cleaned_item = _deep_clean_nulls(item)
                if cleaned_item is not None and cleaned_item != {} and cleaned_item != []:
                    cleaned.append(cleaned_item)
        return cleaned
    elif hasattr(obj, '__dict__'):
        # 处理数据类对象，递归转换为字典并清理
        return _deep_clean_nulls(obj.__dict__)
    else:
        # 基础类型直接返回
        return obj

def _format_result(result) -> dict:
    """格式化操作结果为MCP响应格式，彻底清理所有null/None值"""
    response = {
        'success': result.success,
    }

    if result.success:
        if result.data is not None:
            # 处理数据类型转换
            if hasattr(result.data, '__dict__'):
                # 如果是数据类，转换为字典并清理
                data_dict = result.data.__dict__
                cleaned_data = _deep_clean_nulls(data_dict)
                response.update(cleaned_data)
            elif isinstance(result.data, list):
                # 如果是列表，处理每个元素并清理
                raw_data = [
                    item.__dict__ if hasattr(item, '__dict__') else item
                    for item in result.data
                ]
                cleaned_data = _deep_clean_nulls(raw_data)
                if cleaned_data:  # 只在有数据时添加
                    response['data'] = cleaned_data
            else:
                response['data'] = result.data

        if result.metadata:
            cleaned_metadata = _deep_clean_nulls(result.metadata)
            response.update(cleaned_metadata)

        if result.message:
            response['message'] = result.message
    else:
        response['error'] = result.error

    # 最终清理整个响应对象
    return _deep_clean_nulls(response)

def test_simple():
    """简化测试"""
    print("=== 简化测试根源性修复 ===")

    # 创建测试数据
    cell_diff1 = CellDifference(
        coordinate="A1",
        difference_type=DifferenceType.CELL_CHANGED,
        old_value="旧值",
        new_value="新值"
        # old_format, new_format, sheet_name 使用默认None
    )

    cell_diff2 = CellDifference(
        coordinate="SHEET",
        difference_type=DifferenceType.SHEET_ADDED,
        sheet_name="新工作表"
        # 其他字段使用默认None
    )

    sheet_comp = SheetComparison(
        sheet_name="测试工作表",
        differences=[cell_diff1, cell_diff2],
        total_differences=2
        # structural_changes 使用默认None
    )

    result = OperationResult(
        success=True,
        message="成功比较Excel文件",
        data=sheet_comp,
        metadata={
            "total_differences": 2,
            "empty_metadata": None,
            "null_list": [],
            "nested_null": {"valid_field": "有效值", "null_field": None}
        }
    )

    # 格式化并检查结果
    formatted_result = _format_result(result)
    json_str = json.dumps(formatted_result, ensure_ascii=False, indent=2)

    print(f"JSON长度: {len(json_str)}")
    null_count = json_str.count('null')
    print(f"JSON中null的数量: {null_count}")

    if null_count == 0:
        print("✅ 根源性修复成功！完全清除了null值")
    else:
        print("❌ 仍有null值存在")

    print("\n=== 格式化结果 ===")
    print(json_str)

if __name__ == "__main__":
    test_simple()
