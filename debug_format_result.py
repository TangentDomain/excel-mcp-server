#!/usr/bin/env python3
"""
_format_result 方法调试脚本

用于调试和理解 _format_result 函数的实际行为
"""

import json
from enum import Enum
from dataclasses import dataclass
from typing import Any, Dict, List, Optional


# ==================== 复制 _format_result 函数 ====================

def _format_result(result) -> Dict[str, Any]:
    """
    格式化操作结果为MCP响应格式，使用JSON序列化简化方案

    Args:
        result: OperationResult对象

    Returns:
        格式化后的字典，已清理null值，并转换为紧凑数组格式
    """
    import json

    def _convert_to_compact_array_format(data):
        """
        将结构化比较结果转换为紧凑的数组格式

        Args:
            data: StructuredDataComparison 数据对象

        Returns:
            转换后的紧凑格式数据
        """
        if not isinstance(data, dict) or 'row_differences' not in data:
            return data

        row_differences = data.get('row_differences', [])
        if not row_differences:
            return data

        # 检查是否已经是数组格式（避免重复转换）
        if (isinstance(row_differences, list) and
            len(row_differences) > 0 and
            isinstance(row_differences[0], list) and
            len(row_differences[0]) == 6 and
            row_differences[0] == ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"]):
            return data

        # 转换为紧凑数组格式
        compact_differences = []

        # 第一行：字段定义
        field_definitions = ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"]
        compact_differences.append(field_definitions)

        # 后续行：实际数据
        for diff in row_differences:
            if isinstance(diff, dict):
                # 转换字段级差异为数组格式
                field_diffs = diff.get('detailed_field_differences', [])
                compact_field_diffs = None

                if field_diffs:
                    compact_field_diffs = []
                    for field_diff in field_diffs:
                        if isinstance(field_diff, dict):
                            # 数组格式：[field_name, old_value, new_value, change_type]
                            compact_field_diffs.append([
                                field_diff.get('field_name', ''),
                                field_diff.get('old_value', ''),
                                field_diff.get('new_value', ''),
                                field_diff.get('change_type', '')
                            ])

                # 主要差异数据数组：按字段定义顺序
                compact_row = [
                    diff.get('row_id', ''),
                    diff.get('difference_type', ''),
                    diff.get('row_index1', 0),
                    diff.get('row_index2', 0),
                    diff.get('sheet_name', ''),
                    compact_field_diffs
                ]
                compact_differences.append(compact_row)

        # 创建新的数据副本，替换row_differences
        new_data = data.copy()
        new_data['row_differences'] = compact_differences

        return new_data

    def _deep_clean_nulls(obj):
        """递归深度清理对象中的null/None值"""
        if isinstance(obj, dict):
            cleaned = {}
            for key, value in obj.items():
                if value is not None:
                    cleaned_value = _deep_clean_nulls(value)
                    if cleaned_value is not None and cleaned_value != {} and cleaned_value != []:
                        cleaned[key] = cleaned_value
            return cleaned
        elif isinstance(obj, list):
            cleaned = []
            for item in obj:
                if item is not None:
                    cleaned_item = _deep_clean_nulls(item)
                    if cleaned_item is not None and cleaned_item != {} and cleaned_item != []:
                        cleaned.append(cleaned_item)
            return cleaned
        else:
            return obj

    # 步骤1: 先转成JSON字符串（自动处理dataclass）
    try:
        def json_serializer(obj):
            """自定义JSON序列化器，专门处理dataclass和枚举"""
            if isinstance(obj, Enum):
                return obj.value
            elif hasattr(obj, '__dict__'):
                return obj.__dict__
            else:
                return str(obj)

        json_str = json.dumps(result, default=json_serializer, ensure_ascii=False)
        # 步骤2: 再转回字典
        result_dict = json.loads(json_str)

        # 步骤3: 转换为紧凑数组格式（仅用于结构化比较结果）
        if result_dict.get('data'):
            result_dict['data'] = _convert_to_compact_array_format(result_dict['data'])

        # 步骤4: 应用null清理
        cleaned_dict = _deep_clean_nulls(result_dict)
        return cleaned_dict
    except Exception as e:
        print(f"⚠️ JSON序列化失败: {str(e)}")
        # 如果JSON方案失败，回退到原始方案
        response = {
            'success': result.success,
        }

        if result.success:
            # 统一数据处理，避免重复
            if result.data is not None:
                if hasattr(result.data, '__dict__'):
                    # 如果是数据类，转换为字典并放在data字段中
                    response['data'] = result.data.__dict__
                elif isinstance(result.data, list):
                    # 如果是列表，处理每个元素并放在data字段中
                    response['data'] = [
                        item.__dict__ if hasattr(item, '__dict__') else item
                        for item in result.data
                    ]
                else:
                    response['data'] = result.data

            # 分离处理metadata，避免键冲突
            if result.metadata:
                response['metadata'] = result.metadata

            if result.message:
                response['message'] = result.message
        else:
            response['error'] = result.error

        return response


# ==================== 测试数据模型 ====================

@dataclass
class OperationResult:
    """操作结果"""
    success: bool
    message: Optional[str] = None
    error: Optional[str] = None
    data: Optional[Any] = None
    metadata: Optional[Dict[str, Any]] = None


class MockStatus(Enum):
    """测试用枚举"""
    ACTIVE = "active"
    INACTIVE = "inactive"


# ==================== 调试测试 ====================

def debug_basic_result():
    """调试基本结果格式化"""
    print("🔍 调试测试 1: 基本成功结果")
    print("-" * 50)

    result = OperationResult(
        success=True,
        message="操作成功",
        data={"test": "data", "count": 42},
        metadata={"timestamp": "2025-08-22"}
    )

    print("📥 输入:")
    print(f"   success: {result.success}")
    print(f"   message: {result.message}")
    print(f"   data: {result.data}")
    print(f"   metadata: {result.metadata}")

    formatted = _format_result(result)

    print("📤 输出:")
    print(json.dumps(formatted, indent=2, ensure_ascii=False))

    print("🔑 输出字段:")
    for key in formatted.keys():
        print(f"   - {key}: {formatted[key]}")


def debug_enum_result():
    """调试枚举结果格式化"""
    print("\n🔍 调试测试 2: 枚举类型序列化")
    print("-" * 50)

    result = OperationResult(
        success=True,
        data={
            "status": MockStatus.ACTIVE,
        }
    )

    print("📥 输入:")
    print(f"   success: {result.success}")
    print(f"   data: {result.data}")
    print(f"   data['status']: {result.data['status']} (type: {type(result.data['status'])})")

    formatted = _format_result(result)

    print("📤 输出:")
    print(json.dumps(formatted, indent=2, ensure_ascii=False))

    print("🔑 输出字段:")
    for key in formatted.keys():
        print(f"   - {key}: {formatted[key]}")


def debug_dataclass_result():
    """调试数据类结果格式化"""
    print("\n🔍 调试测试 3: 数据类序列化")
    print("-" * 50)

    @dataclass
    class TestData:
        name: str
        value: int

    test_obj = TestData(name="测试", value=123)

    result = OperationResult(
        success=True,
        data={"test_obj": test_obj}
    )

    print("📥 输入:")
    print(f"   success: {result.success}")
    print(f"   data: {result.data}")
    print(f"   test_obj: {test_obj} (type: {type(test_obj)})")
    print(f"   has __dict__: {hasattr(test_obj, '__dict__')}")
    print(f"   __dict__: {test_obj.__dict__}")

    formatted = _format_result(result)

    print("📤 输出:")
    print(json.dumps(formatted, indent=2, ensure_ascii=False))


def debug_fallback_mechanism():
    """调试回退机制"""
    print("\n🔍 调试测试 4: 回退机制")
    print("-" * 50)

    class UnserializableObject:
        def __str__(self):
            return "UnserializableObject"

        def __init__(self):
            # 创建一个不可序列化的循环引用
            self.ref = self

    unserializable = UnserializableObject()

    result = OperationResult(
        success=True,
        data={"good": "data", "bad": unserializable},
        metadata={"info": "test"}
    )

    print("📥 输入:")
    print(f"   success: {result.success}")
    print(f"   data: {result.data}")
    print(f"   metadata: {result.metadata}")

    formatted = _format_result(result)

    print("📤 输出:")
    print(json.dumps(formatted, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    print("🐛 _format_result 方法调试分析")
    print("=" * 60)

    debug_basic_result()
    debug_enum_result()
    debug_dataclass_result()
    debug_fallback_mechanism()
