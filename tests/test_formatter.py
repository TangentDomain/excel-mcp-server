#!/usr/bin/env python3
"""
Excel MCP Server - 格式化工具单元测试

全面测试 src/utils/formatter.py 中的格式化功能，确保代码质量和可靠性。

测试覆盖范围：
1. format_operation_result 主函数的各种场景
2. JSON序列化功能测试
3. 紧凑数组格式转换测试
4. null值深度清理测试
5. 错误回退机制测试
6. 边界情况和异常处理测试
"""

import pytest
import json
import sys
from pathlib import Path
from enum import Enum
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# 导入要测试的模块
from src.utils.formatter import (
    format_operation_result,
    _serialize_to_json_dict,
    _convert_to_compact_array_format,
    _deep_clean_nulls,
    _fallback_format_result
)

# ==================== 测试数据模型 ====================

@dataclass
class OperationResult:
    """模拟操作结果数据类"""
    success: bool
    message: Optional[str] = None
    error: Optional[str] = None
    data: Optional[Any] = None
    metadata: Optional[Dict[str, Any]] = None


@dataclass
class MockFieldDifference:
    """模拟字段差异数据类"""
    field_name: str
    old_value: Any
    new_value: Any
    change_type: str


@dataclass  
class MockRowDifference:
    """模拟行差异数据类"""
    row_id: Any
    difference_type: str
    row_index1: int
    row_index2: int
    sheet_name: str
    detailed_field_differences: Optional[List[MockFieldDifference]] = None


@dataclass
class MockStructuredDataComparison:
    """模拟结构化数据比较结果"""
    sheet_name: str
    exists_in_file1: bool
    exists_in_file2: bool
    total_differences: int
    row_differences: List[MockRowDifference]


class MockEnum(Enum):
    """测试枚举"""
    ACTIVE = "active"
    INACTIVE = "inactive"
    PENDING = "pending"


class DifferenceType(Enum):
    """差异类型枚举"""
    ROW_ADDED = "row_added"
    ROW_REMOVED = "row_removed"
    ROW_MODIFIED = "row_modified"


# ==================== 测试类 ====================

class TestFormatOperationResult:
    """format_operation_result 主函数测试"""
    
    def test_basic_success_result(self):
        """测试基本成功结果格式化"""
        result = OperationResult(
            success=True,
            message="操作成功",
            data={"test": "data", "count": 42},
            metadata={"timestamp": "2025-08-22"}
        )
        
        formatted = format_operation_result(result)
        
        assert formatted["success"] is True
        assert formatted["message"] == "操作成功"
        assert formatted["data"]["test"] == "data"
        assert formatted["data"]["count"] == 42
        assert formatted["metadata"]["timestamp"] == "2025-08-22"
    
    def test_basic_failure_result(self):
        """测试基本失败结果格式化"""
        result = OperationResult(
            success=False,
            error="文件不存在",
            message="读取失败"
        )
        
        formatted = format_operation_result(result)
        
        assert formatted["success"] is False
        assert formatted["error"] == "文件不存在"
        assert formatted["message"] == "读取失败"
        assert "data" not in formatted
    
    def test_enum_serialization(self):
        """测试枚举类型序列化"""
        result = OperationResult(
            success=True,
            data={
                "status": MockEnum.ACTIVE,
                "type": DifferenceType.ROW_ADDED,
                "multiple_enums": [MockEnum.PENDING, MockEnum.INACTIVE]
            }
        )
        
        formatted = format_operation_result(result)
        
        assert formatted["data"]["status"] == "active"
        assert formatted["data"]["type"] == "row_added"
        assert formatted["data"]["multiple_enums"] == ["pending", "inactive"]
    
    def test_dataclass_serialization(self):
        """测试数据类序列化"""
        field_diff = MockFieldDifference(
            field_name="技能名称",
            old_value="火球术",
            new_value="冰箭术",
            change_type="text_change"
        )
        
        result = OperationResult(
            success=True,
            data={"field_diff": field_diff}
        )
        
        formatted = format_operation_result(result)
        
        field_data = formatted["data"]["field_diff"]
        assert field_data["field_name"] == "技能名称"
        assert field_data["old_value"] == "火球术"
        assert field_data["new_value"] == "冰箭术"
        assert field_data["change_type"] == "text_change"
    
    def test_null_cleaning(self):
        """测试null值清理功能"""
        result = OperationResult(
            success=True,
            data={
                "valid_data": "保留",
                "null_field": None,
                "empty_dict": {},
                "empty_list": [],
                "nested": {
                    "value": "保留",
                    "null_value": None,
                    "empty_nested_dict": {},
                    "empty_nested_list": []
                },
                "list_with_nulls": [1, None, "test", None, {}, []]
            },
            metadata={"valid_meta": "value", "null_meta": None}
        )
        
        formatted = format_operation_result(result)
        
        # null值和空容器应该被清理
        data = formatted["data"]
        assert "null_field" not in data
        assert "empty_dict" not in data
        assert "empty_list" not in data
        assert "null_value" not in data["nested"]
        assert "empty_nested_dict" not in data["nested"]
        assert "empty_nested_list" not in data["nested"]
        assert "null_meta" not in formatted["metadata"]
        
        # 有效值应该保留
        assert data["valid_data"] == "保留"
        assert data["nested"]["value"] == "保留"
        assert data["list_with_nulls"] == [1, "test"]
        assert formatted["metadata"]["valid_meta"] == "value"
    
    def test_compact_array_format_conversion(self):
        """测试紧凑数组格式转换"""
        # 创建包含行差异的结构化比较数据
        field_diff1 = MockFieldDifference("技能名称", "火球术", "冰箭术", "text_change")
        field_diff2 = MockFieldDifference("伤害", 100, 150, "numeric_change")
        
        row_diff1 = MockRowDifference(
            row_id="1001",
            difference_type="row_modified",
            row_index1=5,
            row_index2=7,
            sheet_name="TrSkill",
            detailed_field_differences=[field_diff1, field_diff2]
        )
        
        row_diff2 = MockRowDifference(
            row_id="1002",
            difference_type="row_added", 
            row_index1=0,
            row_index2=8,
            sheet_name="TrSkill"
            # detailed_field_differences 使用默认值None
        )
        
        structured_data = MockStructuredDataComparison(
            sheet_name="TrSkill比较",
            exists_in_file1=True,
            exists_in_file2=True,
            total_differences=2,
            row_differences=[row_diff1, row_diff2]
        )
        
        result = OperationResult(success=True, data=structured_data)
        formatted = format_operation_result(result)
        
        # 检查是否转换为紧凑数组格式
        row_diffs = formatted["data"]["row_differences"]
        assert isinstance(row_diffs, list)
        assert len(row_diffs) >= 3  # 头部 + 至少2行数据
        
        # 检查头部字段定义
        header = row_diffs[0]
        expected_header = ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"]
        assert header == expected_header
        
        # 检查第一行数据（包含字段差异）
        first_row = row_diffs[1]
        assert first_row[0] == "1001"  # row_id
        assert first_row[1] == "row_modified"  # difference_type
        assert first_row[2] == 5  # row_index1
        assert first_row[3] == 7  # row_index2
        assert first_row[4] == "TrSkill"  # sheet_name
        
        # 检查字段差异数组格式
        field_diffs = first_row[5]
        assert isinstance(field_diffs, list)
        assert len(field_diffs) == 2
        assert field_diffs[0] == ["技能名称", "火球术", "冰箭术", "text_change"]
        assert field_diffs[1] == ["伤害", 100, 150, "numeric_change"]
        
        # 检查第二行数据（无字段差异）
        second_row = row_diffs[2]
        assert second_row[0] == "1002"
        assert second_row[1] == "row_added"
        assert len(second_row) == 6  # 包含所有6个字段
        assert second_row[5] is None  # 没有字段差异
    
    def test_prevent_duplicate_conversion(self):
        """测试防止重复转换已是紧凑格式的数据"""
        # 创建已经是紧凑数组格式的数据
        already_compact_data = {
            "row_differences": [
                ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"],
                ["1001", "row_added", 0, 5, "TrSkill", None],
                ["1002", "row_modified", 3, 6, "TrSkill", [["技能名称", "旧值", "新值", "text_change"]]]
            ],
            "total_differences": 2
        }
        
        result = OperationResult(success=True, data=already_compact_data)
        formatted = format_operation_result(result)
        
        # 数据应该保持不变，不重复转换
        row_diffs = formatted["data"]["row_differences"]
        assert len(row_diffs) == 3  # 头部 + 2行数据
        expected_header = ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"]
        assert row_diffs[0] == expected_header
        # 检查数据行（第二行是没有field_differences的）
        assert row_diffs[1][:5] == ["1001", "row_added", 0, 5, "TrSkill"]
        assert row_diffs[2][5] == [["技能名称", "旧值", "新值", "text_change"]]
    
    def test_json_serialization_fallback(self):
        """测试JSON序列化失败时的回退机制"""
        # 创建包含不可序列化对象的数据
        class UnserializableObject:
            def __str__(self):
                return "UnserializableObject"
            
            def __init__(self):
                # 创建循环引用导致序列化失败
                self.ref = self
        
        unserializable = UnserializableObject()
        
        result = OperationResult(
            success=True,
            data={"good": "data", "bad": unserializable},
            metadata={"info": "test"}
        )
        
        # 应该触发回退机制但仍能格式化
        formatted = format_operation_result(result)
        
        assert formatted["success"] is True
        # 回退机制会将数据放在data字段中
        assert "data" in formatted or "good" in formatted
    
    def test_empty_data_handling(self):
        """测试空数据处理"""
        # 空数据
        result1 = OperationResult(success=True, data=None)
        formatted1 = format_operation_result(result1)
        assert formatted1["success"] is True
        assert "data" not in formatted1
        
        # 空字典数据
        result2 = OperationResult(success=True, data={})
        formatted2 = format_operation_result(result2)
        assert formatted2["success"] is True
        # 空字典会被null清理移除
        assert "data" not in formatted2
        
        # 空列表数据
        result3 = OperationResult(success=True, data=[])
        formatted3 = format_operation_result(result3)
        assert formatted3["success"] is True
        # 空列表会被null清理移除
        assert "data" not in formatted3
    
    def test_complex_nested_structure(self):
        """测试复杂嵌套结构处理"""
        complex_data = {
            "level1": {
                "level2": {
                    "level3": {
                        "value": "深层数据",
                        "null_field": None,
                        "enum": MockEnum.ACTIVE,
                        "list": [1, None, {"nested": "value", "empty": None}]
                    }
                }
            },
            "top_level_enum": DifferenceType.ROW_ADDED
        }
        
        result = OperationResult(success=True, data=complex_data)
        formatted = format_operation_result(result)
        
        # 检查深层嵌套是否正确处理
        level3 = formatted["data"]["level1"]["level2"]["level3"]
        assert level3["value"] == "深层数据"
        assert level3["enum"] == "active"
        assert "null_field" not in level3
        assert level3["list"] == [1, {"nested": "value"}]
        assert formatted["data"]["top_level_enum"] == "row_added"


class TestHelperFunctions:
    """辅助函数单元测试"""
    
    def test_serialize_to_json_dict_success(self):
        """测试JSON序列化成功场景"""
        @dataclass
        class TestData:
            name: str
            status: MockEnum
            count: int
        
        test_obj = TestData("测试", MockEnum.ACTIVE, 42)
        result = OperationResult(success=True, data=test_obj)
        
        serialized = _serialize_to_json_dict(result)
        
        assert serialized["success"] is True
        assert serialized["data"]["name"] == "测试"
        assert serialized["data"]["status"] == "active"
        assert serialized["data"]["count"] == 42
    
    def test_serialize_to_json_dict_failure(self):
        """测试JSON序列化失败场景"""
        class UnserializableClass:
            def __init__(self):
                self.ref = self  # 循环引用
        
        unserializable = UnserializableClass()
        result = OperationResult(success=True, data=unserializable)
        
        # 应该抛出异常
        with pytest.raises(Exception):
            _serialize_to_json_dict(result)
    
    def test_convert_to_compact_array_format_valid_data(self):
        """测试有效数据的紧凑数组格式转换"""
        data = {
            "row_differences": [
                {
                    "row_id": "1001",
                    "difference_type": "row_added",
                    "row_index1": 0,
                    "row_index2": 5,
                    "sheet_name": "TrSkill",
                    "detailed_field_differences": [
                        {
                            "field_name": "技能名称",
                            "old_value": "",
                            "new_value": "火球术",
                            "change_type": "text_change"
                        }
                    ]
                }
            ],
            "total_differences": 1
        }
        
        result = _convert_to_compact_array_format(data)
        
        row_diffs = result["row_differences"]
        assert len(row_diffs) == 2  # 头部 + 1行数据
        
        # 检查头部
        expected_header = ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"]
        assert row_diffs[0] == expected_header
        
        # 检查数据行
        data_row = row_diffs[1]
        assert data_row[0] == "1001"
        assert data_row[1] == "row_added"
        assert data_row[5][0] == ["技能名称", "", "火球术", "text_change"]
    
    def test_convert_to_compact_array_format_invalid_data(self):
        """测试无效数据不进行转换"""
        # 没有row_differences的数据
        data1 = {"total_differences": 0}
        result1 = _convert_to_compact_array_format(data1)
        assert result1 == data1
        
        # 非字典数据
        data2 = ["not", "a", "dict"]
        result2 = _convert_to_compact_array_format(data2)
        assert result2 == data2
        
        # 空的row_differences
        data3 = {"row_differences": []}
        result3 = _convert_to_compact_array_format(data3)
        assert result3 == data3
    
    def test_convert_to_compact_array_format_already_compact(self):
        """测试已经是紧凑格式的数据不重复转换"""
        data = {
            "row_differences": [
                ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"],
                ["1001", "row_added", 0, 5, "TrSkill", None]
            ]
        }
        
        result = _convert_to_compact_array_format(data)
        
        # 应该保持不变
        assert result == data
    
    def test_deep_clean_nulls_dict(self):
        """测试字典的null值清理"""
        data = {
            "valid": "value",
            "null_field": None,
            "empty_dict": {},
            "empty_list": [],
            "nested": {
                "keep": "this",
                "remove": None
            }
        }
        
        result = _deep_clean_nulls(data)
        
        assert "null_field" not in result
        assert "empty_dict" not in result
        assert "empty_list" not in result
        assert result["valid"] == "value"
        assert result["nested"]["keep"] == "this"
        assert "remove" not in result["nested"]
    
    def test_deep_clean_nulls_list(self):
        """测试列表的null值清理"""
        data = [1, None, "test", None, {}, [], {"valid": "data", "null": None}]
        
        result = _deep_clean_nulls(data)
        
        assert result == [1, "test", {"valid": "data"}]
    
    def test_deep_clean_nulls_primitives(self):
        """测试原始类型不受影响"""
        assert _deep_clean_nulls("string") == "string"
        assert _deep_clean_nulls(42) == 42
        assert _deep_clean_nulls(True) is True
        assert _deep_clean_nulls(None) is None
    
    def test_fallback_format_result_success(self):
        """测试回退格式化成功场景"""
        result = OperationResult(
            success=True,
            data={"test": "value"},
            metadata={"info": "meta"},
            message="success"
        )
        
        formatted = _fallback_format_result(result, Exception("test"))
        
        assert formatted["success"] is True
        assert formatted["data"]["test"] == "value"
        assert formatted["metadata"]["info"] == "meta"
        assert formatted["message"] == "success"
    
    def test_fallback_format_result_failure(self):
        """测试回退格式化失败场景"""
        result = OperationResult(
            success=False,
            error="test error",
            message="failure"
        )
        
        formatted = _fallback_format_result(result, Exception("test"))
        
        assert formatted["success"] is False
        assert formatted["error"] == "test error"
        # 失败情况下，message不会被包含（根据实际实现）
        assert "data" not in formatted
    
    def test_fallback_format_result_with_unserializable_data(self):
        """测试回退格式化包含不可序列化数据的场景"""
        class UnserializableData:
            def __init__(self):
                self.ref = self
        
        result = OperationResult(
            success=True,
            data=UnserializableData()
        )
        
        formatted = _fallback_format_result(result, Exception("test"))
        
        assert formatted["success"] is True
        # 不可序列化的对象在回退方案中会转换为其__dict__属性
        assert "data" in formatted
        assert isinstance(formatted["data"], dict)


# ==================== 性能和边界测试 ====================

class TestPerformanceAndEdgeCases:
    """性能和边界情况测试"""
    
    def test_large_data_handling(self):
        """测试大数据量处理"""
        # 创建包含大量行差异的数据
        large_row_differences = []
        for i in range(1000):
            large_row_differences.append(MockRowDifference(
                row_id=f"id_{i}",
                difference_type="row_modified",
                row_index1=i,
                row_index2=i + 1000,
                sheet_name="LargeSheet",
                detailed_field_differences=[
                    MockFieldDifference(f"field_{j}", f"old_{i}_{j}", f"new_{i}_{j}", "text_change")
                    for j in range(5)
                ]
            ))
        
        structured_data = MockStructuredDataComparison(
            sheet_name="大数据测试",
            exists_in_file1=True,
            exists_in_file2=True,
            total_differences=1000,
            row_differences=large_row_differences
        )
        
        result = OperationResult(success=True, data=structured_data)
        
        # 应该能够处理大数据量而不出错
        formatted = format_operation_result(result)
        
        assert formatted["success"] is True
        assert formatted["data"]["total_differences"] == 1000
        row_diffs = formatted["data"]["row_differences"]
        assert len(row_diffs) == 1001  # 头部 + 1000行数据
    
    def test_deep_nested_structure(self):
        """测试深度嵌套结构"""
        # 创建50层深的嵌套结构
        deep_data = {"value": "leaf"}
        for i in range(50):
            deep_data = {"level": i, "nested": deep_data, "null": None}
        
        result = OperationResult(success=True, data=deep_data)
        formatted = format_operation_result(result)
        
        assert formatted["success"] is True
        # 应该能够处理深度嵌套而不栈溢出
        current = formatted["data"]
        for i in range(50):
            assert current["level"] == 49 - i
            assert "null" not in current  # null值应该被清理
            current = current["nested"]
        assert current["value"] == "leaf"


# ==================== 运行测试 ====================

if __name__ == "__main__":
    # 运行所有测试
    pytest.main([__file__, "-v", "--tb=short"])
