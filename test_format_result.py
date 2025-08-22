#!/usr/bin/env python3
"""
_format_result 方法测试脚本

测试 server.py 中的 _format_result 函数的各种场景，
包括成功/失败场景、数据类型处理、紧凑数组格式转换、null值清理等。
"""

import sys
import json
from enum import Enum
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Union

# 添加 src 目录到路径
sys.path.append('src')

# 导入要测试的函数
from server import _format_result

# 导入数据模型
from models.types import OperationResult, DifferenceType, MatchType


# ==================== 测试数据类和枚举 ====================

@dataclass
class MockFieldDifference:
    """模拟字段差异对象"""
    field_name: str
    old_value: Any
    new_value: Any
    change_type: str


@dataclass
class MockRowDifference:
    """模拟行差异对象"""
    row_id: Any
    difference_type: DifferenceType
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


class MockStatus(Enum):
    """测试用枚举"""
    ACTIVE = "active"
    INACTIVE = "inactive"


# ==================== 测试函数 ====================

class TestFormatResult:
    """_format_result 方法测试类"""

    def __init__(self):
        self.test_count = 0
        self.passed_count = 0
        self.failed_count = 0

    def run_test(self, test_name: str, test_func):
        """执行单个测试"""
        self.test_count += 1
        print(f"\n🧪 测试 {self.test_count}: {test_name}")
        print("-" * 60)

        try:
            test_func()
            self.passed_count += 1
            print("✅ 测试通过")
        except Exception as e:
            self.failed_count += 1
            print(f"❌ 测试失败: {str(e)}")
            import traceback
            print(traceback.format_exc())

    def test_basic_success_result(self):
        """测试基本成功结果"""
        result = OperationResult(
            success=True,
            message="操作成功",
            data={"test": "data", "count": 42},
            metadata={"timestamp": "2025-08-22"}
        )

        formatted = _format_result(result)

        assert formatted["success"] is True
        assert formatted["message"] == "操作成功"
        assert formatted["test"] == "data"
        assert formatted["count"] == 42
        assert formatted["timestamp"] == "2025-08-22"
        print("📋 基本成功结果格式化正确")

    def test_basic_failure_result(self):
        """测试基本失败结果"""
        result = OperationResult(
            success=False,
            error="文件不存在",
            message="读取失败"
        )

        formatted = _format_result(result)

        assert formatted["success"] is False
        assert formatted["error"] == "文件不存在"
        print("📋 基本失败结果格式化正确")

    def test_enum_serialization(self):
        """测试枚举类型序列化"""
        result = OperationResult(
            success=True,
            data={
                "status": MockStatus.ACTIVE,
                "match_type": MatchType.VALUE,
                "diff_type": DifferenceType.ROW_ADDED
            }
        )

        formatted = _format_result(result)

        assert formatted["status"] == "active"
        assert formatted["match_type"] == "value"
        assert formatted["diff_type"] == "row_added"
        print("📋 枚举类型序列化正确")

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

        formatted = _format_result(result)

        field_data = formatted["field_diff"]
        assert field_data["field_name"] == "技能名称"
        assert field_data["old_value"] == "火球术"
        assert field_data["new_value"] == "冰箭术"
        assert field_data["change_type"] == "text_change"
        print("📋 数据类序列化正确")

    def test_null_cleaning(self):
        """测试null值清理"""
        result = OperationResult(
            success=True,
            data={
                "name": "测试",
                "empty_field": None,
                "nested": {
                    "value": "保留",
                    "null_value": None,
                    "empty_list": [],
                    "empty_dict": {}
                },
                "list_with_nulls": [1, None, "test", None, {}]
            },
            metadata={"key": None, "valid": "value"}
        )

        formatted = _format_result(result)

        # null 值应该被清理掉
        assert "empty_field" not in formatted
        assert "null_value" not in formatted.get("nested", {})
        assert "empty_list" not in formatted.get("nested", {})
        assert "empty_dict" not in formatted.get("nested", {})
        assert "key" not in formatted

        # 有效值应该保留
        assert formatted["name"] == "测试"
        assert formatted["nested"]["value"] == "保留"
        assert formatted["valid"] == "value"
        assert formatted["list_with_nulls"] == [1, "test"]
        print("📋 null值清理正确")

    def test_compact_array_format_conversion(self):
        """测试紧凑数组格式转换"""
        # 创建包含行差异的结构化比较数据
        field_diff1 = MockFieldDifference("技能名称", "火球术", "冰箭术", "text_change")
        field_diff2 = MockFieldDifference("伤害", 100, 150, "numeric_change")

        row_diff1 = MockRowDifference(
            row_id="1001",
            difference_type=DifferenceType.ROW_MODIFIED,
            row_index1=5,
            row_index2=7,
            sheet_name="TrSkill",
            detailed_field_differences=[field_diff1, field_diff2]
        )

        row_diff2 = MockRowDifference(
            row_id="1002",
            difference_type=DifferenceType.ROW_ADDED,
            row_index1=0,
            row_index2=8,
            sheet_name="TrSkill"
        )

        structured_data = MockStructuredDataComparison(
            sheet_name="TrSkill比较",
            exists_in_file1=True,
            exists_in_file2=True,
            total_differences=2,
            row_differences=[row_diff1, row_diff2]
        )

        result = OperationResult(
            success=True,
            data=structured_data
        )

        formatted = _format_result(result)

        # 检查是否转换为紧凑数组格式
        row_diffs = formatted["data"]["row_differences"]
        assert isinstance(row_diffs, list)
        assert len(row_diffs) >= 3  # 头部 + 至少2行数据

        # 检查头部字段定义
        header = row_diffs[0]
        assert header == ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"]

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
        assert second_row[5] is None  # 没有字段差异

        print("📋 紧凑数组格式转换正确")

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

        result = OperationResult(
            success=True,
            data=already_compact_data
        )

        formatted = _format_result(result)

        # 数据应该保持不变，不重复转换
        row_diffs = formatted["data"]["row_differences"]
        assert len(row_diffs) == 3  # 头部 + 2行数据
        assert row_diffs[0] == ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"]
        assert row_diffs[1] == ["1001", "row_added", 0, 5, "TrSkill", None]
        assert row_diffs[2][5] == [["技能名称", "旧值", "新值", "text_change"]]

        print("📋 防止重复转换功能正确")

    def test_json_serialization_fallback(self):
        """测试JSON序列化失败时的回退机制"""
        # 创建包含不可序列化对象的数据
        class UnserializableObject:
            def __str__(self):
                return "UnserializableObject"

        unserializable = UnserializableObject()

        result = OperationResult(
            success=True,
            data={"serializable": "data", "unserializable": unserializable},
            metadata={"info": "test"}
        )

        # 这应该触发回退机制
        formatted = _format_result(result)

        # 应该仍然能够格式化，使用回退方案
        assert formatted["success"] is True
        assert "serializable" in formatted or "data" in formatted
        print("📋 JSON序列化回退机制工作正常")

    def test_empty_data_handling(self):
        """测试空数据处理"""
        # 空数据
        result1 = OperationResult(success=True, data=None)
        formatted1 = _format_result(result1)
        assert formatted1["success"] is True
        assert "data" not in formatted1 or formatted1.get("data") is None

        # 空字典数据
        result2 = OperationResult(success=True, data={})
        formatted2 = _format_result(result2)
        assert formatted2["success"] is True

        # 空列表数据
        result3 = OperationResult(success=True, data=[])
        formatted3 = _format_result(result3)
        assert formatted3["success"] is True

        print("📋 空数据处理正确")

    def test_complex_nested_structure(self):
        """测试复杂嵌套结构"""
        complex_data = {
            "level1": {
                "level2": {
                    "level3": {
                        "value": "深层数据",
                        "null_field": None,
                        "enum": MockStatus.ACTIVE,
                        "list": [1, None, {"nested": "value", "empty": None}]
                    }
                }
            },
            "top_level_enum": DifferenceType.ROW_ADDED
        }

        result = OperationResult(success=True, data=complex_data)
        formatted = _format_result(result)

        # 检查深层嵌套是否正确处理
        level3 = formatted["level1"]["level2"]["level3"]
        assert level3["value"] == "深层数据"
        assert level3["enum"] == "active"
        assert "null_field" not in level3
        assert level3["list"] == [1, {"nested": "value"}]
        assert formatted["top_level_enum"] == "row_added"

        print("📋 复杂嵌套结构处理正确")

    def run_all_tests(self):
        """运行所有测试"""
        print("🚀 开始测试 _format_result 方法")
        print("=" * 80)

        test_methods = [
            ("基本成功结果", self.test_basic_success_result),
            ("基本失败结果", self.test_basic_failure_result),
            ("枚举类型序列化", self.test_enum_serialization),
            ("数据类序列化", self.test_dataclass_serialization),
            ("null值清理", self.test_null_cleaning),
            ("紧凑数组格式转换", self.test_compact_array_format_conversion),
            ("防止重复转换", self.test_prevent_duplicate_conversion),
            ("JSON序列化回退", self.test_json_serialization_fallback),
            ("空数据处理", self.test_empty_data_handling),
            ("复杂嵌套结构", self.test_complex_nested_structure)
        ]

        for test_name, test_method in test_methods:
            self.run_test(test_name, test_method)

        # 输出测试总结
        print("\n" + "=" * 80)
        print("🏁 测试完成")
        print(f"📊 总计: {self.test_count} 个测试")
        print(f"✅ 通过: {self.passed_count} 个")
        print(f"❌ 失败: {self.failed_count} 个")

        if self.failed_count == 0:
            print("🎉 所有测试通过！_format_result 方法工作正常")
        else:
            print(f"⚠️  有 {self.failed_count} 个测试失败，需要检查代码")

        return self.failed_count == 0


# ==================== 主程序 ====================

if __name__ == "__main__":
    tester = TestFormatResult()
    success = tester.run_all_tests()

    if success:
        print("\n🔧 _format_result 方法测试完成，功能正常")
    else:
        print("\n🚨 _format_result 方法存在问题，请检查代码")
        sys.exit(1)
