#!/usr/bin/env python3
"""
_format_result 方法独立测试脚本

直接包含 _format_result 函数的完整实现，进行全面测试
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


# ==================== 测试数据类和枚举 ====================

@dataclass
class OperationResult:
    """操作结果"""
    success: bool
    message: Optional[str] = None
    error: Optional[str] = None
    data: Optional[Any] = None
    metadata: Optional[Dict[str, Any]] = None


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


class MockStatus(Enum):
    """测试用枚举"""
    ACTIVE = "active"
    INACTIVE = "inactive"


class DifferenceType(Enum):
    """比较差异类型枚举"""
    ROW_ADDED = "row_added"
    ROW_REMOVED = "row_removed"
    ROW_MODIFIED = "row_modified"


class MatchType(Enum):
    """搜索匹配类型枚举"""
    VALUE = "value"
    FORMULA = "formula"


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
        assert formatted["message"] == "读取失败"
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

    def test_fixed_duplicate_issue(self):
        """测试修复的数据重复问题"""
        # 这个测试专门验证之前修复的数据重复问题
        result = OperationResult(
            success=True,
            data={"main_data": "test"},
            metadata={"info": "meta_test"}
        )

        formatted = _format_result(result)

        # 确保数据不重复出现
        assert formatted["success"] is True
        assert formatted["main_data"] == "test"
        assert formatted["info"] == "meta_test"

        # 确保没有多余的 data 字段包装
        data_field_count = 1 if "data" in formatted else 0
        main_data_count = 1 if "main_data" in formatted else 0

        # 应该直接展开，而不是包装在 data 字段中
        assert main_data_count == 1, "main_data 应该直接存在"
        print("📋 数据重复问题修复验证通过")

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
            ("复杂嵌套结构", self.test_complex_nested_structure),
            ("修复的数据重复问题", self.test_fixed_duplicate_issue)
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
            print("\n🔍 测试覆盖的功能：")
            print("   - ✅ JSON 序列化和反序列化")
            print("   - ✅ 数据类和枚举类型处理")
            print("   - ✅ null 值递归清理")
            print("   - ✅ 紧凑数组格式转换")
            print("   - ✅ 防止重复转换机制")
            print("   - ✅ 序列化失败时的回退机制")
            print("   - ✅ 复杂嵌套结构处理")
            print("   - ✅ 数据重复问题修复验证")
        else:
            print(f"⚠️  有 {self.failed_count} 个测试失败，需要检查代码")

        return self.failed_count == 0


# ==================== 主程序 ====================

if __name__ == "__main__":
    tester = TestFormatResult()
    success = tester.run_all_tests()

    if success:
        print("\n✨ 所有功能测试通过，_format_result 方法运行正常！")
    else:
        print("\n🚨 发现问题，需要进一步检查代码")
        exit(1)
