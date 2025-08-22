#!/usr/bin/env python3
"""
测试用户简化的JSON序列化方案
"""

import json
import sys
import os
from enum import Enum

# 添加src路径以便导入
sys.path.insert(0, os.path.join(os.getcwd(), 'src'))

def test_json_approach():
    """测试JSON序列化方案"""
    print("=== 测试JSON序列化简化方案 ===")

    try:
        # 导入必要的类
        from models.types import OperationResult, CellDifference, DifferenceType, SheetComparison

        # 创建测试数据 - 包含大量None值
        cell_diff1 = CellDifference(
            coordinate="A1",
            difference_type=DifferenceType.VALUE_CHANGED,
            old_value="旧值",
            new_value="新值"
            # old_format, new_format, sheet_name 默认为None
        )

        cell_diff2 = CellDifference(
            coordinate="SHEET",
            difference_type=DifferenceType.SHEET_ADDED,
            sheet_name="新工作表"
            # 其他字段默认为None
        )

        sheet_comp = SheetComparison(
            sheet_name="测试工作表",
            exists_in_file1=True,
            exists_in_file2=True,
            differences=[cell_diff1, cell_diff2],
            total_differences=2,
            structural_changes={}
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

        print("=== 原始数据分析 ===")
        print(f"CellDifference包含的None字段数量: {sum(1 for k, v in cell_diff1.__dict__.items() if v is None)}")
        print(f"SheetComparison包含的None字段数量: {sum(1 for k, v in sheet_comp.__dict__.items() if v is None)}")

        # 使用新的JSON方案
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

        # 步骤1: 转成JSON字符串
        def json_serializer(obj):
            """自定义JSON序列化器，专门处理dataclass和枚举"""
            if isinstance(obj, Enum):
                return obj.value
            elif hasattr(obj, '__dict__'):
                return obj.__dict__
            else:
                return str(obj)

        json_str = json.dumps(result, default=json_serializer, ensure_ascii=False)
        print(f"\n=== JSON字符串分析 ===")
        print(f"JSON字符串长度: {len(json_str)}")
        print(f"JSON中null的数量: {json_str.count('null')}")

        # 步骤2: 转回字典
        result_dict = json.loads(json_str)

        # 步骤3: 清理null值
        cleaned_dict = _deep_clean_nulls(result_dict)

        # 最终结果分析
        final_json = json.dumps(cleaned_dict, ensure_ascii=False, indent=2)
        print(f"\n=== 最终结果分析 ===")
        print(f"清理后JSON长度: {len(final_json)}")

        # 更精确地检测JSON null值（不包括字符串中的null）
        import re
        json_null_pattern = r':\s*null\b|^\s*null\b|\[\s*null\b'  # 匹配真正的JSON null值
        json_null_matches = re.findall(json_null_pattern, final_json)
        real_null_count = len(json_null_matches)

        print(f"清理后真正的JSON null数量: {real_null_count}")

        if real_null_count == 0:
            print("✅ 成功！用户的JSON序列化方案完美解决了null值问题")
        else:
            print("❌ 仍有真正的null值残留")
            for match in json_null_matches:
                print(f"发现null模式: {repr(match)}")

        print(f"\n=== 压缩效果 ===")
        print(f"原始长度: {len(json_str)}")
        print(f"清理后长度: {len(final_json)}")
        print(f"压缩率: {((len(json_str) - len(final_json)) / len(json_str) * 100):.1f}%")

        print(f"\n=== 清理后的完整结果 ===")
        print(final_json)

    except Exception as e:
        print(f"测试失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_json_approach()
