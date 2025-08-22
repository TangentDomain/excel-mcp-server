#!/usr/bin/env python3

import json
from enum import Enum
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

@dataclass
class MockFieldDifference:
    field_name: str
    old_value: Any
    new_value: Any
    change_type: str

@dataclass
class MockRowDifference:
    row_id: Any
    difference_type: str
    row_index1: int
    row_index2: int
    sheet_name: str
    detailed_field_differences: Optional[List[MockFieldDifference]] = None

@dataclass
class MockStructuredDataComparison:
    sheet_name: str
    exists_in_file1: bool
    exists_in_file2: bool
    total_differences: int
    row_differences: List[MockRowDifference]

@dataclass
class OperationResult:
    success: bool
    message: Optional[str] = None
    error: Optional[str] = None
    data: Optional[Any] = None
    metadata: Optional[Dict[str, Any]] = None

# 创建测试数据，完全模拟测试用例
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

# 模拟JSON序列化过程
def _json_serializer(obj):
    if isinstance(obj, Enum):
        return obj.value
    elif hasattr(obj, '__dict__'):
        return obj.__dict__
    else:
        return str(obj)

# 序列化和反序列化
json_str = json.dumps(result, default=_json_serializer, ensure_ascii=False)
serialized_dict = json.loads(json_str)

print("第二个row_difference:")
row_diff2_dict = serialized_dict['data']['row_differences'][1]
print(json.dumps(row_diff2_dict, indent=2, ensure_ascii=False))

print("\n键值对：")
for key, value in row_diff2_dict.items():
    print(f"  {key}: {value} (type: {type(value).__name__})")

# 模拟_build_compact_array的逻辑
def convert_field_differences_to_array(field_diffs):
    if not field_diffs:
        return None

    compact_field_diffs = []
    for field_diff in field_diffs:
        if isinstance(field_diff, dict):
            compact_field_diffs.append([
                field_diff.get('field_name', ''),
                field_diff.get('old_value', ''),
                field_diff.get('new_value', ''),
                field_diff.get('change_type', '')
            ])

    return compact_field_diffs

field_diffs = row_diff2_dict.get('detailed_field_differences', [])
print(f"\nfield_diffs: {field_diffs}")
print(f"not field_diffs: {not field_diffs}")

compact_field_diffs = convert_field_differences_to_array(field_diffs)
print(f"compact_field_diffs: {compact_field_diffs}")

# 构建紧凑行
compact_row = [
    row_diff2_dict.get('row_id', ''),
    row_diff2_dict.get('difference_type', ''),
    row_diff2_dict.get('row_index1', 0),
    row_diff2_dict.get('row_index2', 0),
    row_diff2_dict.get('sheet_name', ''),
    compact_field_diffs
]

print(f"\ncompact_row: {compact_row}")
print(f"compact_row length: {len(compact_row)}")
print(f"Last element (index 5): {compact_row[5]}")
