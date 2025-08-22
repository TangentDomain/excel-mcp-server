#!/usr/bin/env python3

import json
from enum import Enum
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

@dataclass  
class MockRowDifference:
    row_id: Any
    difference_type: str
    row_index1: int
    row_index2: int
    sheet_name: str
    detailed_field_differences: Optional[List] = None

# 测试JSON序列化的行为
row_diff = MockRowDifference(
    row_id="1002",
    difference_type="row_added", 
    row_index1=0,
    row_index2=8,
    sheet_name="TrSkill"
    # detailed_field_differences 使用默认值None
)

# 模拟JSON序列化过程
def _json_serializer(obj):
    if hasattr(obj, '__dict__'):
        return obj.__dict__
    else:
        return str(obj)

# 转换为JSON字符串再解析回字典
json_str = json.dumps(row_diff, default=_json_serializer, ensure_ascii=False)
print("JSON字符串:", json_str)

parsed = json.loads(json_str)
print("解析后的字典:", parsed)
print("detailed_field_differences键是否存在:", 'detailed_field_differences' in parsed)
print("detailed_field_differences的值:", parsed.get('detailed_field_differences'))

# 测试构建紧凑数组的行为
field_diffs = parsed.get('detailed_field_differences', [])
print("获取的field_diffs:", field_diffs)

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

compact_field_diffs = convert_field_differences_to_array(field_diffs)
print("转换后的compact_field_diffs:", compact_field_diffs)

# 构建紧凑数组行
compact_row = [
    parsed.get('row_id', ''),
    parsed.get('difference_type', ''),
    parsed.get('row_index1', 0),
    parsed.get('row_index2', 0),
    parsed.get('sheet_name', ''),
    compact_field_diffs
]

print("最终的compact_row:", compact_row)
print("compact_row长度:", len(compact_row))
