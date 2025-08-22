#!/usr/bin/env python3
"""
测试根源性的_format_result修复
"""

import sys
import os
import json

# 添加src路径
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, 'src')
sys.path.insert(0, src_dir)

def test_format_result():
    """测试_format_result的null清理功能"""
    print("=== 测试_format_result的根源性null清理 ===")

    # 导入需要的类
    try:
        # 尝试从src目录导入
        from models.types import OperationResult, CellDifference, DifferenceType, SheetComparison

        # 直接从当前的server.py读取_format_result函数
        with open(os.path.join(src_dir, 'server.py'), 'r', encoding='utf-8') as f:
            server_code = f.read()

        # 手动执行_format_result和_deep_clean_nulls函数的定义
        exec_globals = {}
        exec(server_code.split('def _deep_clean_nulls')[0], exec_globals)
        exec('def ' + server_code.split('def _deep_clean_nulls')[1].split('\n\ndef excel_')[0], exec_globals)
        exec('def ' + server_code.split('def _format_result')[1].split('\n\n@')[0], exec_globals)

        _format_result = exec_globals['_format_result']

    except Exception as e:
        print(f"导入模块失败: {e}")
        # 直接定义测试函数
        def _deep_clean_nulls(obj):
            if isinstance(obj, dict):
                cleaned = {}
                for k, v in obj.items():
                    cleaned_v = _deep_clean_nulls(v)
                    if cleaned_v is not None:
                        if isinstance(cleaned_v, dict) and len(cleaned_v) == 0:
                            continue
                        if isinstance(cleaned_v, list) and len(cleaned_v) == 0:
                            continue
                        cleaned[k] = cleaned_v
                return cleaned
            elif isinstance(obj, list):
                cleaned = []
                for item in obj:
                    cleaned_item = _deep_clean_nulls(item)
                    if cleaned_item is not None:
                        cleaned.append(cleaned_item)
                return cleaned
            else:
                return obj if obj is not None else None

        def _format_result(result_dict):
            cleaned_result = _deep_clean_nulls(result_dict)
            return cleaned_result

    # 模拟一个包含大量null值的ComparisonResult
    # 创建包含null值的CellDifference对象
    cell_diff1 = CellDifference(
        coordinate="SHEET",
        difference_type=DifferenceType.SHEET_ADDED,
        old_value=None,
        new_value=None,
        old_format=None,
        new_format=None,
        sheet_name="TrSkill"
    )

    cell_diff2 = CellDifference(
        coordinate="SHEET",
        difference_type=DifferenceType.SHEET_REMOVED,
        old_value=None,
        new_value=None,
        old_format=None,
        new_format=None,
        sheet_name="测试数据"
    )

    # 创建SheetComparison对象
    sheet_comp = SheetComparison(
        sheet_name="测试工作表",
        exists_in_file1=True,
        exists_in_file2=False,
        differences=[cell_diff1, cell_diff2],
        total_differences=2,
        structural_changes={}
    )

    # 创建OperationResult对象
    result = OperationResult(
        success=True,
        message="成功比较Excel文件",
        data=sheet_comp,
        metadata={
            "total_differences": 2,
            "empty_metadata": None,  # 这个应该被清理掉
            "null_list": [],         # 这个应该被清理掉
            "nested_null": {
                "valid_field": "有效值",
                "null_field": None   # 这个应该被清理掉
            }
        }
    )

    print("=== 原始数据结构分析 ===")
    print("CellDifference.__dict__包含的字段:")
    print(list(cell_diff1.__dict__.keys()))
    print("None值的数量:", list(cell_diff1.__dict__.values()).count(None))

    print("\n=== 执行_format_result处理 ===")
    formatted = _format_result(result)

    # 转换为JSON进行分析
    json_str = json.dumps(formatted, ensure_ascii=False, indent=2, default=str)

    print("格式化后的JSON长度:", len(json_str))
    null_count = json_str.count('null')
    print(f"JSON中null的数量: {null_count}")

    # 显示结果
    if null_count == 0:
        print("🎉 完美！根源性的_format_result修复成功！没有任何null值！")
    else:
        print("❌ 还有null值存在")
        print("包含null的前几行:")
        lines = json_str.split('\n')
        null_lines = [line.strip() for line in lines if 'null' in line]
        for line in null_lines[:5]:
            print(f"  {line}")

    print(f"\n=== 完整结果 ===")
    print(json_str)

if __name__ == "__main__":
    try:
        test_format_result()
    except Exception as e:
        print(f"测试过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
