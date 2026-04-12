#!/usr/bin/env python3
"""批量精简server.py中的工具docstring，减少token消耗。
策略：只保留核心功能描述+关键参数，去掉冗长的场景/技巧/最佳实践。"""

import re

# 手动定义每个工具的精简docstring
SLIM_DOCS = {
    "excel_list_sheets": '列出Excel文件中的所有工作表名称。查询前先用此工具确认工作表存在。',

    "excel_search": '在Excel中搜索匹配pattern的单元格。支持正则、大小写、全词匹配。',
    
    "excel_search_directory": '在目录下所有Excel文件中搜索内容。支持文件类型过滤和递归搜索。',

    "excel_get_range": '读取指定范围的数据。返回{headers, data, shape}。支持include_formatting获取样式。',

    "excel_get_headers": '提取工作表表头信息。支持双行表头（中文描述+英文字段名）。不传sheet_name获取所有表的表头。',

    "excel_update_range": '写入数据到指定范围。preserve_formulas=True时保留已有公式不被覆盖。',

    "excel_assess_data_impact": '评估修改操作的影响范围。返回受影响行数、关键值变化等，修改前必用。',

    "excel_get_operation_history": '查看最近的Excel操作记录。可按文件过滤。',

    "excel_create_backup": '为Excel文件创建备份。备份存放在同级backup目录。',

    "excel_restore_backup": '从备份文件恢复Excel。target_path不传则覆盖原文件。',

    "excel_list_backups": '列出文件的所有备份版本及时间。',

    "excel_insert_rows": '在指定位置插入空行。row_index从0开始。',

    "excel_insert_columns": '在指定位置插入空列。column_index从1开始。',

    "excel_find_last_row": '查找工作表最后一行。可指定列来找该列最后一个有值的行。追加数据前必用。',

    "excel_create_file": '创建新Excel文件。可指定初始工作表名称列表。',

    "excel_export_to_csv": '将工作表导出为CSV。支持指定编码和分隔符。',

    "excel_import_from_csv": '从CSV创建Excel文件。支持编码和分隔符配置。',

    "excel_convert_format": 'Excel/CSV/JSON格式互转。',

    "excel_merge_files": '合并多个Excel文件。merge_mode: sheets(每个文件一个表) | append(纵向追加) | columns(横向拼接)。',

    "excel_get_file_info": '获取文件元数据：大小、工作表数、行列范围等。',

    "excel_create_sheet": '创建新工作表。可指定插入位置index。',

    "excel_delete_sheet": '删除指定工作表。',

    "excel_rename_sheet": '重命名工作表。',

    "excel_copy_sheet": '复制工作表（含数据和格式）。可指定目标文件。',

    "excel_rename_column": '修改表头（列名）。只改header_row指定的行。',

    "excel_upsert_row": '按key_column+key_value查找行，存在则更新，不存在则插入。update_columns指定要更新的列。',

    "excel_batch_insert_rows": '批量插入多行数据。data为字典列表，header_row指定表头行号。',

    "excel_delete_rows": '删除行。支持按索引(row_index)或条件(where_column+where_value)删除。',

    "excel_delete_columns": '删除指定位置开始的列。column_index从1开始。',

    "excel_set_formula": '在单元格写入Excel公式。',

    "excel_evaluate_formula": '临时计算公式结果，不修改文件。',

    "excel_query": 'SQL查询引擎。支持WHERE/JOIN/GROUP BY/ORDER BY/LIMIT/子查询。query_expression示例:\n'
                   '  "SELECT * FROM 技能表 WHERE 伤害 > 100"\n'
                   '  "SELECT a.名称, b.效果 FROM 技能表 a JOIN 装备表 b ON a.ID = b.技能ID"\n'
                   '  "SELECT 类型, COUNT(*) as 数量 FROM 技能表 GROUP BY 类型"',

    "excel_update_query": 'SQL批量修改。dry_run=True预览变更不实际写入。示例:\n'
                          '  "UPDATE 技能表 SET 伤害 = 200 WHERE 等级 >= 5"',

    "excel_describe_table": '获取表结构信息：列名、类型、样本数据、非空统计。',

    "excel_format_cells": '设置单元格样式。formatting字段: bold/italic/underline/font_size/font_color/bg_color/'
                          'number_format/alignment/wrap_text/border_style。只传需要修改的字段。',

    "excel_merge_cells": '合并指定范围为一个大单元格。',

    "excel_unmerge_cells": '取消合并，恢复为独立单元格。',

    "excel_set_borders": '为范围设置边框。border_style: thin/thick/double/dotted/dashed。',

    "excel_set_row_height": '设置行高（磅值）。',

    "excel_set_column_width": '设置列宽（字符单位）。',

    "excel_compare_files": '逐单元格比较两个Excel文件差异。',

    "excel_check_duplicate_ids": '扫描ID列，返回重复值及所在行号。',

    "excel_compare_sheets": '比较两个工作表的差异：新增/删除/修改的行和列。',

    "excel_server_stats": '服务器状态：缓存、调用次数、运行时间。',

    "excel_batch_update_ranges": '批量更新多个范围。updates为[{range, data}]列表。',

    "excel_merge_multiple_files": '合并多个文件。merge_mode: append(纵向追加) | sheets(分表合并)。',

    "excel_create_chart": '在工作表中创建图表。chart_type: line/bar/pie/scatter/area等。',

    "excel_list_charts": '列出工作表中的所有图表信息。',

    "excel_set_data_validation": '设置数据验证规则。validation_type: list/whole_number/decimal/date/text_length/custom。',

    "excel_clear_validation": '清除数据验证规则。',

    "excel_add_conditional_format": '添加条件格式规则。支持高亮/数据条/色阶/图标集。',

    "excel_clear_conditional_format": '清除条件格式。',

    "excel_write_only_override": '大文件高性能覆盖写入。range_spec: "sheet!A1:D10"。不读取已有内容，直接覆盖。适合批量导入场景。',
}


def slim_docstring(name: str, original: str) -> str:
    """用精简版替换原始docstring"""
    if name in SLIM_DOCS:
        return SLIM_DOCS[name]
    # 未手动定义的，提取第一行作为摘要
    for line in original.strip().split('\n'):
        line = line.strip().strip('*').strip()
        if line:
            return line
    return original.strip().split('\n')[0].strip()


def main():
    with open('src/excel_mcp_server_fastmcp/server.py', 'r', encoding='utf-8') as f:
        content = f.read()

    # Match: @mcp.tool() ... def excel_xxx(...): \n    """..."""
    pattern = r'(@mcp\.tool\(\)\s*\n@_track_call\s*\ndef (excel_\w+)\([^)]*\)[^:]*:\s*\n)(\s*"""[\s\S]*?""")'

    count = 0
    total_before = 0
    total_after = 0

    def replacer(m):
        nonlocal count, total_before, total_after
        prefix = m.group(1)
        name = m.group(2)
        original_doc = m.group(3)
        total_before += len(original_doc)
        
        new_doc = slim_docstring(name, original_doc)
        new_doc_block = f'    """{new_doc}"""'
        total_after += len(new_doc_block)
        count += 1
        return prefix + new_doc_block

    new_content = re.sub(pattern, replacer, content)

    with open('src/excel_mcp_server_fastmcp/server.py', 'w', encoding='utf-8') as f:
        f.write(new_content)

    saved = total_before - total_after
    print(f"Processed {count} tools")
    print(f"Before: {total_before} chars (~{total_before//4} tokens)")
    print(f"After:  {total_after} chars (~{total_after//4} tokens)")
    print(f"Saved:  {saved} chars (~{saved//4} tokens, {saved*100//total_before}% reduction)")


if __name__ == '__main__':
    main()
