#!/usr/bin/env python3
"""
精确地为16个工具添加返回信息描述
不修改文件的其他部分，只插入返回信息
"""

def add_return_info():
    with open('src/excel_mcp_server_fastmcp/server.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 定义16个工具的返回信息，使用原始文件的emoji风格
    return_info_map = {
        'excel_export_to_csv': '\n**📊 返回信息**:\n• **exported_path**: 导出的CSV文件完整路径\n• **original_sheet**: 原始工作表名称\n• **encoding**: 使用的编码格式\n• **file_size**: 导出后的文件大小（字节）\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_import_from_csv': '\n**📊 返回信息**:\n• **output_path**: 生成的Excel文件完整路径\n• **imported_rows**: 导入的行数（含标题行）\n• **imported_columns**: 导入的列数\n• **sheet_name**: 工作表名称\n• **encoding**: 使用的编码格式\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_convert_format': '\n**📊 返回信息**:\n• **input_format**: 输入文件原始格式\n• **output_format**: 输出文件目标格式\n• **converted_path**: 转换后的文件完整路径\n• **sheet_count**: 工作表数量（Excel相关格式）\n• **row_count**: 总行数\n• **column_count**: 总列数\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_merge_files': '\n**📊 返回信息**:\n• **merge_mode**: 使用的合并模式（sheets/append/horizontal）\n• **input_count**: 输入文件数量\n• **input_files**: 输入文件路径列表\n• **output_path**: 合并后的输出文件路径\n• **merged_sheets**: 合并后包含的工作表数量\n• **total_rows**: 合并后的总行数\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_create_sheet': '\n**📊 返回信息**:\n• **sheet_name**: 创建的工作表名称\n• **sheet_index**: 工作表在文件中的位置（0-based）\n• **total_sheets**: 文件中总的工作表数量\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_copy_sheet': '\n**📊 返回信息**:\n• **source_sheet**: 源工作表名称\n• **copied_sheet**: 复制后的工作表名称\n• **new_index**: 新工作表在文件中的位置（0-based）\n• **copied_rows**: 复制的行数\n• **copied_columns**: 复制的列数\n• **streaming_used**: 是否使用流式复制\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_rename_column': '\n**📊 返回信息**:\n• **old_header**: 原始列名\n• **new_header**: 新列名\n• **header_row**: 修改的表头行号\n• **sheet_name**: 工作表名称\n• **success**: 重命名是否成功\n• **message**: 操作结果说明\n',
        'excel_batch_insert_rows': '\n**📊 返回信息**:\n• **inserted_rows**: 成功插入的行数\n• **inserted_position**: 插入的起始行位置\n• **affected_sheet**: 工作表名称\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_delete_rows': '\n**📊 返回信息**:\n• **deleted_rows**: 删除的行数\n• **start_row**: 删除的起始行位置\n• **end_row**: 删除的结束行位置\n• **affected_sheet**: 工作表名称\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_delete_columns': '\n**📊 返回信息**:\n• **deleted_columns**: 删除的列数\n• **start_col**: 删除的起始列位置（字母）\n• **end_col**: 删除的结束列位置（字母）\n• **affected_sheet**: 工作表名称\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_set_formula': '\n**📊 返回信息**:\n• **formula_applied**: 应用的公式（原始公式）\n• **formula_range**: 公式应用的单元格范围\n• **affected_sheet**: 工作表名称\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_merge_cells': '\n**📊 返回信息**:\n• **merged_range**: 合并的单元格范围\n• **cell_count**: 合并的单元格数量\n• **merged_value**: 合并后的单元格值\n• **affected_sheet**: 工作表名称\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_set_row_height': '\n**📊 返回信息**:\n• **row_range**: 设置行高的行号范围\n• **height**: 设置的行高值（磅）\n• **affected_sheet**: 工作表名称\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_set_column_width': '\n**📊 返回信息**:\n• **column_range**: 设置列宽的列号范围\n• **width**: 设置的列宽值（字符数）\n• **affected_sheet**: 工作表名称\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_compare_sheets': '\n**📊 返回信息**:\n• **comparison_result**: 比较结果详情\n• **differences_found**: 发现的差异数量\n• **added_records**: 新增的记录数量\n• **deleted_records**: 删除的记录数量\n• **modified_records**: 修改的记录数量\n• **unchanged_records**: 未变的记录数量\n• **comparison_summary**: 对比总结信息\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
        'excel_server_stats': '\n**📊 返回信息**:\n• **tool_calls**: 各工具调用次数统计\n• **execution_stats**: 执行时间统计（平均耗时、最慢工具）\n• **error_rates**: 各工具错误率统计\n• **global_errors**: 全局错误类型统计（按错误分类的计数）\n• **server_status**: 服务器运行状态（内存、连接等）\n• **performance_metrics**: 性能指标（QPS、响应时间等）\n• **success**: 操作是否成功\n• **message**: 状态消息或错误信息\n',
    }
    
    import re
    
    # 逐个工具处理
    updated_count = 0
    for tool_name, return_info in return_info_map.items():
        # 查找每个工具函数的docstring结束位置（三引号结束前）
        # 模式：找到函数def..."""..."""之间的内容
        pattern = rf'(def {tool_name}\([^)]*\)[^:]*:.*?"""(.*?)""")'
        match = re.search(pattern, content, re.DOTALL)
        
        if match:
            # 获取整个匹配（函数定义+docstring）
            full_match = match.group(0)
            # 检查是否已有返回信息
            if '📊 返回信息' in full_match:
                print(f"  跳过 {tool_name}: 已有返回信息")
                continue
            
            # 在最后一个"""前插入返回信息
            # 找到最后一个"""
            last_quote_pos = full_match.rfind('"""')
            if last_quote_pos > 0:
                # 在"""之前插入返回信息
                new_match = full_match[:last_quote_pos] + return_info + '"""'
                content = content.replace(full_match, new_match)
                updated_count += 1
                print(f"  ✓ {tool_name}")
            else:
                print(f"  ✗ {tool_name}: 未找到docstring结束位置")
        else:
            print(f"  ✗ {tool_name}: 未找到函数定义")
    
    # 写回文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"\n共更新 {updated_count} 个工具的返回信息")

if __name__ == "__main__":
    add_return_info()