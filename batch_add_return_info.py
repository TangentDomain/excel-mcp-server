#!/usr/bin/env python3
"""
批量添加缺失的工具返回信息
"""

import re

def add_return_info_to_tools():
    # 读取文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 工具列表及其对应的返回信息模板
    tools_to_update = {
        'excel_batch_insert_rows': {
            'return_info': """**📊 返回信息**:
• **inserted_rows**: 成功插入的行数
• **inserted_position**: 插入的起始行位置
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        },
        'excel_delete_rows': {
            'return_info': """**📊 返回信息**:
• **deleted_rows**: 删除的行数
• **start_row**: 删除的起始行位置
• **end_row**: 删除的结束行位置
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        },
        'excel_delete_columns': {
            'return_info': """**📊 返回信息**:
• **deleted_columns**: 删除的列数
• **start_col**: 删除的起始列位置（字母）
• **end_col**: 删除的结束列位置（字母）
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        },
        'excel_set_formula': {
            'return_info': """**📊 返回信息**:
• **formula_applied**: 应用的公式（原始公式）
• **formula_range**: 公式应用的单元格范围
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        },
        'excel_merge_cells': {
            'return_info': """**📊 返回信息**:
• **merged_range**: 合并的单元格范围
• **cell_count**: 合并的单元格数量
• **merged_value**: 合并后的单元格值
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        },
        'excel_set_row_height': {
            'return_info': """**📊 返回信息**:
• **row_range**: 设置行高的行号范围
• **height**: 设置的行高值（磅）
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        },
        'excel_set_column_width': {
            'return_info': """**📊 返回信息**:
• **column_range**: 设置列宽的列号范围
• **width**: 设置的列宽值（字符数）
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        },
        'excel_compare_sheets': {
            'return_info': """**📊 返回信息**:
• **comparison_result**: 比较结果详情
• **differences_found**: 发现的差异数量
• **difference_details**: 具体差异描述
• **compared_files**: 比较的文件和工作表
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        },
        'excel_server_stats': {
            'return_info': """**📊 返回信息**:
• **tool_calls**: 各工具调用次数统计
• **execution_stats**: 执行时间统计（平均/最慢）
• **error_rates**: 各工具错误率统计
• **global_errors**: 全局错误类型统计
• **server_status**: 服务器运行状态
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        }
    }
    
    # 逐个更新工具
    updated_tools = []
    for tool_name, info in tools_to_update.items():
        # 查找工具的参数说明部分
        pattern = rf'(def {tool_name}.*?{re.escape("**⚡ 使用建议**")})(.*?)(?=def |$)'
        match = re.search(pattern, content, re.DOTALL)
        
        if match:
            # 插入返回信息
            new_content = match.group(1) + f"\n\n{info['return_info']}\n" + match.group(2)
            content = content.replace(match.group(0), new_content)
            updated_tools.append(tool_name)
        else:
            print(f"警告: 未找到工具 {tool_name} 的位置")
    
    # 写回文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"成功更新 {len(updated_tools)} 个工具的返回信息:")
    for tool in updated_tools:
        print(f"  ✅ {tool}")

if __name__ == "__main__":
    add_return_info_to_tools()