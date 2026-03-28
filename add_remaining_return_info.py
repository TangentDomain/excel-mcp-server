#!/usr/bin/env python3
"""
最后处理剩余3个工具的返回信息
"""

def add_remaining_return_info():
    # 手动添加剩余3个工具的返回信息
    
    # 读取文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 为 excel_batch_insert_rows 添加返回信息
    pattern1 = r'(def excel_batch_insert_rows.*?)(\n.*?\n.*?⚡ 使用建议.*?)(\n    """[^"]*""")'
    match1 = re.search(pattern1, content, re.DOTALL)
    if match1:
        return_info1 = """**📊 返回信息**:
• **inserted_rows**: 成功插入的行数
• **inserted_position**: 插入的起始行位置
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        new_content1 = match1.group(1) + "\n" + return_info1 + match1.group(3)
        content = content.replace(match1.group(0), new_content1)
    
    # 为 excel_set_formula 添加返回信息
    pattern2 = r'(def excel_set_formula.*?)(\n.*?\n.*?🔧 参数说明.*?)(\n.*?\n.*?⚡ 使用建议.*?)(\n    """[^"]*""")'
    match2 = re.search(pattern2, content, re.DOTALL)
    if match2:
        return_info2 = """**📊 返回信息**:
• **formula_applied**: 应用的公式（原始公式）
• **formula_range**: 公式应用的单元格范围
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        new_content2 = match2.group(1) + "\n" + return_info2 + match2.group(4)
        content = content.replace(match2.group(0), new_content2)
    
    # 为 excel_set_row_height 添加返回信息
    pattern3 = r'(def excel_set_row_height.*?)(\n.*?\n.*?🔧 参数说明.*?)(\n.*?\n.*?⚡ 使用建议.*?)(\n    """[^"]*""")'
    match3 = re.search(pattern3, content, re.DOTALL)
    if match3:
        return_info3 = """**📊 返回信息**:
• **row_range**: 设置行高的行号范围
• **height**: 设置的行高值（磅）
• **affected_sheet**: 工作表名称
• **success**: 操作是否成功
• **message**: 状态消息或错误信息"""
        new_content3 = match3.group(1) + "\n" + return_info3 + match3.group(4)
        content = content.replace(match3.group(0), new_content3)
    
    # 写回文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print("已添加剩余3个工具的返回信息")

if __name__ == "__main__":
    import re
    add_remaining_return_info()