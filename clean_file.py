#!/usr/bin/env python3
"""
清理损坏的文档行
"""

import re

def clean_file():
    # 读取文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 移除损坏的文档行（使用简单替换）
    lines = content.split('\n')
    cleaned_lines = []
    for line in lines:
        if not '• ****:' in line:
            cleaned_lines.append(line)
    
    # 写回文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'w', encoding='utf-8') as f:
        f.write('\n'.join(cleaned_lines))
    
    print("损坏行清理完成")

if __name__ == "__main__":
    clean_file()