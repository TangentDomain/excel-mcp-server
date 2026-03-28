#!/usr/bin/env python3
"""
修复文件中的特殊字符语法错误
"""

import re

def fix_special_characters():
    # 读取文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 定义需要替换的特殊字符模式
    replacements = {
        '，': ',',
        '。': '.',
        '（': '(',
        '）': ')',
        '【': '[',
        '】': ']',
        '：': ':',
        '；': ';',
        '！': '!',
        '？': '?',
        '…': '...',
        '—': '-',
        '～': '~',
        '＆': '&',
        '＊': '*',
        '＋': '+',
        '＝': '=',
        '＜': '<',
        '＞': '>',
        '｜': '|',
        '｛': '{',
        '｝': '}',
        '／': '/',
        '＼': '\\',
        '＄': '$',
        '＃': '#',
        '＠': '@',
        '％': '%',
        '／': '/',
    }
    
    # 执行替换
    for old_char, new_char in replacements.items():
        content = content.replace(old_char, new_char)
    
    # 写回文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print("特殊字符替换完成")

if __name__ == "__main__":
    fix_special_characters()