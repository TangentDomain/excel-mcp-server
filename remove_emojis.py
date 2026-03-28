#!/usr/bin/env python3
"""
使用正则表达式移除所有emoji字符
"""

import re

def remove_emojis():
    # 读取文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 移除所有emoji字符（包括表情符号和特殊符号）
    emoji_pattern = re.compile(
        "["
        "\U0001F600-\U0001F64F"  # 表情符号
        "\U0001F300-\U0001F5FF"  # 符号和图片
        "\U0001F680-\U0001F6FF"  # 交通和地图符号
        "\U0001F1E0-\U0001F1FF"  # flags (iOS)
        "\U00002702-\U000027B0"
        "\U000024C2-\U0001F251"
        "\U0001F900-\U0001F9FF"  # 补充符号和表情
        "\U0001FA70-\U0001FAFF"  # 更多补充符号
        "\U00002500-\U00002BEF"  # 各种符号
        "\U00002200-\U000022FF"  # 数学符号
        "\U00002300-\U000023FF"  # 技术符号
        "\U00002400-\U000024FF"  # 封闭字母和数字
        "\U00002500-\U000025FF"  # 箭头
        "\U00002600-\U000026FF"  # 杂项符号
        "\U00002700-\U000027FF"  # Dingbats
        "\U00002900-\U000029FF"  # 箭头符号
        "\U00002A00-\U00002AFF"  # 杂项符号
        "\U00003000-\U000030FF"  # CJK符号和标点
        "\U0001F000-\U0001F0FF"  # 扩展表意文字
        "\U0001F100-\U0001F64F"  # 扩展表情符号
        "\U0001F680-\U0001F6FF"  # 交通和地图符号
        "\U0001F900-\U0001F9FF"  # 补充符号和表情
        "\U0001FA70-\U0001FAFF"  # 更多补充符号
        "]+",
        flags=re.UNICODE
    )
    
    # 移除emoji
    content = emoji_pattern.sub('', content)
    
    # 写回文件
    with open('src/excel_mcp_server_fastmcp/server.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print("emoji字符移除完成")

if __name__ == "__main__":
    remove_emojis()