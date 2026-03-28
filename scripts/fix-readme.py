#!/usr/bin/env python3
"""
README文档同步脚本
检查并修复README.md和README.en.md的格式问题和版本一致性
"""

import re
from pathlib import Path

def fix_badge_format(content):
    """修复徽章格式，确保正确的closing tags"""
    # 修复缺失的closing bracket
    content = re.sub(r'\[!\[([^\]]+)\]\(([^)]+)\)\s*$', r'[![\1](\2)]', content)
    # 修复缺失的closing bracket for badges
    content = re.sub(r'\[!\[([^\]]+)\]\(([^)]+)\)\s*$', r'[![\1](\2)]', content)
    return content

def fix_version_text(content, version="1.6.37"):
    """修复版本号文本问题，确保版本号只在正确位置出现"""
    # 移除文本中不正确的版本号嵌入（包括v1.6.36和1.6.36的各种变体）
    patterns = [
        r'v1\.6\.36',
        r'1\.6\.36',
        r'Na1\.6\.37',
        r'Ser1\.6\.37',
        r'Nati1\.6\.37',
        r'Co1\.6\.37',
        r'u1\.6\.37',
        r'de1\.6\.37',
        r'De1\.6\.37'
    ]
    
    for pattern in patterns:
        content = re.sub(pattern, version, content)
    
    # 修复版本徽章格式
    content = re.sub(r'\[!\[([^\]]+)\]\(([^)]+)\)\s*\]\[!\[([^\]]*)\]\]\(([^)]+)\)', 
                     r'[![\1](\2)](https://img.shields.io/badge/\3-\4.svg)', content)
    return content

def fix_html_tags(content):
    """修复HTML标签格式"""
    # 确保所有div标签正确闭合
    content = re.sub(r'<div\s+align="center">\s*(.*?)\s*</div>', 
                     lambda m: f'<div align="center">{m.group(1).strip()}</div>', 
                     content, flags=re.DOTALL)
    return content

def clean_extra_spaces(content):
    """清理多余的空格和换行"""
    # 清理多余的空白行
    content = re.sub(r'\n\s*\n\s*\n', '\n\n', content)
    # 清理行尾空格
    content = re.sub(r'\s+\n', '\n', content)
    return content

def fix_readme_content(content):
    """修复README内容的主要问题"""
    # 修复徽章格式
    content = fix_badge_format(content)
    # 修复版本号
    content = fix_version_text(content)
    # 修复HTML标签
    content = fix_html_tags(content)
    # 清理空格
    content = clean_extra_spaces(content)
    return content

def ensure_consistency_between_readmes():
    """确保中英文README版本信息一致"""
    # 读取两个README文件
    with open('README.md', 'r', encoding='utf-8') as f:
        zh_content = f.read()
    
    with open('README.en.md', 'r', encoding='utf-8') as f:
        en_content = f.read()
    
    # 获取版本号
    zh_version = re.search(r'!\[工具数量\].*?v([0-9]+\.[0-9]+\.[0-9]+)', zh_content)
    en_version = re.search(r'!\[Tools\].*?v([0-9]+\.[0-9]+\.[0-9]+)', en_content)
    
    if zh_version and en_version:
        zh_ver = zh_version.group(1)
        en_ver = en_version.group(1)
        
        if zh_ver != en_ver:
            print(f"⚠️ 版本不一致: README.md {zh_ver} vs README.en.md {en_ver}")
            # 使用中文README的版本号作为权威版本
            en_content = fix_version_text(en_content, zh_ver)
            print(f"✅ 已同步版本号至 {zh_ver}")
    
    return zh_content, en_content

def main():
    print("🔍 开始README文档同步和格式修复...")
    
    # 确保两个README版本一致
    zh_content, en_content = ensure_consistency_between_readmes()
    
    # 修复内容问题
    zh_content = fix_readme_content(zh_content)
    en_content = fix_readme_content(en_content)
    
    # 备份并写入修复后的内容
    for filename, content in [('README.md', zh_content), ('README.en.md', en_content)]:
        backup_path = f"{filename}.backup"
        with open(backup_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"✅ {filename} 已修复并备份")
    
    print("📊 README文档同步完成")
    
    return 0

if __name__ == "__main__":
    import sys
    sys.exit(main())