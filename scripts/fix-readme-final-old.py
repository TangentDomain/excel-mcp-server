#!/usr/bin/env python3
"""
README文档同步脚本 - 精确修复版本号嵌入问题
"""

import re
from pathlib import Path

def fix_specific_issues(content):
    """修复特定的版本号嵌入问题"""
    # 精确修复特定的错误模式，确保所有正则表达式都正确转义
    replacements = [
        # Core Features行 - 修复版本号出现在中间的问题
        (r'numerical balancing,\s*1\.6\.37 comparison', 'numerical balancing, version comparison'),
        # Quick Nav行 - 修复标题中的版本号
        (r'## 🔍 Quick\s+1\.6\.37', '## 🔍 Quick Nav'),
        # Game Scenarios锚点 - 修复链接中的版本号
        (r'Game Scenarios\]\(#-game-1\.6\.37', 'Game Scenarios](#-game-development-scenarios'),
        # Architecture行 - 修复架构描述中的版本号
        (r'\*\*Architecture\*\* \| MCP\s+1\.6\.37\s+Mode', '**Architecture** | MCP Server Mode'),
        # AI Integration行 - 修复AI集成描述中的版本号
        (r'\*\*AI Integration\*\* \| ✅\s*1\.6\.37\s*MCP Support', '**AI Integration** | ✅ Native MCP Support'),
        # Test行 - 修复测试覆盖率描述中的版本号
        (r'\*\*Test\s+1\.6\.37\s*\*\*', '**Test Coverage** |'),
        # Installation行 - 修复安装命令中的版本号
        (r'✅ `1\.6\.37 One-line Command', '✅ `uvx excel-mcp-server-fastmcp'),
    ]
    
    for pattern, replacement in replacements:
        content = re.sub(pattern, replacement, content, flags=re.IGNORECASE)
    
    return content

def fix_badge_format(content):
    """修复徽章格式，确保正确的closing tags"""
    # 确保每个badge都有正确的closing bracket
    content = re.sub(r'\[!\[([^\]]+)\]\(([^)]+)\)\s*$', r'[![\1](\2)]', content, flags=re.MULTILINE)
    return content

def fix_html_structure(content):
    """修复HTML结构"""
    # 确保div标签正确闭合
    content = re.sub(r'<div\s+align="center">\s*(.*?)\s*</div>', 
                     lambda m: f'<div align="center">{m.group(1).strip()}</div>', 
                     content, flags=re.DOTALL)
    return content

def clean_formatting(content):
    """清理格式问题"""
    # 清理多余空行
    content = re.sub(r'\n\s*\n\s*\n', '\n\n', content)
    # 清理行尾空格
    content = re.sub(r'\s+\n', '\n', content)
    return content

def sync_readme_versions():
    """同步中英文README版本信息"""
    with open('README.md', 'r', encoding='utf-8') as f:
        zh_content = f.read()
    
    with open('README.en.md', 'r', encoding='utf-8') as f:
        en_content = f.read()
    
    # 从中文README提取正确的版本号
    version_match = re.search(r'!\[工具数量\].*?v([0-9]+\.[0-9]+\.[0-9]+)', zh_content)
    if version_match:
        correct_version = version_match.group(1)
        print(f"✅ 权威版本号: {correct_version}")
    else:
        correct_version = "1.6.37"
        print(f"⚠️ 未找到版本号，使用默认: {correct_version}")
    
    return zh_content, en_content, correct_version

def main():
    print("🔍 开始精确的README文档修复...")
    
    # 同步版本信息
    zh_content, en_content, correct_version = sync_readme_versions()
    
    # 对两个README应用修复
    for filename, content in [('README.md', zh_content), ('README.en.md', en_content)]:
        print(f"📄 修复 {filename}...")
        
        # 应用修复
        content = fix_specific_issues(content)
        content = fix_badge_format(content)
        content = fix_html_structure(content)
        content = clean_formatting(content)
        
        # 备份并写入
        backup_path = f"{filename}.backup"
        with open(backup_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"✅ {filename} 修复完成")
    
    print("📊 README文档修复完成")
    return 0

if __name__ == "__main__":
    import sys
    sys.exit(main())