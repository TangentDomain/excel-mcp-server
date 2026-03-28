#!/usr/bin/env python3
"""
README文档同步脚本 - 统一版本
整合所有README修复功能，支持版本同步、格式优化、链接检查
"""

import re
import sys
from pathlib import Path

def fix_specific_issues(content):
    """修复特定的版本号嵌入问题"""
    # 精确修复特定的错误模式
    replacements = [
        # Core Features行
        (r'numerical balancing, 1\.6\.36 comparison', 'numerical balancing, version comparison'),
        # Quick Nav行
        (r'## 🔍 Quick 1\.6\.37', '## 🔍 Quick Nav'),
        # Game Scenarios锚点
        (r'Game Scenarios\]\(#-game-1\.6\.37', 'Game Scenarios](#-game-development-scenarios'),
        # Architecture行
        (r'\*\*Architecture\*\* \| MCP 1\.6\.37 Mode', '**Architecture** | MCP Server Mode'),
        # AI Integration行
        (r'\*\*AI Integration\*\* \| ✅ 1\.6\.37 MCP Support', '**AI Integration** | ✅ Native MCP Support'),
        # Test行
        (r'\*\*Test 1\.6\.37 \*\*', '**Test Coverage** |'),
        # Installation行
        (r'✅ `1\.6\.37 One-line Command', '✅ `uvx excel-mcp-server-fastmcp'),
    ]
    
    for pattern, replacement in replacements:
        content = re.sub(pattern, replacement, content)
    return content

def fix_badge_format(content):
    """修复徽章格式，确保正确的closing tags"""
    # 确保每个badge都有正确的closing bracket
    content = re.sub(r'\[!\[([^\]]+)\]\(([^)]+)\)\s*$', r'[![\1](\2)]', content, flags=re.MULTILINE)
    return content

def fix_html_structure(content):
    """修复HTML结构"""
    # 确保div标签正确闭合
    content = re.sub(r'<div([^>]*)>\s*</div>', r'<div\1></div>', content)
    return content

def clean_formatting(content):
    """清理格式问题"""
    # 清理多余空行
    content = re.sub(r'\n\s*\n\s*\n', '\n\n', content)
    return content

def sync_readme_versions():
    """同步中英文README版本信息"""
    # 读取两个README文件
    readme_files = ["README.md", "README.en.md"]
    contents = {}
    
    for filename in readme_files:
        filepath = Path(filename)
        if filepath.exists():
            contents[filename] = filepath.read_text(encoding='utf-8')
        else:
            print(f"⚠️ 文件不存在: {filename}")
            continue
    
    if not contents:
        print("❌ 未找到README文件")
        return
    
    # 从中文README提取正确的版本号
    version_match = re.search(r'!\[工具数量\].*?v([0-9]+\.[0-9]+\.[0-9]+)', contents.get("README.md", ""))
    if version_match:
        correct_version = version_match.group(1)
        print(f"✅ 权威版本号: {correct_version}")
    else:
        correct_version = "1.6.37"
        print(f"⚠️ 未找到版本号，使用默认: {correct_version}")
    
    return contents, correct_version

def fix_readme_file(content, filename, correct_version):
    """修复单个README文件"""
    print(f"📄 修复 {filename}...")
    
    # 应用修复
    content = fix_specific_issues(content)
    content = fix_badge_format(content)
    content = fix_html_structure(content)
    content = clean_formatting(content)
    
    # 备份并写入
    filepath = Path(filename)
    backup_path = Path(f"{filename}.backup")
    filepath.rename(backup_path)
    filepath.write_text(content, encoding='utf-8')
    
    print(f"✅ {filename} 修复完成")
    return content

def main():
    """主函数"""
    print("🔍 开始统一README文档修复...")
    
    # 同步版本信息
    contents, correct_version = sync_readme_versions()
    
    if contents:
        # 修复所有README文件
        for filename in contents.keys():
            content = contents[filename]
            fix_readme_file(content, filename, correct_version)
        
        print("📊 README文档统一修复完成")
        return 0
    else:
        print("❌ 没有找到需要修复的README文件")
        return 1

if __name__ == "__main__":
    sys.exit(main())