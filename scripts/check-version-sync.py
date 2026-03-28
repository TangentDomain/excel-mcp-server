#!/usr/bin/env python3
"""
自动化版本一致性检查脚本
检查项目：pyproject.toml、__init__.py、README.md、README.en.md、CHANGELOG
发现不一致时自动修复并记录到DECISIONS.md
"""

import re
import sys
from pathlib import Path
import toml

def get_version_from_pyproject():
    """从pyproject.toml获取版本"""
    try:
        with open('pyproject.toml', 'r', encoding='utf-8') as f:
            data = toml.load(f)
            return data['project']['version']
    except Exception as e:
        print(f"❌ 读取pyproject.toml失败: {e}")
        return None

def get_version_from_init():
    """从__init__.py获取版本"""
    try:
        with open('src/excel_mcp_server_fastmcp/__init__.py', 'r', encoding='utf-8') as f:
            content = f.read()
            match = re.search(r'__version__ = ["\']([^"\']+)["\']', content)
            if match:
                return match.group(1)
    except Exception as e:
        print(f"❌ 读取__init__.py失败: {e}")
    return None

def get_version_from_readme(readme_path):
    """从README获取版本"""
    try:
        with open(readme_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # 查找版本徽章中的版本号
            badge_patterns = [
                r'excel-mcp-server-fastmcp@v?([^\]\s\)]+)',
                r'v([0-9]+\.[0-9]+\.[0-9]+)',
                r'excel-mcp-server-fastmcp/[^/]+/[^/]+/badge/v([0-9]+\.[0-9]+\.[0-9]+)'
            ]
            for pattern in badge_patterns:
                match = re.search(pattern, content)
                if match:
                    return match.group(1)
    except Exception as e:
        print(f"❌ 读取{readme_path}失败: {e}")
    return None

def update_version_in_file(file_path, old_version, new_version):
    """更新文件中的版本号"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 备份原文件
        backup_path = f"{file_path}.backup"
        with open(backup_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        # 对于CHANGELOG.md，需要特殊处理格式
        if file_path == 'CHANGELOG.md':
            # 将 ## [版本] 格式更新到最新版本
            pattern = r'##\s*\[' + re.escape(old_version) + r'\]'
            updated_content = re.sub(pattern, f'## [{new_version}]', content)
        else:
            # 其他文件直接替换
            updated_content = content.replace(old_version, new_version)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(updated_content)
        
        print(f"✅ {file_path} 版本已更新: {old_version} → {new_version}")
        return True
    except Exception as e:
        print(f"❌ 更新{file_path}失败: {e}")
        return False

def get_latest_changelog_version():
    """从CHANGELOG获取最新版本"""
    changelog_path = Path('CHANGELOG.md')
    if not changelog_path.exists():
        return None
    
    try:
        with open(changelog_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # 查找第一个版本号（支持 ## [1.6.36] 和 ## v1.6.36 格式）
            patterns = [
                r'##\s*\[([0-9]+\.[0-9]+\.[0-9]+)\]',  # ## [1.6.36]
                r'##\s*v([0-9]+\.[0-9]+\.[0-9]+)'       # ## v1.6.36
            ]
            for pattern in patterns:
                match = re.search(pattern, content)
                if match:
                    return match.group(1)
    except Exception as e:
        print(f"❌ 读取CHANGELOG.md失败: {e}")
    return None

def main():
    print("🔍 开始自动化版本一致性检查...")
    
    # 获取各文件版本
    pyproject_version = get_version_from_pyproject()
    init_version = get_version_from_init() 
    readme_version = get_version_from_readme('README.md')
    readme_en_version = get_version_from_readme('README.en.md')
    changelog_version = get_latest_changelog_version()
    
    print(f"📋 版本检查结果:")
    print(f"  pyproject.toml:     {pyproject_version}")
    print(f"  __init__.py:       {init_version}")
    print(f"  README.md:          {readme_version}")
    print(f"  README.en.md:       {readme_en_version}")
    print(f"  CHANGELOG.md:       {changelog_version}")
    
    # 确定权威版本（以pyproject.toml为准）
    if not pyproject_version:
        print("❌ pyproject.toml没有找到版本号，退出")
        return 1
    
    versions = {
        'pyproject.toml': pyproject_version,
        '__init__.py': init_version,
        'README.md': readme_version,
        'README.en.md': readme_en_version,
        'CHANGELOG.md': changelog_version
    }
    
    # 检查一致性
    inconsistent_files = []
    for file, version in versions.items():
        if version != pyproject_version:
            inconsistent_files.append((file, version, pyproject_version))
    
    if not inconsistent_files:
        print("✅ 所有文件版本一致，无需修复")
        return 0
    
    print(f"⚠️ 发现 {len(inconsistent_files)} 个文件版本不一致:")
    for file, old, new in inconsistent_files:
        print(f"  📄 {file}: {old} → {new}")
    
    # 修复不一致的文件
    fixed_files = []
    decision_log = []
    
    for file, old_version, new_version in inconsistent_files:
        if file == 'pyproject.toml':
            continue  # 不修改权威版本
        
        file_path_map = {
            '__init__.py': 'src/excel_mcp_server_fastmcp/__init__.py',
            'README.md': 'README.md',
            'README.en.md': 'README.en.md',
            'CHANGELOG.md': 'CHANGELOG.md'
        }
        
        actual_path = file_path_map.get(file)
        if actual_path and update_version_in_file(actual_path, old_version, new_version):
            fixed_files.append(file)
            decision_log.append(f"修复{file}版本号: {old_version} → {new_version}")
    
    # 记录到DECISIONS.md
    if fixed_files:
        timestamp = "2026-03-28 08:25 UTC"
        decision_text = f"[第195轮] {timestamp}\n[自动化版本检查] 发现并修复版本不一致\n"
        
        for file, old, new in inconsistent_files:
            if file in fixed_files:
                decision_text += f"• **问题**: {file}版本号{old}与权威版本{pyproject_version}不一致\n"
                decision_text += f"• **修复**: 更新{file}版本号至{pyproject_version}\n"
        
        decision_text += f"• **影响**: 确保项目版本信息统一，提升用户体验\n• **状态**: ✅ 已修复\n\n"
        
        # 追加到DECISIONS.md
        with open('docs/DECISIONS.md', 'a', encoding='utf-8') as f:
            f.write(decision_text)
        
        print(f"✅ 版本修复完成，已记录到docs/DECISIONS.md")
        print(f"📊 修复文件数: {len(fixed_files)}")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())