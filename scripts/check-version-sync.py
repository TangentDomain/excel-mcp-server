#!/usr/bin/env python3
"""
自动化版本检查脚本
检查项目版本一致性，自动修复不一致问题
"""

import re
import toml
import sys
from pathlib import Path
import shutil

def get_version_from_toml():
    """从pyproject.toml获取版本"""
    try:
        with open("pyproject.toml", "r", encoding="utf-8") as f:
            data = toml.load(f)
            return data["project"]["version"]
    except Exception as e:
        print(f"❌ 读取pyproject.toml失败: {e}")
        return None

def get_version_from_init():
    """从__init__.py获取版本"""
    try:
        with open("src/excel_mcp_server_fastmcp/__init__.py", "r", encoding="utf-8") as f:
            content = f.read()
            match = re.search(r'__version__ = ["\']([^"\']+)["\']', content)
            if match:
                return match.group(1)
    except Exception as e:
        print(f"❌ 读取__init__.py失败: {e}")
    return None

def get_version_from_readme_en():
    """从README.en.md获取版本（只获取MCP相关版本信息）"""
    try:
        with open("README.en.md", "r", encoding="utf-8") as f:
            content = f.read()
            # 只查找uvx命令相关的版本信息
            match = re.search(r'uvx excel-mcp-server-fastmcp.*?@([^\s\)]+)', content)
            if match:
                return match.group(1)
    except Exception as e:
        print(f"❌ 读取README.en.md失败: {e}")
    return None

def get_version_from_readme():
    """从README.md获取版本（只获取MCP相关版本信息）"""
    try:
        with open("README.md", "r", encoding="utf-8") as f:
            content = f.read()
            # 只查找uvx命令相关的版本信息
            match = re.search(r'uvx excel-mcp-server-fastmcp.*?@([^\s\)]+)', content)
            if match:
                return match.group(1)
    except Exception as e:
        print(f"❌ 读取README.md失败: {e}")
    return None

def get_version_from_changelog():
    """从CHANGELOG获取最新版本"""
    try:
        with open("CHANGELOG.md", "r", encoding="utf-8") as f:
            content = f.read()
            # 查找第一个版本号，支持v1.6.41和1.6.41两种格式
            match = re.search(r'## \[v?([^\]]+)\]', content)
            if match:
                return match.group(1)
    except Exception as e:
        print(f"❌ 读取CHANGELOG.md失败: {e}")
    return None

def fix_version_consistency(target_version):
    """修复版本一致性"""
    changes = []
    
    # 检查并修复 __init__.py
    init_version = get_version_from_init()
    if init_version != target_version:
        try:
            with open("src/excel_mcp_server_fastmcp/__init__.py", "r", encoding="utf-8") as f:
                content = f.read()
            
            new_content = re.sub(
                r'__version__ = [\'"][^\'\"][\'"]',
                f'__version__ = \'{target_version}\'',
                content
            )
            
            with open("src/excel_mcp_server_fastmcp/__init__.py", "w", encoding="utf-8") as f:
                f.write(new_content)
            
            changes.append(f"__init__.py: {init_version} → {target_version}")
        except Exception as e:
            print(f"❌ 修复__init__.py失败: {e}")
    
    # 检查并修复 README.en.md
    readme_en_version = get_version_from_readme_en()
    if readme_en_version != target_version:
        try:
            with open("README.en.md", "r", encoding="utf-8") as f:
                content = f.read()
            
            new_content = re.sub(
                r'uvx excel-mcp-server-fastmcp.*?@([^\s\)]+)',
                f'uvx excel-mcp-server-fastmcp@{target_version}',
                content
            )
            
            with open("README.en.md", "w", encoding="utf-8") as f:
                f.write(new_content)
            
            changes.append(f"README.en.md: {readme_en_version} → {target_version}")
        except Exception as e:
            print(f"❌ 修复README.en.md失败: {e}")
    
    # 检查并修复 README.md
    readme_version = get_version_from_readme()
    if readme_version != target_version:
        try:
            with open("README.md", "r", encoding="utf-8") as f:
                content = f.read()
            
            new_content = re.sub(
                r'uvx excel-mcp-server-fastmcp.*?@([^\s\)]+)',
                f'uvx excel-mcp-server-fastmcp@{target_version}',
                content
            )
            
            with open("README.md", "w", encoding="utf-8") as f:
                f.write(new_content)
            
            changes.append(f"README.md: {readme_version} → {target_version}")
        except Exception as e:
            print(f"❌ 修复README.md失败: {e}")
    
    return changes

def main():
    """主函数"""
    print("🔍 开始自动化版本检查...")
    
    # 获取pyproject.toml版本作为基准
    base_version = get_version_from_toml()
    if not base_version:
        print("❌ 无法获取基准版本，退出")
        sys.exit(1)
    
    print(f"📋 基准版本（pyproject.toml）: {base_version}")
    
    # 检查各文件版本
    versions = {
        "pyproject.toml": base_version,
        "__init__.py": get_version_from_init(),
        "README.en.md": get_version_from_readme_en(),
        "README.md": get_version_from_readme(),
        "CHANGELOG.md": get_version_from_changelog()
    }
    
    print("\n📊 版本检查结果:")
    inconsistent = []
    for file, version in versions.items():
        status = "✅" if version == base_version else "❌"
        print(f"  {status} {file}: {version or '未找到'}")
        if version != base_version:
            inconsistent.append(file)
    
    # 处理不一致的版本
    if inconsistent:
        print(f"\n⚠️  发现 {len(inconsistent)} 个文件版本不一致，开始修复...")
        changes = fix_version_consistency(base_version)
        
        if changes:
            print("\n✅ 版本修复完成:")
            for change in changes:
                print(f"  📝 {change}")
            
            # 记录到DECISIONS.md
            decision_entry = f"[自我进化建议] 版本一致性检查与自动化修复 → 创建check-version-sync.py脚本，自动检测并修复版本不一致问题"
            with open("docs/DECISIONS.md", "a", encoding="utf-8") as f:
                f.write(f"- {decision_entry}\n")
            
            print(f"\n📝 已记录修复操作到docs/DECISIONS.md")
        else:
            print("\n❌ 修复失败")
    else:
        print("\n✅ 所有文件版本一致，无需修复")

if __name__ == "__main__":
    main()