#!/usr/bin/env python3
"""
版本一致性检查脚本
检查 pyproject.toml、__init__.py、README.md、README.en.md 的版本一致性
"""
import json
import re
import sys
from pathlib import Path

def get_version_from_pyproject():
    """从pyproject.toml获取版本"""
    pyproject_path = Path("pyproject.toml")
    if not pyproject_path.exists():
        return None
    
    content = pyproject_path.read_text(encoding='utf-8')
    match = re.search(r'version = "([^"]+)"', content)
    return match.group(1) if match else None

def get_version_from_init():
    """从__init__.py获取版本"""
    init_path = Path("src/excel_mcp_server_fastmcp/__init__.py")
    if not init_path.exists():
        return None
    
    content = init_path.read_text(encoding='utf-8')
    match = re.search(r'__version__ = ["\']([^"\']+)["\']', content)
    return match.group(1) if match else None

def get_version_from_readme():
    """从README.md获取版本"""
    readme_path = Path("README.md")
    if not readme_path.exists():
        return None
    
    content = readme_path.read_text(encoding='utf-8')
    match = re.search(r'\[!\[PyPI\].*v([\d.]+)\]', content)
    return match.group(1) if match else None

def get_version_from_readme_en():
    """从README.en.md获取版本"""
    readme_en_path = Path("README.en.md")
    if not readme_en_path.exists():
        return None
    
    content = readme_en_path.read_text(encoding='utf-8')
    match = re.search(r'\[!\[PyPI\].*v([\d.]+)\]', content)
    return match.group(1) if match else None

def main():
    """主检查函数"""
    print("🔍 版本一致性检查...")
    
    # 获取各文件版本
    versions = {
        'pyproject.toml': get_version_from_pyproject(),
        '__init__.py': get_version_from_init(),
        'README.md': get_version_from_readme(),
        'README.en.md': get_version_from_readme_en(),
    }
    
    print(f"版本信息:")
    for file, version in versions.items():
        print(f"  {file}: {version}")
    
    # 检查一致性
    version_values = [v for v in versions.values() if v is not None]
    if len(set(version_values)) == 1:
        print("✅ 所有版本一致")
        return True
    else:
        print("❌ 版本不一致!")
        print("🔧 开始自动修复...")
        
        # 找到最新版本（通常是最大的版本号）
        latest_version = max(version_values)
        print(f"🎯 目标版本: {latest_version}")
        
        # 修复版本不一致
        fixed_files = []
        for file, current_version in versions.items():
            if current_version != latest_version:
                print(f"📝 修复 {file}: {current_version} → {latest_version}")
                
                if file == 'pyproject.toml':
                    content = Path(file).read_text(encoding='utf-8')
                    new_content = re.sub(
                        r'version = "[^"]+"',
                        f'version = "{latest_version}"',
                        content
                    )
                    Path(file).write_text(new_content, encoding='utf-8')
                    
                elif file == '__init__.py':
                    content = Path(file).read_text(encoding='utf-8')
                    new_content = re.sub(
                        r'__version__ = [\'"][^\'\"][\'"]',
                        f'__version__ = \'{latest_version}\'',
                        content
                    )
                    Path(file).write_text(new_content, encoding='utf-8')
                    
                elif file in ['README.md', 'README.en.md']:
                    content = Path(file).read_text(encoding='utf-8')
                    new_content = re.sub(
                        r'\[!\[PyPI\][^\]]*v[\d.]+[^\]]*\]',
                        f'[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/v{latest_version})',
                        content
                    )
                    Path(file).write_text(new_content, encoding='utf-8')
                
                fixed_files.append(file)
        
        print(f"✅ 修复完成: {', '.join(fixed_files)}")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)