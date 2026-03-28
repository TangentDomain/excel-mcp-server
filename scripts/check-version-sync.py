#!/usr/bin/env python3
"""
自动化版本一致性检查脚本
检查并修复pyproject.toml、__init__.py、README.md、README.en.md中的版本不一致问题
"""

import re
import json
import sys
from pathlib import Path
from datetime import datetime

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent
DECISIONS_FILE = PROJECT_ROOT / "docs" / "DECISIONS.md"

# 要检查的文件及其版本模式
FILES_TO_CHECK = {
    "pyproject.toml": r'version\s*=\s*["\']([^"\']+)["\']',
    "src/excel_mcp_server_fastmcp/__init__.py": r'__version__\s*=\s*["\']([^"\']+)["\']',
    "README.md": r'pypi\.org/project/[^/]+/v([^"\'\s\)]+)',
    "README.en.md": r'pypi\.org/project/[^/]+/v([^"\'\s\)]+)'
}

def get_version_from_file(file_path, pattern):
    """从文件中提取版本号"""
    try:
        content = file_path.read_text(encoding='utf-8')
        match = re.search(pattern, content)
        if match:
            return match.group(1)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return None

def check_and_fix_versions():
    """检查并修复版本一致性"""
    print("🔍 检查版本一致性...")
    
    versions = {}
    inconsistencies = []
    
    # 获取所有文件的版本
    for file_name, pattern in FILES_TO_CHECK.items():
        file_path = PROJECT_ROOT / file_name
        if file_path.exists():
            version = get_version_from_file(file_path, pattern)
            versions[file_name] = version
            if version:
                print(f"  ✅ {file_name}: v{version}")
            else:
                print(f"  ❌ {file_name}: 无法提取版本")
        else:
            print(f"  ⚠️  {file_name}: 文件不存在")
    
    # 检查一致性
    unique_versions = set(v for v in versions.values() if v is not None)
    if len(unique_versions) == 1:
        current_version = list(unique_versions)[0]
        print(f"🎉 所有文件版本一致: v{current_version}")
        return True, None
    else:
        print("❌ 发现版本不一致:")
        for file_name, version in versions.items():
            print(f"  - {file_name}: {version or 'N/A'}")
        
        # 确定主版本（pyproject.toml为基准）
        main_version = versions.get("pyproject.toml")
        if main_version:
            print(f"🔧 以pyproject.toml为基准，修复其他文件为 v{main_version}")
            
            # 修复不一致的文件
            for file_name, pattern in FILES_TO_CHECK.items():
                if file_name in versions and versions[file_name] != main_version:
                    file_path = PROJECT_ROOT / file_name
                    if file_path.exists():
                        content = file_path.read_text(encoding='utf-8')
                        
                        # 替换版本号
                        new_content = re.sub(pattern, f'\\g<0>'.replace(versions[file_name], main_version), content)
                        
                        if new_content != content:
                            file_path.write_text(new_content, encoding='utf-8')
                            print(f"  📝 修复 {file_name}: {versions[file_name]} → {main_version}")
                            inconsistencies.append(f"{file_name}: {versions[file_name]} → {main_version}")
                        else:
                            print(f"  ⚠️  {file_name}: 版本修复失败，请检查格式")
        
        return False, inconsistencies

def log_to_decisions(message):
    """记录决策到DECISIONS.md"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M UTC")
    log_entry = f"- **版本同步自动修复**: {message} - {timestamp}\n"
    
    # 读取现有内容
    if DECISIONS_FILE.exists():
        existing_content = DECISIONS_FILE.read_text(encoding='utf-8')
        
        # 在末尾添加新条目
        if existing_content.strip() and not existing_content.endswith('\n'):
            existing_content += '\n\n'
        existing_content += log_entry
        
        DECISIONS_FILE.write_text(existing_content, encoding='utf-8')
    else:
        DECISIONS_FILE.write_text(log_entry, encoding='utf-8')

def main():
    """主函数"""
    print(f"🚀 自动版本一致性检查 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    try:
        # 检查并修复
        is_consistent, inconsistencies = check_and_fix_versions()
        
        if not is_consistent and inconsistencies:
            # 记录修复日志
            fix_summary = f"修复版本不一致: {', '.join(inconsistencies)}"
            log_to_decisions(fix_summary)
            print(f"📝 已记录修复日志到 docs/DECISIONS.md")
        
        print(f"✅ 版本检查完成")
        return 0 if is_consistent else 1
        
    except Exception as e:
        print(f"❌ 版本检查失败: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())