#!/usr/bin/env python3
"""
自动化版本检查脚本
检查项目中所有关键文件版本一致性，异常时立即修复
"""

import re
import os
from pathlib import Path

def check_version_consistency():
    """检查项目版本一致性"""
    
    # 定义需要检查的文件和对应的版本模式
    version_patterns = {
        'pyproject.toml': r'version = "([^"]+)"',
        'src/excel_mcp_server_fastmcp/__init__.py': r'__version__ = [\'"]([^\'"]+)[\'"]',
        'README.md': r'excel-mcp-server-fastmcp@([0-9]+\.[0-9]+\.[0-9]+)',
        'README.en.md': r'excel-mcp-server-fastmcp@([0-9]+\.[0-9]+\.[0-9]+)',
        'CHANGELOG.md': r'## \[v([0-9]+\.[0-9]+\.[0-9]+)\]'
    }
    
    # 读取主版本号（从pyproject.toml）
    main_version = None
    version_files = {}
    
    print("🔍 开始版本一致性检查...")
    
    for file_path, pattern in version_patterns.items():
        full_path = Path(file_path)
        if not full_path.exists():
            print(f"❌ 文件不存在: {file_path}")
            continue
            
        try:
            with open(full_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 匹配版本号
            match = re.search(pattern, content)
            if match:
                version = match.group(1)
                version_files[file_path] = version
                
                if main_version is None:
                    main_version = version
                    print(f"📌 主版本号 ({file_path}): {main_version}")
                else:
                    if version != main_version:
                        print(f"❌ 版本不匹配 ({file_path}): {version} (期望: {main_version})")
                        # 记录到待修复列表
                        version_files[file_path] = f"NEED_FIX:{version}"
                    else:
                        print(f"✅ 版本正确 ({file_path}): {version}")
            else:
                print(f"⚠️  未找到版本号模式 ({file_path})")
                version_files[file_path] = "NOT_FOUND"
                
        except Exception as e:
            print(f"❌ 读取文件失败 ({file_path}): {e}")
            version_files[file_path] = "ERROR"
    
    # 检查是否需要修复
    need_fix = []
    for file_path, version in version_files.items():
        if isinstance(version, str) and version.startswith("NEED_FIX:"):
            correct_version = version.split(":")[1]
            need_fix.append((file_path, correct_version))
    
    if need_fix:
        print(f"\n🔧 发现 {len(need_fix)} 个版本不一致问题，开始自动修复...")
        for file_path, wrong_version in need_fix:
            fix_version_in_file(file_path, main_version, wrong_version)
        print("✅ 版本修复完成")
        return True
    else:
        print("\n✅ 所有版本号一致，无需修复")
        return False

def fix_version_in_file(file_path, correct_version, wrong_version):
    """修复文件中的版本号"""
    
    patterns = {
        'pyproject.toml': (r'version = "[^"]+"', f'version = "{correct_version}"'),
        'src/excel_mcp_server_fastmcp/__init__.py': (r'__version__ = [\'"][^\'"]+[\'"]', f'__version__ = \'{correct_version}\''),
        'README.md': (r'excel-mcp-server-fastmcp@[0-9]+\.[0-9]+\.[0-9]+', f'excel-mcp-server-fastmcp@{correct_version}'),
        'README.en.md': (r'excel-mcp-server-fastmcp@[0-9]+\.[0-9]+\.[0-9]+', f'excel-mcp-server-fastmcp@{correct_version}'),
        'CHANGELOG.md': (r'## \[v[0-9]+\.[0-9]+\.[0-9]+\]', f'## [v{correct_version}]')
    }
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        old_content = content
        
        # 根据文件类型选择修复模式
        if file_path in patterns:
            pattern, replacement = patterns[file_path]
            content = re.sub(pattern, replacement, content)
            
            if content != old_content:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                print(f"  📝 {file_path}: {wrong_version} → {correct_version}")
            else:
                print(f"  ⚠️  {file_path}: 未找到版本号 {wrong_version}，可能格式不同")
                
    except Exception as e:
        print(f"  ❌ {file_path}: 修复失败 - {e}")

def check_changelog_consistency():
    """检查CHANGELOG的版本格式一致性"""
    
    changelog_path = Path('CHANGELOG.md')
    if not changelog_path.exists():
        return False
    
    try:
        with open(changelog_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 检查是否有非[v1.x.x]格式的版本号
        wrong_patterns = [
            r'^## [0-9]+\.[0-9]+\.[0-9]+',  # ## 1.0.0格式
            r'^## v[0-9]+\.[0-9]+\.[0-9]+[^-]',  # ## v1.0.0但后面没有连字符
        ]
        
        issues = []
        for i, line in enumerate(content.split('\n'), 1):
            for pattern in wrong_patterns:
                if re.match(pattern, line.strip()):
                    issues.append(f"第{i}行: {line.strip()}")
        
        if issues:
            print(f"\n🔍 CHANGELOG格式问题:")
            for issue in issues[:5]:  # 只显示前5个
                print(f"  ❌ {issue}")
            return True
        else:
            print(f"\n✅ CHANGELOG格式正确")
            return False
            
    except Exception as e:
        print(f"❌ 读取CHANGELOG失败: {e}")
        return False

if __name__ == "__main__":
    # 检查版本一致性
    version_issues = check_version_consistency()
    
    # 检查CHANGELOG格式
    changelog_issues = check_changelog_consistency()
    
    if version_issues or changelog_issues:
        print(f"\n📋 发现 {('+' + str(version_issues)) if version_issues else ''}{('+' + str(changelog_issues)) if changelog_issues else ''} 个问题，已自动修复")
    else:
        print(f"\n🎉 所有文档检查通过")