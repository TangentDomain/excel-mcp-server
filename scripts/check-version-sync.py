#!/usr/bin/env python3
"""
自动化版本检查脚本
<<<<<<< HEAD
检查项目版本一致性，自动修复不一致问题
"""

import re
import toml
=======
确保 pyproject.toml、__init__.py、README.md、README.en.md、CHANGELOG 版本一致性
"""

import re
>>>>>>> feature/REQ-027-evolution
import sys
from pathlib import Path
import shutil

<<<<<<< HEAD
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
=======
def get_version_from_pyproject():
    """从 pyproject.toml 获取版本"""
    pyproject_path = Path("pyproject.toml")
    if not pyproject_path.exists():
        return None
    
    content = pyproject_path.read_text(encoding='utf-8')
    match = re.search(r'version\s*=\s*["\']([^"\']+)["\']', content)
    return match.group(1) if match else None

def get_version_from_init():
    """从 __init__.py 获取版本"""
    init_path = Path("src/excel_mcp_server_fastmcp/__init__.py")
    if not init_path.exists():
        return None
    
    content = init_path.read_text(encoding='utf-8')
    match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
    return match.group(1) if match else None

def get_version_from_readme(readme_path):
    """从 README 获取版本"""
    if not readme_path.exists():
        return None
    
    content = readme_path.read_text(encoding='utf-8')
    
    # 查找版本徽章
    badge_match = re.search(r'badge/version-v([0-9]+\.[0-9]+\.[0-9]+)-blue', content)
    if badge_match:
        return badge_match.group(1)
    
    # 查找当前版本行
    current_match = re.search(r'✅\s*v([0-9]+\.[0-9]+\.[0-9]+)', content)
    if current_match:
        return current_match.group(1)
    
    return None

def get_version_from_changelog():
    """从 CHANGELOG 获取最新版本"""
    changelog_path = Path("CHANGELOG.md")
    if not changelog_path.exists():
        return None
    
    content = changelog_path.read_text(encoding='utf-8')
    
    # 查找最新版本条目（支持 v1.0.0 或 1.0.0 格式）
    lines = content.split('\n')
    for line in lines:
        if line.strip().startswith('## [') and ('v' in line or '[1.' in line):
            match = re.search(r'v?([0-9]+\.[0-9]+\.[0-9]+)', line)
            if match:
                return match.group(1)
    
    return None

def check_version_consistency():
    """检查版本一致性并返回修复建议"""
    versions = {
        'pyproject.toml': get_version_from_pyproject(),
        '__init__.py': get_version_from_init(),
        'README.md': get_version_from_readme(Path("README.md")),
        'README.en.md': get_version_from_readme(Path("README.en.md")),
        'CHANGELOG.md': get_version_from_changelog()
    }
    
    # 过滤掉None值
    valid_versions = {k: v for k, v in versions.items() if v}
    
    if not valid_versions:
        print("❌ 未找到任何版本信息")
        return False
    
    # 检查一致性
    unique_versions = set(valid_versions.values())
    if len(unique_versions) == 1:
        version = list(unique_versions)[0]
        print(f"✅ 版本一致: v{version}")
        return True
    else:
        print("❌ 版本不一致:")
        for file, version in valid_versions.items():
            print(f"  {file}: v{version}")
        return False

def get_latest_version():
    """获取最新版本号"""
    global versions
    versions = [v for v in versions.values() if v]
    if not versions:
        return None
    
    # 简单的版本号比较
    versions.sort(reverse=True)
    return versions[0]

def get_latest_version_from_files():
    """从文件获取最新版本号"""
    # 从pyproject.toml获取版本作为基准
    return get_version_from_pyproject()

def auto_fix_versions():
    """自动修复版本不一致问题"""
    latest_version = get_latest_version_from_files()
    if not latest_version:
        print("❌ 无法确定最新版本号")
        return False
    
    print(f"🔄 自动修复版本到: v{latest_version}")
    files_updated = []
    
    # 更新 pyproject.toml
    pyproject_path = Path("pyproject.toml")
    if pyproject_path.exists():
        content = pyproject_path.read_text(encoding='utf-8')
        new_content = re.sub(r'version\s*=\s*["\'][^"\']+["\']', f'version = "{latest_version}"', content)
        pyproject_path.write_text(new_content, encoding='utf-8')
        files_updated.append("pyproject.toml")
    
    # 更新 __init__.py
    init_path = Path("src/excel_mcp_server_fastmcp/__init__.py")
    if init_path.exists():
        content = init_path.read_text(encoding='utf-8')
        new_content = re.sub(r'__version__\s*=\s*["\'][^"\']+["\']', f'__version__ = "{latest_version}"', content)
        init_path.write_text(new_content, encoding='utf-8')
        files_updated.append("__init__.py")
    
    # 更新 README.md 的版本徽章和当前版本
    readme_path = Path("README.md")
    if readme_path.exists():
        content = readme_path.read_text(encoding='utf-8')
        
        # 更新版本徽章
        content = re.sub(r'badge/version-v([0-9]+\.[0-9]+\.[0-9]+)-blue', f'badge/version-v{latest_version}-blue', content)
        
        # 更新当前版本行
        content = re.sub(r'✅\s*v([0-9]+\.[0-9]+\.[0-9]+)', f'✅ v{latest_version}', content)
        
        readme_path.write_text(content, encoding='utf-8')
        files_updated.append("README.md")
    
    # 更新 README.en.md 的版本徽章和当前版本
    readme_en_path = Path("README.en.md")
    if readme_en_path.exists():
        content = readme_en_path.read_text(encoding='utf-8')
        
        # 更新版本徽章
        content = re.sub(r'badge/version-v([0-9]+\.[0-9]+\.[0-9]+)-blue', f'badge/version-v{latest_version}-blue', content)
        
        # 更新当前版本行
        content = re.sub(r'✅\s*v([0-9]+\.[0-9]+\.[0-9]+)', f'✅ v{latest_version}', content)
        
        readme_en_path.write_text(content, encoding='utf-8')
        files_updated.append("README.en.md")
    
    return files_updated

def main():
    """主函数"""
    print("🔍 开始版本一致性检查...")
    
    if check_version_consistency():
        print("✅ 版本检查通过")
        return 0
    else:
        print("🔧 发现版本不一致，尝试自动修复...")
        updated_files = auto_fix_versions()
        
        if updated_files:
            print(f"✅ 已修复文件: {', '.join(updated_files)}")
            return 0
        else:
            print("❌ 自动修复失败")
            return 1

if __name__ == "__main__":
    sys.exit(main())
>>>>>>> feature/REQ-027-evolution
