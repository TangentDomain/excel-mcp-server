#!/usr/bin/env python3
"""
自动化版本检查脚本
确保项目各处版本信息一致，自动修复不一致问题
"""
import re
import json
import os
from pathlib import Path
from typing import Dict, Any, List, Tuple

class VersionChecker:
    def __init__(self, project_dir: str = "."):
        self.project_dir = Path(project_dir)
        self.issues: List[str] = []
        self.fixed: List[str] = []
        
    def read_version_from_file(self, file_path: str, pattern: str) -> str:
        """从文件读取版本号"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                match = re.search(pattern, content)
                return match.group(1) if match else ""
        except Exception as e:
            self.issues.append(f"读取文件 {file_path} 失败: {e}")
            return ""
    
    def write_version_to_file(self, file_path: str, pattern: str, new_version: str) -> bool:
        """写入版本号到文件"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            new_content = re.sub(pattern, new_version, content)
            if new_content != content:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                self.fixed.append(f"修复 {file_path}: {pattern} -> {new_version}")
                return True
        except Exception as e:
            self.issues.append(f"写入文件 {file_path} 失败: {e}")
        return False
    
    def check_and_fix_versions(self) -> Tuple[bool, List[str]]:
        """检查并修复版本不一致"""
        versions = {}
        
        # 从各个文件读取版本
        versions['pyproject'] = self.read_version_from_file("pyproject.toml", r'version = "([^"]+)"')
        versions['init'] = self.read_version_from_file("src/excel_mcp_server_fastmcp/__init__.py", r'__version__ = "([^"]+)"')
        versions['readme'] = self.read_version_from_file("README.md", r'v([0-9]+\.[0-9]+\.[0-9]+)')
        versions['readme_en'] = self.read_version_from_file("README.en.md", r'v([0-9]+\.[0-9]+\.[0-9]+)')
        versions['changelog'] = self.read_version_from_file("CHANGELOG.md", r'##\[([^\]]+)\]')
        
        # 检查一致性
        unique_versions = set(v for v in versions.values() if v)
        if len(unique_versions) <= 1:
            return True, []
        
        # 获取主版本（从pyproject.toml）
        main_version = versions['pyproject']
        if not main_version:
            self.issues.append("无法从pyproject.toml获取主版本号")
            return False, self.issues
        
        # 修复不一致的版本
        files_to_fix = [
            ("src/excel_mcp_server_fastmcp/__init__.py", r'__version__ = "([^"]+)"', f'__version__ = "{main_version}"'),
            ("README.md", r'v([^"\s]+)', f'v{main_version}'),
            ("README.en.md", r'v([^"\s]+)', f'v{main_version}'),
        ]
        
        # 只修复非CHANGELOG文件（CHANGELOG需要历史记录）
        for file_path, pattern, new_version in files_to_fix:
            if versions.get(file_path.replace("src/", "").replace("__init__.py", "").replace("/", "").replace("excel_mcp_server_fastmcp", "")) != main_version:
                self.write_version_to_file(file_path, pattern, new_version)
        
        # 更新CHANGELOG，添加新的版本记录
        changelog_version = versions['changelog']
        if changelog_version != main_version:
            self.update_changelog(main_version)
        
        return len(self.issues) == 0, self.issues
    
    def update_changelog(self, version: str):
        """更新CHANGELOG，添加新版本记录"""
        try:
            changelog_path = self.project_dir / "CHANGELOG.md"
            with open(changelog_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 检查是否已存在该版本
            if f'## v{version}' in content:
                return
            
            # 创建新版本记录
            new_entry = f"""
##[{version}] - {self.get_current_date()}
- 自动化版本同步：确保各处版本信息一致性
- 版本号统一：pyproject.toml、__init__.py、README.md、README.en.md

"""
            
            # 在第一个版本号前插入新记录
            pattern = r'##\[([^\]]+)\]'
            match = re.search(pattern, content)
            if match:
                insert_pos = match.start()
                new_content = content[:insert_pos] + new_entry + content[insert_pos:]
                with open(changelog_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                self.fixed.append(f"更新CHANGELOG: 添加[{version}]")
            else:
                self.issues.append("无法在CHANGELOG中找到版本号位置")
        except Exception as e:
            self.issues.append(f"更新CHANGELOG失败: {e}")
    
    def get_current_date(self) -> str:
        """获取当前日期"""
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d")
    
    def cleanup_old_versions(self, keep_latest: int = 5):
        """清理旧版本检查历史文件"""
        cleanup_files = [
            "old_version_check_results.json",
            "version_check_history.json"
        ]
        
        for file_name in cleanup_files:
            file_path = self.project_dir / file_name
            if file_path.exists():
                try:
                    file_path.unlink()
                    self.fixed.append(f"清理旧版本记录: {file_name}")
                except Exception as e:
                    self.issues.append(f"清理 {file_name} 失败: {e}")

def main():
    """主函数"""
    checker = VersionChecker()
    success, issues = checker.check_and_fix_versions()
    
    # 清理旧版本检查历史
    checker.cleanup_old_versions()
    
    # 输出结果
    if success:
        print("✅ 版本检查通过：所有文件版本信息一致")
        if checker.fixed:
            print("🔧 已修复的版本问题:")
            for fix in checker.fixed:
                print(f"  - {fix}")
    else:
        print("❌ 版本检查发现问题:")
        for issue in issues:
            print(f"  - {issue}")
    
    return 0 if success else 1

if __name__ == "__main__":
    exit(main())