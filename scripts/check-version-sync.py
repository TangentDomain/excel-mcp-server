#!/usr/bin/env python3
"""
版本一致性检查与自动化修复脚本
自动检测并修复pyproject.toml、__init__.py、README.md、README.en.md、CHANGELOG.md之间的版本信息同步
"""

import re
import json
import sys
from pathlib import Path
from typing import Dict, Optional, Tuple

class VersionSyncChecker:
    def __init__(self, project_root: str = "."):
        self.project_root = Path(project_root)
        self.version_files = {
            "pyproject.toml": self._extract_pyproject_version,
            "src/excel_mcp_server_fastmcp/__init__.py": self._extract_init_version, 
            "README.md": self._extract_readme_version,
            "README.en.md": self._extract_readme_version,
            "CHANGELOG.md": self._extract_changelog_version
        }
        self.versions = {}
        self.inconsistencies = []
        
    def _extract_pyproject_version(self, content: str) -> Optional[str]:
        """从pyproject.toml中提取版本"""
        try:
            import toml
            data = toml.loads(content)
            return data.get("project", {}).get("version")
        except ImportError:
            # fallback to regex if toml not available
            match = re.search(r'version\s*=\s*["\']([^"\']+)["\']', content)
            return match.group(1) if match else None
        except:
            return None
    
    def _extract_init_version(self, content: str) -> Optional[str]:
        """从__init__.py中提取版本"""
        match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
        return match.group(1) if match else None
    
    def _extract_readme_version(self, content: str) -> Optional[str]:
        """从README文件中提取版本"""
        # 查找常见的版本标记模式
        patterns = [
            r'version[:\s]+([^\n\s]+)',
            r'v([0-9]+\.[0-9]+\.[0-9]+)',
            r'excel-mcp-server-fastmcp[^\n]*?([0-9]+\.[0-9]+\.[0-9]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                version = match.group(1)
                # 确保版本格式正确
                if re.match(r'^[0-9]+\.[0-9]+\.[0-9]+$', version):
                    return version
        return None
    
    def _extract_changelog_version(self, content: str) -> Optional[str]:
        """从CHANGELOG.md中提取最新版本"""
        # 查找第一个非[Unreleased]的版本
        lines = content.split('\n')
        for line in lines:
            match = re.search(r'##\s*\[?v?([0-9]+\.[0-9]+\.[0-9]+)', line)
            if match and '[Unreleased]' not in line:
                return match.group(1)
        return None
    
    def read_file(self, filepath: Path) -> Optional[str]:
        """读取文件内容"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            print(f"Warning: Could not read {filepath}: {e}")
            return None
    
    def check_versions(self) -> Dict[str, Optional[str]]:
        """检查所有文件的版本"""
        for filename, extractor in self.version_files.items():
            filepath = self.project_root / filename
            if filepath.exists():
                content = self.read_file(filepath)
                if content:
                    self.versions[filename] = extractor(content)
                    print(f"✓ {filename}: {self.versions[filename]}")
                else:
                    self.versions[filename] = None
                    print(f"✗ {filename}: Could not extract version")
            else:
                self.versions[filename] = None
                print(f"✗ {filename}: File not found")
        
        return self.versions
    
    def find_inconsistencies(self) -> list:
        """发现版本不一致问题"""
        versions = [v for v in self.versions.values() if v is not None]
        if len(set(versions)) <= 1:
            return []  # 所有版本一致
        
        self.inconsistencies = []
        reference_version = self._get_reference_version()
        
        for filename, version in self.versions.items():
            if version and version != reference_version:
                self.inconsistencies.append({
                    "file": filename,
                    "current": version,
                    "expected": reference_version
                })
        
        return self.inconsistencies
    
    def _get_reference_version(self) -> Optional[str]:
        """获取参考版本（优先使用pyproject.toml的版本）"""
        return self.versions.get("pyproject.toml")
    
    def fix_versions(self, target_version: Optional[str] = None) -> bool:
        """修复版本不一致问题"""
        if not target_version:
            target_version = self._get_reference_version()
        
        if not target_version:
            print("✗ Cannot determine target version")
            return False
        
        fixed_files = []
        
        for filename, version in self.versions.items():
            if version and version != target_version:
                filepath = self.project_root / filename
                content = self.read_file(filepath)
                if content:
                    fixed_content = self._fix_file_content(content, target_version, filename)
                    if fixed_content:
                        with open(filepath, 'w', encoding='utf-8') as f:
                            f.write(fixed_content)
                        fixed_files.append(filename)
                        print(f"✓ Fixed {filename}: {version} → {target_version}")
        
        return len(fixed_files) > 0
    
    def _fix_file_content(self, content: str, target_version: str, filename: str) -> Optional[str]:
        """修复文件内容中的版本"""
        if filename == "pyproject.toml":
            # 使用toml库如果可用，否则简单替换
            try:
                import toml
                data = toml.loads(content)
                if data.get("project", {}).get("version") != target_version:
                    data["project"]["version"] = target_version
                    return toml.dumps(data)
            except ImportError:
                content = re.sub(
                    r'version\s*=\s*["\'][^"\']*["\']',
                    f'version = "{target_version}"',
                    content
                )
        
        elif filename == "__init__.py":
            content = re.sub(
                r'__version__\s*=\s*["\'][^"\']*["\']',
                f'__version__ = "{target_version}"',
                content
            )
        
        elif filename in ["README.md", "README.en.md"]:
            # 修复多种版本标记模式
            patterns = [
                (r'version[:\s]+[^\n\s]+', f'version: {target_version}'),
                (r'v([0-9]+\.[0-9]+\.[0-9]+)', f'v{target_version}'),
                (r'excel-mcp-server-fastmcp[^\n]*?([0-9]+\.[0-9]+\.[0-9]+)', 
                 f'excel-mcp-server-fastmcp v{target_version}')
            ]
            
            for pattern, replacement in patterns:
                content = re.sub(pattern, replacement, content, flags=re.IGNORECASE)
        
        elif filename == "CHANGELOG.md":
            # 确保第一个非Unreleased版本号正确
            content = re.sub(
                r'##\s*(?:\[Unreleased\]|v[0-9]+\.[0-9]+\.[0-9]+)',
                f'## v{target_version}',
                content,
                count=1
            )
        
        return content
    
    def generate_report(self) -> str:
        """生成检查报告"""
        report = []
        report.append("=== 版本一致性检查报告 ===")
        report.append("")
        
        # 检查结果
        report.append("版本检查结果:")
        for filename, version in self.versions.items():
            status = "✓" if version else "✗"
            report.append(f"  {status} {filename}: {version}")
        
        report.append("")
        
        # 不一致问题
        if self.inconsistencies:
            report.append("发现版本不一致问题:")
            for issue in self.inconsistencies:
                report.append(f"  ✗ {issue['file']}: {issue['current']} → {issue['expected']}")
        else:
            report.append("✓ 所有文件版本一致")
        
        return "\n".join(report)

def main():
    """主函数"""
    print("开始版本一致性检查...")
    
    checker = VersionSyncChecker()
    
    # 检查版本
    versions = checker.check_versions()
    
    # 发现不一致
    inconsistencies = checker.find_inconsistencies()
    
    # 生成报告
    report = checker.generate_report()
    print(report)
    
    # 如果有不一致，询问是否修复
    if inconsistencies:
        print("\n发现版本不一致，尝试自动修复...")
        if checker.fix_versions():
            print("✓ 版本修复完成")
            # 重新检查
            checker.check_versions()
            checker.find_inconsistencies()
            final_report = checker.generate_report()
            print("\n最终检查结果:")
            print(final_report)
        else:
            print("✗ 版本修复失败")
            return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())