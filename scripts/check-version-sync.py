#!/usr/bin/env python3
"""
版本一致性检查脚本
确保项目中所有版本信息保持同步

检查项目：
1. pyproject.toml
2. src/excel_mcp_server_fastmcp/__init__.py  
3. README.md
4. README.en.md
5. CHANGELOG.md

执行方式：
python3 scripts/check-version-sync.py

输出：
- 发现不一致时自动修复
- 记录到DECISIONS.md
- 生成检查历史
"""

import re
import json
import os
from pathlib import Path
from typing import Dict, Tuple, List, Optional
from datetime import datetime

class VersionChecker:
    """版本一致性检查器"""
    
    def __init__(self, project_root: str = "."):
        self.project_root = Path(project_root)
        self.results = {
            "timestamp": datetime.now().isoformat(),
            "checks": {},
            "inconsistencies": [],
            "auto_fixes": [],
            "current_version": None
        }
        
    def extract_version_from_pyproject(self) -> Optional[str]:
        """从pyproject.toml提取版本"""
        pyproject_path = self.project_root / "pyproject.toml"
        if not pyproject_path.exists():
            return None
            
        with open(pyproject_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # 匹配版本号行
        match = re.search(r'version\s*=\s*["\']([^"\']+)["\']', content)
        if match:
            return match.group(1)
        return None
    
    def extract_version_from_init(self) -> Optional[str]:
        """从__init__.py提取版本"""
        init_path = self.project_root / "src" / "excel_mcp_server_fastmcp" / "__init__.py"
        if not init_path.exists():
            return None
            
        with open(init_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # 匹配 __version__ 变量
        match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
        if match:
            return match.group(1)
        return None
    
    def extract_version_from_readme(self) -> Dict[str, Optional[str]]:
        """从README文件提取版本"""
        versions = {}
        
        # 检查中文README
        readme_zh_path = self.project_root / "README.md"
        if readme_zh_path.exists():
            with open(readme_zh_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            # 查找版本号（通常在标题或安装部分）
            match = re.search(r'v\d+\.\d+\.\d+|版本[:：]\s*v?\d+\.\d+\.\d+', content)
            if match:
                versions["zh"] = match.group(0)
            else:
                versions["zh"] = None
                
        # 检查英文README
        readme_en_path = self.project_root / "README.en.md"
        if readme_en_path.exists():
            with open(readme_en_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            match = re.search(r'v\d+\.\d+\.\d+|version[:：]\s*v?\d+\.\d+\.\d+', content, re.IGNORECASE)
            if match:
                versions["en"] = match.group(0)
            else:
                versions["en"] = None
                
        return versions
    
    def extract_version_from_changelog(self) -> Optional[str]:
        """从CHANGELOG.md提取最新版本"""
        changelog_path = self.project_root / "CHANGELOG.md"
        if not changelog_path.exists():
            return None
            
        with open(changelog_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # 查找第一个版本号（通常是最新版本）
        match = re.search(r'^#+\s*[vV]?\d+\.\d+\.\d+', content, re.MULTILINE)
        if match:
            return match.group(0).strip('#').strip()
        return None
    
    def check_consistency(self) -> bool:
        """检查所有版本信息的一致性"""
        # 提取所有版本信息
        pyproject_ver = self.extract_version_from_pyproject()
        init_ver = self.extract_version_from_init()
        readme_vers = self.extract_version_from_readme()
        changelog_ver = self.extract_version_from_changelog()
        
        # 确定当前版本（优先使用pyproject.toml）
        if pyproject_ver:
            self.results["current_version"] = pyproject_ver
        
        # 记录检查结果
        self.results["checks"] = {
            "pyproject.toml": pyproject_ver,
            "__init__.py": init_ver,
            "README.md": readme_vers.get("zh"),
            "README.en.md": readme_vers.get("en"),
            "CHANGELOG.md": changelog_ver
        }
        
        # 检查一致性
        primary_version = pyproject_ver or init_ver
        inconsistencies = []
        
        if pyproject_ver and init_ver and pyproject_ver != init_ver:
            inconsistencies.append("pyproject.toml和__init__.py版本不一致")
            
        if primary_version:
            if readme_vers.get("zh") and primary_version not in readme_vers["zh"]:
                inconsistencies.append(f"README.md缺少版本信息或版本不匹配: {primary_version}")
            if readme_vers.get("en") and primary_version not in readme_vers["en"]:
                inconsistencies.append(f"README.en.md缺少版本信息或版本不匹配: {primary_version}")
                
        if changelog_ver and primary_version and changelog_ver != primary_version:
            # changelog的版本可能格式不同，需要更宽松的检查
            if not (changelog_ver.replace('v', '') == primary_version.replace('v', '')):
                inconsistencies.append(f"CHANGELOG.md版本与主版本不一致: {primary_version} vs {changelog_ver}")
        
        self.results["inconsistencies"] = inconsistencies
        
        return len(inconsistencies) == 0
    
    def auto_fix(self) -> List[str]:
        """自动修复版本不一致问题"""
        fixes = []
        primary_version = self.results.get("current_version")
        
        if not primary_version:
            return fixes
            
        # 修复__init__.py版本
        init_path = self.project_root / "src" / "excel_mcp_server_fastmcp" / "__init__.py"
        if init_path.exists():
            with open(init_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            old_version = self.extract_version_from_init()
            if old_version and old_version != primary_version:
                new_content = re.sub(
                    r'__version__\s*=\s*["\'][^"\']+["\']',
                    f'__version__ = "{primary_version}"',
                    content
                )
                
                with open(init_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                    
                fixes.append(f"更新__init__.py版本: {old_version} → {primary_version}")
        
        # 修复README.md版本
        readme_zh_path = self.project_root / "README.md"
        if readme_zh_path.exists():
            with open(readme_zh_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            # 查找现有版本并更新
            version_patterns = [
                r'v\d+\.\d+\.\d+',
                r'版本[:：]\s*v?\d+\.\d+\.\d+'
            ]
            
            for pattern in version_patterns:
                matches = re.findall(pattern, content)
                if matches:
                    new_content = re.sub(
                        pattern,
                        primary_version,
                        content
                    )
                    
                    if new_content != content:
                        with open(readme_zh_path, 'w', encoding='utf-8') as f:
                            f.write(new_content)
                        fixes.append(f"更新README.md版本: {matches[0]} → {primary_version}")
                        break
        
        # 修复README.en.md版本
        readme_en_path = self.project_root / "README.en.md"
        if readme_en_path.exists():
            with open(readme_en_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            version_patterns = [
                r'v\d+\.\d+\.\d+',
                r'version[:：]\s*v?\d+\.\d+\.\d+',
                r'VERSION[:：]\s*v?\d+\.\d+\.\d+'
            ]
            
            for pattern in version_patterns:
                matches = re.findall(pattern, content, re.IGNORECASE)
                if matches:
                    new_content = re.sub(
                        pattern,
                        primary_version,
                        content,
                        flags=re.IGNORECASE
                    )
                    
                    if new_content != content:
                        with open(readme_en_path, 'w', encoding='utf-8') as f:
                            f.write(new_content)
                        fixes.append(f"更新README.en.md版本: {matches[0]} → {primary_version}")
                        break
        
        self.results["auto_fixes"] = fixes
        return fixes
    
    def generate_report(self) -> str:
        """生成检查报告"""
        report = []
        report.append(f"版本一致性检查报告 - {self.results['timestamp']}")
        report.append("=" * 50)
        
        if self.results["current_version"]:
            report.append(f"当前主版本: {self.results['current_version']}")
        
        report.append("\n检查结果:")
        for file, version in self.results["checks"].items():
            status = "✅" if version else "❌"
            report.append(f"  {status} {file}: {version or '未找到版本信息'}")
        
        if self.results["inconsistencies"]:
            report.append(f"\n发现 {len(self.results['inconsistencies'])} 个不一致:")
            for issue in self.results["inconsistencies"]:
                report.append(f"  ❌ {issue}")
        else:
            report.append("\n✅ 所有版本信息一致!")
        
        if self.results["auto_fixes"]:
            report.append(f"\n自动修复 {len(self.results['auto_fixes'])} 项:")
            for fix in self.results["auto_fixes"]:
                report.append(f"  🔧 {fix}")
        
        return "\n".join(report)
    
    def save_history(self):
        """保存检查历史"""
        history_dir = self.project_root / ".version_check_history"
        history_dir.mkdir(exist_ok=True)
        
        history_file = history_dir / f"check_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        with open(history_file, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)


def main():
    """主函数"""
    print("🔍 开始版本一致性检查...")
    
    checker = VersionChecker()
    
    # 检查一致性
    is_consistent = checker.check_consistency()
    
    print("\n" + "=" * 60)
    print(checker.generate_report())
    
    if not is_consistent:
        print("\n🔧 尝试自动修复...")
        fixes = checker.auto_fix()
        
        if fixes:
            print("✅ 自动修复完成!")
            print("\n修复后结果:")
            print(checker.generate_report())
            
            # 更新DECISIONS.md记录
            decision_content = f"""
### [版本同步自动化] 第{len(checker.results['auto_fixes'])}次修复
- **时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} UTC
- **检查结果**: 发现{len(checker.results['inconsistencies'])}个不一致
- **自动修复**: {len(fixes)}项成功修复
- **修复内容**: {', '.join(fixes)}
- **效果**: 确保项目中所有版本信息保持同步，提升用户体验
- **依据**: RULES.md版本检查自动化规则
"""
            
            # 追加到DECISIONS.md
            with open("docs/DECISIONS.md", 'a', encoding='utf-8') as f:
                f.write(decision_content)
                
            print(f"\n📝 已记录到 docs/DECISIONS.md")
        
        # 保存检查历史
        checker.save_history()
        print("📊 检查历史已保存到 .version_check_history/")
        
    else:
        print("\n✅ 无需修复，版本信息已同步!")
    
    return 0 if is_consistent else 1


if __name__ == "__main__":
    exit(main())