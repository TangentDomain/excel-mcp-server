#!/usr/bin/env python3
"""
文档同步检查脚本
确保中英文README文档内容一致
"""
import re
import json
from pathlib import Path
from typing import Dict, List, Tuple, Any

class DocumentationSyncChecker:
    def __init__(self, project_dir: str = "."):
        self.project_dir = Path(project_dir)
        self.issues: List[str] = []
        self.synced_items: List[str] = []
        
    def extract_sections_from_readme(self, readme_path: str) -> Dict[str, str]:
        """从README中提取主要部分"""
        sections = {}
        
        try:
            with open(readme_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 提取标题
            h1_pattern = r'^#\s+(.+)$'
            h2_pattern = r'^##\s+(.+)$'
            
            current_h1 = None
            current_h2 = None
            
            for line in content.split('\n'):
                h1_match = re.match(h1_pattern, line)
                h2_match = re.match(h2_pattern, line)
                
                if h1_match:
                    current_h1 = h1_match.group(1)
                    current_h2 = None
                    sections[current_h1] = ""
                elif h2_match and current_h1:
                    current_h2 = h2_match.group(1)
                    sections[f"{current_h1} > {current_h2}"] = ""
                elif current_h1:
                    section_key = current_h2 if current_h2 else current_h1
                    if section_key in sections:
                        sections[section_key] += line + "\n"
            
        except Exception as e:
            self.issues.append(f"读取 {readme_path} 失败: {e}")
        
        return sections
    
    def compare_versions(self) -> bool:
        """比较版本号是否一致"""
        versions = {}
        
        # 从pyproject.toml获取版本
        try:
            with open(self.project_dir / "pyproject.toml", 'r', encoding='utf-8') as f:
                content = f.read()
                version_match = re.search(r'version = "([^"]+)"', content)
                if version_match:
                    versions["pyproject"] = version_match.group(1)
        except:
            self.issues.append("无法读取pyproject.toml版本")
        
        # 从README获取版本
        readme_sections = self.extract_sections_from_readme("README.md")
        for section, content in readme_sections.items():
            if "徽章" in section or "Badge" in section or "version" in section.lower():
                version_match = re.search(r'v([0-9]+\.[0-9]+\.[0-9]+)', content)
                if version_match:
                    versions["readme"] = version_match.group(1)
                    break
        
        # 从英文README获取版本
        readme_en_sections = self.extract_sections_from_readme("README.en.md")
        for section, content in readme_en_sections.items():
            if "Badge" in section or "version" in section.lower():
                version_match = re.search(r'v([0-9]+\.[0-9]+\.[0-9]+)', content)
                if version_match:
                    versions["readme_en"] = version_match.group(1)
                    break
        
        # 检查一致性
        unique_versions = set(v for v in versions.values() if v)
        return len(unique_versions) <= 1
    
    def compare_badge_counts(self) -> Tuple[bool, str]:
        """比较测试用例数量徽章"""
        readme_sections = self.extract_sections_from_readme("README.md")
        readme_en_sections = self.extract_sections_from_readme("README.en.md")
        
        readme_tests = readme_sections.get("徽章", "")
        readme_en_tests = readme_en_sections.get("Badges", "")
        
        # 提取测试数量
        readme_count = re.search(r'(\d+)', readme_tests)
        readme_en_count = re.search(r'(\d+)', readme_en_tests)
        
        readme_num = int(readme_count.group(1)) if readme_count else 0
        readme_en_num = int(readme_en_count.group(1)) if readme_en_count else 0
        
        return readme_num == readme_en_num, f"README测试数: {readme_num}, README.en测试数: {readme_en_num}"
    
    def check_feature_consistency(self) -> List[str]:
        """检查功能描述一致性"""
        inconsistencies = []
        
        readme_sections = self.extract_sections_from_readme("README.md")
        readme_en_sections = self.extract_sections_from_readme("README.en.md")
        
        # 检查主要功能章节
        key_features = [
            "特性", "功能", "Features", "Core Features"
        ]
        
        for zh_feature in key_features:
            if zh_feature in readme_sections:
                # 查找对应的英文功能
                en_feature = None
                for en_key in readme_en_sections.keys():
                    if any(keyword in en_key.lower() for keyword in ["feature", "core", "capability"]):
                        en_feature = en_key
                        break
                
                if en_feature:
                    zh_content = readme_sections[zh_feature].strip()
                    en_content = readme_en_sections[en_feature].strip()
                    
                    # 比较内容长度（粗略的相似度检查）
                    if len(zh_content) > 100 and len(en_content) > 100:
                        if abs(len(zh_content) - len(en_content)) / max(len(zh_content), len(en_content)) > 0.5:
                            inconsistencies.append(f"功能描述长度差异较大: {zh_feature} vs {en_feature}")
        
        return inconsistencies
    
    def generate_sync_report(self) -> Dict[str, Any]:
        """生成同步报告"""
        report = {
            "timestamp": "2026-03-28T07:30:00Z",
            "version_sync": self.compare_versions(),
            "badge_sync": {},
            "feature_consistency": [],
            "overall_status": "✅ 同步良好"
        }
        
        # 检查徽章同步
        badge_sync, badge_info = self.compare_badge_counts()
        report["badge_sync"] = {
            "sync": badge_sync,
            "info": badge_info
        }
        
        # 检查功能一致性
        report["feature_consistency"] = self.check_feature_consistency()
        
        # 确定整体状态
        if not report["version_sync"]:
            report["overall_status"] = "❌ 版本不同步"
        elif not report["badge_sync"]["sync"]:
            report["overall_status"] = "⚠️ 徽章不同步"
        elif len(report["feature_consistency"]) > 0:
            report["overall_status"] = "⚠️ 功能描述不一致"
        
        return report
    
    def save_report(self, report: Dict[str, Any]):
        """保存同步报告"""
        report_path = self.project_dir / "docs-sync-report.json"
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        
        return report_path

def main():
    """主函数"""
    checker = DocumentationSyncChecker()
    
    print("📋 执行文档同步检查...")
    report = checker.generate_sync_report()
    
    # 保存报告
    report_path = checker.save_report(report)
    print(f"📄 同步报告已保存: {report_path}")
    
    # 输出结果
    print(f"\n📊 文档同步状态: {report['overall_status']}")
    print(f"🔄 版本同步: {'✅' if report['version_sync'] else '❌'}")
    print(f"🏷️ 徽章同步: {'✅' if report['badge_sync']['sync'] else '❌'} ({report['badge_sync']['info']})")
    
    if report['feature_consistency']:
        print(f"⚠️ 功能描述不一致:")
        for issue in report['feature_consistency']:
            print(f"  - {issue}")
    
    return 0 if report['overall_status'] == "✅ 同步良好" else 1

if __name__ == "__main__":
    exit(main())