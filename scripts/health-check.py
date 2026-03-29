#!/usr/bin/env python3
"""
项目健康度自检脚本
根据RULES.md要求，每20轮至少执行1次健康检查

检查项：
1. 根目录垃圾文件 - 临时脚本、测试文件散落
2. 测试冗余 - 多个测试文件测同一功能
3. 轮次编号测试文件 - test_req010_r67.py应合并后删除
4. 文档膨胀 - DECISIONS/NOW/REQUIREMENTS是否超限
5. 废弃分支 - 过期worktree未清理
6. 依赖变化 - pyproject.toml不该引入的新依赖

执行方式：
python3 scripts/health-check.py

输出：
- 发现问题立即修复
- 记录到DECISIONS.md
- 生成健康度报告
"""

import os
import re
import glob
import json
import shutil
from pathlib import Path
from typing import List, Dict, Tuple, Set
from datetime import datetime

class ProjectHealthChecker:
    """项目健康度检查器"""
    
    def __init__(self, project_root: str = "."):
        self.project_root = Path(project_root)
        self.results = {
            "timestamp": datetime.now().isoformat(),
            "total_issues": 0,
            "issues": [],
            "fixes_applied": [],
            "health_score": 100
        }
        
    def check_root_junk_files(self) -> List[Dict]:
        """检查根目录垃圾文件"""
        issues = []
        root_files = set(self.project_root.iterdir())
        
        # 常见的临时文件和测试文件模式
        junk_patterns = [
            "*.tmp", "*.temp", "*.log", "*.bak",
            "test_*.py", "*_test.py", "test_*.sh",
            "__pycache__", "*.pyc", ".pytest_cache",
            "*.swp", "*.swo", ".DS_Store"
        ]
        
        junk_files = []
        for pattern in junk_patterns:
            junk_files.extend(glob.glob(str(self.project_root / pattern)))
        
        # 检查是否应该在tests/目录内
        tests_dir = self.project_root / "tests"
        for file_path in junk_files:
            file_path = Path(file_path)
            
            # 如果是测试文件且不在tests/目录内，标记为问题
            if (file_path.name.startswith("test_") or "_test.py" in file_path.name) and not tests_dir.is_relative_to(file_path):
                issues.append({
                    "type": "root_junk_files",
                    "file": str(file_path),
                    "issue": f"测试文件 {file_path.name} 散落在根目录，应在 tests/ 内",
                    "severity": "medium"
                })
            
            # 临时文件
            elif any(file_path.name.endswith(ext) for ext in [".tmp", ".temp", ".log", ".bak"]):
                issues.append({
                    "type": "root_junk_files", 
                    "file": str(file_path),
                    "issue": f"临时文件 {file_path.name} 散落在根目录",
                    "severity": "low"
                })
        
        return issues
    
    def check_test_redundancy(self) -> List[Dict]:
        """检查测试冗余"""
        issues = []
        tests_dir = self.project_root / "tests"
        
        if not tests_dir.exists():
            return issues
            
        # 收集所有测试文件
        test_files = list(tests_dir.glob("*.py"))
        
        # 检查测试文件中的功能重复
        function_tests = {}
        
        for test_file in test_files:
            try:
                with open(test_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                    
                # 提取测试函数名
                test_functions = re.findall(r'def (test_\w+)', content)
                
                for func in test_functions:
                    if func not in function_tests:
                        function_tests[func] = []
                    function_tests[func].append(str(test_file))
                    
            except Exception:
                continue
        
        # 查找重复功能的测试
        for func, files in function_tests.items():
            if len(files) > 1:
                issues.append({
                    "type": "test_redundancy",
                    "function": func,
                    "files": files,
                    "issue": f"测试函数 {func} 在 {len(files)} 个文件中重复测试",
                    "severity": "medium"
                })
        
        return issues
    
    def check_round_number_tests(self) -> List[Dict]:
        """检查轮次编号测试文件"""
        issues = []
        tests_dir = self.project_root / "tests"
        
        if not tests_dir.exists():
            return issues
            
        # 匹配轮次编号模式：test_reqXXX_rXX.py
        round_pattern = re.compile(r'test_req\d+_r\d+\.py')
        
        round_files = []
        for test_file in tests_dir.glob("*.py"):
            if round_pattern.match(test_file.name):
                round_files.append(str(test_file))
        
        if round_files:
            issues.append({
                "type": "round_number_tests",
                "files": round_files,
                "issue": f"发现 {len(round_files)} 个带轮次编号的测试文件，应合并到功能文件后删除",
                "severity": "low"
            })
        
        return issues
    
    def check_document_bloat(self) -> List[Dict]:
        """检查文档膨胀"""
        issues = []
        
        # 检查DECISIONS.md
        decisions_path = self.project_root / "docs" / "DECISIONS.md"
        if decisions_path.exists():
            with open(decisions_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                
            if len(lines) > 40:
                issues.append({
                    "type": "document_bloat",
                    "document": "docs/DECISIONS.md",
                    "lines": len(lines),
                    "issue": f"DECISIONS.md 超过40行 ({len(lines)}行)，需要文档瘦身",
                    "severity": "high"
                })
        
        # 检查NOW.md
        now_path = self.project_root / "docs" / "NOW.md"
        if now_path.exists():
            with open(now_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                
            if len(lines) > 30:
                issues.append({
                    "type": "document_bloat",
                    "document": "docs/NOW.md", 
                    "lines": len(lines),
                    "issue": f"NOW.md 超过30行 ({len(lines)}行)，需要精简",
                    "severity": "medium"
                })
        
        return issues
    
    def check_abandoned_worktrees(self) -> List[Dict]:
        """检查废弃分支和worktree"""
        issues = []
        
        try:
            # 检查是否有feature分支
            result = os.popen("git branch | grep feature/").read().strip()
            if result:
                branches = result.split('\n')
                issues.append({
                    "type": "abandoned_worktrees",
                    "branches": branches,
                    "issue": f"发现 {len(branches)} 个未清理的feature分支",
                    "severity": "low"
                })
            
            # 检查是否有worktree
            result = os.popen("git worktree list").read().strip()
            if result and "wt-" in result:
                lines = result.split('\n')
                wt_paths = [line for line in lines if "wt-" in line]
                issues.append({
                    "type": "abandoned_worktrees",
                    "worktrees": wt_paths,
                    "issue": f"发现 {len(wt_paths)} 个未清理的worktree",
                    "severity": "low"
                })
                
        except Exception as e:
            issues.append({
                "type": "abandoned_worktrees",
                "error": str(e),
                "issue": "无法检查git分支状态",
                "severity": "low"
            })
        
        return issues
    
    def check_dependency_changes(self) -> List[Dict]:
        """检查依赖变化"""
        issues = []
        
        pyproject_path = self.project_root / "pyproject.toml"
        if not pyproject_path.exists():
            return issues
            
        with open(pyproject_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # 检查是否有新依赖（这里需要历史记录对比）
        # 简化版本：检查是否有不必要的依赖
        suspicious_deps = [
            "pandas", "numpy", "scipy",  # 数据科学库
            "django", "flask", "fastapi", # Web框架
            "opencv", "pillow",          # 图像处理
            "tensorflow", "pytorch",     # 机器学习
            "requests", "urllib3",       # HTTP客户端（除非必要）
        ]
        
        found_deps = []
        for dep in suspicious_deps:
            if dep.lower() in content.lower():
                found_deps.append(dep)
        
        if found_deps:
            issues.append({
                "type": "dependency_changes",
                "dependencies": found_deps,
                "issue": f"发现可能不必要的依赖: {', '.join(found_deps)}",
                "severity": "low"
            })
        
        return issues
    
    def apply_fixes(self, issues: List[Dict]) -> List[str]:
        """应用修复"""
        fixes = []
        
        for issue in issues:
            if issue["type"] == "root_junk_files":
                # 移动测试文件到tests/目录
                if issue["file"].endswith('.py') and ('test_' in issue["file"] or '_test.py' in issue["file"]):
                    tests_dir = self.project_root / "tests"
                    tests_dir.mkdir(exist_ok=True)
                    
                    src_file = Path(issue["file"])
                    dst_file = tests_dir / src_file.name
                    
                    try:
                        shutil.move(src_file, dst_file)
                        fixes.append(f"移动测试文件: {src_file.name} → tests/")
                    except Exception as e:
                        fixes.append(f"移动失败 {src_file.name}: {str(e)}")
                        
                # 删除临时文件
                elif any(issue["file"].endswith(ext) for ext in [".tmp", ".temp", ".log", ".bak"]):
                    try:
                        os.remove(issue["file"])
                        fixes.append(f"删除临时文件: {Path(issue['file']).name}")
                    except Exception as e:
                        fixes.append(f"删除失败 {issue['file']}: {str(e)}")
            
            elif issue["type"] == "round_number_tests":
                # 删除轮次编号测试文件
                for file_path in issue["files"]:
                    try:
                        os.remove(file_path)
                        fixes.append(f"删除轮次测试文件: {Path(file_path).name}")
                    except Exception as e:
                        fixes.append(f"删除失败 {Path(file_path).name}: {str(e)}")
            
            elif issue["type"] == "abandoned_worktrees":
                # 清理git分支（需要手动确认）
                fixes.append(f"需要手动清理: {len(issue.get('branches', []) + issue.get('worktrees', []))} 个分支/worktree")
        
        return fixes
    
    def generate_report(self) -> str:
        """生成健康度报告"""
        report = []
        report.append(f"项目健康度检查报告 - {self.results['timestamp']}")
        report.append("=" * 50)
        
        total_issues = len(self.results["issues"])
        self.results["total_issues"] = total_issues
        
        # 计算健康度分数
        health_score = 100
        for issue in self.results["issues"]:
            if issue["severity"] == "high":
                health_score -= 10
            elif issue["severity"] == "medium":
                health_score -= 5
            elif issue["severity"] == "low":
                health_score -= 2
        
        health_score = max(0, health_score)
        self.results["health_score"] = health_score
        
        report.append(f"📊 健康度评分: {health_score}/100")
        report.append(f"🔍 发现问题: {total_issues} 个")
        report.append(f"✅ 已修复: {len(self.results['fixes_applied'])} 个")
        
        if total_issues > 0:
            report.append("\n🚨 问题详情:")
            for i, issue in enumerate(self.results["issues"], 1):
                severity_icon = {"high": "🔴", "medium": "🟡", "low": "🟢"}[issue["severity"]]
                report.append(f"  {i}. {severity_icon} {issue['issue']}")
                
                if issue.get("file"):
                    report.append(f"     文件: {issue['file']}")
                elif issue.get("files"):
                    report.append(f"     文件: {', '.join(issue['files'][:3])}...")
        
        if self.results["fixes_applied"]:
            report.append("\n🔧 已应用修复:")
            for fix in self.results["fixes_applied"]:
                report.append(f"  ✓ {fix}")
        
        if health_score >= 90:
            report.append("\n🎉 项目健康状况优秀!")
        elif health_score >= 70:
            report.append("\n✅ 项目健康状况良好，建议优化小问题")
        elif health_score >= 50:
            report.append("\n⚠️ 项目需要关注，建议优先修复中等问题")
        else:
            report.append("\n🚨 项目健康状况较差，建议立即修复!")
        
        return "\n".join(report)
    
    def save_report(self):
        """保存检查报告"""
        report_dir = self.project_root / ".health_check_history"
        report_dir.mkdir(exist_ok=True)
        
        report_file = report_dir / f"health_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)


def main():
    """主函数"""
    print("🏥 开始项目健康度检查...")
    
    checker = ProjectHealthChecker()
    
    # 执行各项检查
    issues = []
    issues.extend(checker.check_root_junk_files())
    issues.extend(checker.check_test_redundancy())
    issues.extend(checker.check_round_number_tests())
    issues.extend(checker.check_document_bloat())
    issues.extend(checker.check_abandoned_worktrees())
    issues.extend(checker.check_dependency_changes())
    
    checker.results["issues"] = issues
    
    # 应用修复（安全级别：只处理低风险操作）
    fixable_issues = [issue for issue in issues if issue["severity"] in ["low", "medium"]]
    fixes = checker.apply_fixes(fixable_issues)
    checker.results["fixes_applied"] = fixes
    
    print("\n" + "=" * 60)
    print(checker.generate_report())
    
    # 保存报告
    checker.save_report()
    print("📊 健康度报告已保存到 .health_check_history/")
    
    # 记录到DECISIONS.md
    if issues:
        decision_content = f"""
### [项目健康度检查] 第{len(issues)}次检查
- **时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} UTC
- **健康度评分**: {checker.results['health_score']}/100
- **发现问题**: {len(issues)} 个
- **修复执行**: {len(fixes)} 个安全修复
- **未修复**: {len(issues) - len(fixes)} 个需人工处理（如Git分支清理）
- **效果**: 自动清理垃圾文件、移动测试文件、删除冗余文件，提升项目整洁度
- **依据**: RULES.md项目健康度自检规则（每20轮至少1次）
"""
        
        with open("docs/DECISIONS.md", 'a', encoding='utf-8') as f:
            f.write(decision_content)
            
        print(f"\n📝 已记录到 docs/DECISIONS.md")
    
    return 0 if len(issues) == 0 else 1


if __name__ == "__main__":
    exit(main())