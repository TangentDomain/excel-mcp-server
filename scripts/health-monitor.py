#!/usr/bin/env python3
"""
项目健康度持续监控脚本
监控代码质量、测试覆盖、文档状态等关键指标
"""
import os
import re
import json
import subprocess
from pathlib import Path
from typing import Dict, List, Tuple, Any
from datetime import datetime

class ProjectHealthMonitor:
    def __init__(self, project_dir: str = "."):
        self.project_dir = Path(project_dir)
        self.issues: List[str] = []
        self.improvements: List[str] = []
        self.metrics: Dict[str, Any] = {}
        
    def run_command(self, cmd: List[str], description: str) -> Tuple[bool, str]:
        """运行shell命令并返回结果"""
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, cwd=self.project_dir)
            return result.returncode == 0, result.stdout + result.stderr
        except Exception as e:
            return False, f"执行 {cmd} 失败: {e}"
    
    def check_root_directory_cleanliness(self) -> Dict[str, Any]:
        """检查根目录清洁度"""
        metrics = {
            "temp_files": [],
            "backup_files": [],
            "build_artifacts": [],
            "total_issues": 0
        }
        
        # 检查临时文件
        temp_patterns = ["*.tmp", "*.temp", "*.bak", "*.backup", "*~"]
        for pattern in temp_patterns:
            result, output = self.run_command(
                ["find", ".", "-maxdepth", "1", "-name", pattern], 
                f"查找临时文件 {pattern}"
            )
            if result and output.strip():
                files = [f"./{f.strip()}" for f in output.strip().split('\n') if f.strip()]
                metrics["temp_files"].extend(files)
        
        # 检查构建产物
        build_patterns = ["dist/", "build/", "*.egg-info/", "*.whl"]
        for pattern in build_patterns:
            if os.path.exists(pattern):
                metrics["build_artifacts"].append(pattern)
        
        metrics["total_issues"] = len(metrics["temp_files"]) + len(metrics["build_artifacts"])
        
        return metrics
    
    def check_test_redundancy(self) -> Dict[str, Any]:
        """检查测试冗余"""
        metrics = {
            "total_test_files": 0,
            "duplicate_tests": [],
            "large_test_files": [],
            "merged_suggestions": []
        }
        
        # 查找所有测试文件
        result, output = self.run_command(
            ["find", "tests/", "-name", "test_*.py"], 
            "查找测试文件"
        )
        if result and output.strip():
            test_files = [f.strip() for f in output.strip().split('\n') if f.strip()]
            metrics["total_test_files"] = len(test_files)
            
            # 检查大文件（>1000行）
            for test_file in test_files:
                file_path = self.project_dir / test_file
                if file_path.exists():
                    lines = len(file_path.read_text().splitlines())
                    if lines > 1000:
                        metrics["large_test_files"].append({"file": test_file, "lines": lines})
            
            # 检查功能重复的测试文件
            test_contents = {}
            for test_file in test_files:
                file_path = self.project_dir / test_file
                if file_path.exists():
                    content = file_path.read_text()[:1000]  # 只看前1000字符
                    test_contents.setdefault(content, []).append(test_file)
            
            for content, files in test_contents.items():
                if len(files) > 1:
                    metrics["duplicate_tests"].append(files)
                    metrics["merged_suggestions"].append(f"考虑合并: {files}")
        
        return metrics
    
    def check_round_numbered_tests(self) -> Dict[str, Any]:
        """检查轮次编号测试文件"""
        metrics = {
            "round_numbered_files": [],
            "cleanup_suggestions": []
        }
        
        result, output = self.run_command(
            ["find", "tests/", "-name", "*test*req*.py"], 
            "查找REQ编号测试文件"
        )
        if result and output.strip():
            req_files = [f.strip() for f in output.strip().split('\n') if f.strip()]
            metrics["round_numbered_files"] = req_files
            
            # 检查是否可以合并
            for req_file in req_files:
                file_path = self.project_dir / req_file
                if file_path.exists():
                    content = file_path.read_text()
                    # 如果文件很小且内容简单，建议删除
                    if len(content) < 2000 and "def test_" in content.count("def test_") == 1:
                        metrics["cleanup_suggestions"].append(f"删除小文件: {req_file}")
        
        return metrics
    
    def check_branch_cleanup(self) -> Dict[str, Any]:
        """检查分支清理"""
        metrics = {
            "active_branches": [],
            "stale_features": [],
            "cleanup_suggestions": []
        }
        
        # 检查活跃分支
        result, output = self.run_command(
            ["git", "branch", "--list", "feature/*"], 
            "查找feature分支"
        )
        if result:
            branches = [b.strip() for b in output.split('\n') if b.strip()]
            metrics["active_branches"] = branches
            
            # 检查是否有过期分支（超过30天）
            for branch in branches:
                result, output = self.run_command(
                    ["git", "log", "--format=%ci", "-1", branch], 
                    f"检查分支 {branch} 最后提交时间"
                )
                if result and output.strip():
                    commit_date = datetime.strptime(output.strip(), "%Y-%m-%d %H:%M:%S %z")
                    days_old = (datetime.now(commit_date.tzinfo) - commit_date).days
                    if days_old > 30:
                        metrics["stale_features"].append(branch)
                        metrics["cleanup_suggestions"].append(f"清理过期分支: {branch} ({days_old}天)")
        
        return metrics
    
    def check_dependency_safety(self) -> Dict[str, Any]:
        """检查依赖安全性"""
        metrics = {
            "dependencies": [],
            "new_dependencies": [],
            "security_concerns": []
        }
        
        # 读取pyproject.toml
        pyproject_path = self.project_dir / "pyproject.toml"
        if pyproject_path.exists():
            content = pyproject_path.read_text()
            
            # 检查dependencies
            dep_pattern = r'dependencies\s*=\s*\[(.*?)\]'
            match = re.search(dep_pattern, content, re.DOTALL)
            if match:
                deps_text = match.group(1)
                deps = re.findall(r'"([^"]+)"', deps_text)
                metrics["dependencies"] = deps
                
                # 检查是否有新依赖（简单检查）
                base_deps = ["openpyxl", "python-calamine", "mcp", "xlcalculator", "sqlglot"]
                for dep in deps:
                    if dep not in base_deps:
                        metrics["new_dependencies"].append(dep)
                
                # 检查潜在的安全问题
                for dep in deps:
                    if "requests" in dep or "urllib" in dep:
                        metrics["security_concerns"].append(f"网络相关依赖: {dep}")
        
        return metrics
    
    def generate_health_report(self) -> Dict[str, Any]:
        """生成健康度报告"""
        report = {
            "timestamp": datetime.now().isoformat(),
            "overall_score": 100,
            "categories": {},
            "recommendations": [],
            "auto_fixes": []
        }
        
        # 执行各项检查
        report["categories"]["root_cleanliness"] = self.check_root_directory_cleanliness()
        report["categories"]["test_redundancy"] = self.check_test_redundancy()
        report["categories"]["round_numbered_tests"] = self.check_round_numbered_tests()
        report["categories"]["branch_cleanup"] = self.check_branch_cleanup()
        report["categories"]["dependency_safety"] = self.check_dependency_safety()
        
        # 计算总分
        total_issues = 0
        for category in report["categories"].values():
            if isinstance(category, dict) and "total_issues" in category:
                total_issues += category["total_issues"]
        
        report["overall_score"] = max(0, 100 - total_issues * 5)
        
        # 生成建议
        for category_name, category in report["categories"].items():
            if category.get("cleanup_suggestions"):
                report["recommendations"].extend(category["cleanup_suggestions"])
            
            if category.get("merged_suggestions"):
                report["recommendations"].extend(category["merged_suggestions"])
        
        # 自动修复
        report["auto_fixes"] = self.auto_fix_issues(report)
        
        return report
    
    def auto_fix_issues(self, report: Dict[str, Any]) -> List[str]:
        """自动修复可解决的问题"""
        fixes = []
        
        # 清理临时文件
        root_clean = report["categories"]["root_cleanliness"]
        for temp_file in root_clean["temp_files"]:
            try:
                Path(temp_file).unlink()
                fixes.append(f"删除临时文件: {temp_file}")
            except:
                pass
        
        # 清理构建产物
        for build_artifact in root_clean["build_artifacts"]:
            if os.path.exists(build_artifact):
                try:
                    import shutil
                    if os.path.isdir(build_artifact):
                        shutil.rmtree(build_artifact)
                    else:
                        os.unlink(build_artifact)
                    fixes.append(f"清理构建产物: {build_artifact}")
                except:
                    pass
        
        return fixes
    
    def save_report(self, report: Dict[str, Any], filename: str = "health-monitor-result.json"):
        """保存健康度报告"""
        report_path = self.project_dir / filename
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        
        return report_path

def main():
    """主函数"""
    monitor = ProjectHealthMonitor()
    
    print("🔍 执行项目健康度监控...")
    report = monitor.generate_health_report()
    
    # 保存报告
    report_path = monitor.save_report(report)
    print(f"📊 健康度报告已保存: {report_path}")
    
    # 输出摘要
    print(f"\n📈 健康度评分: {report['overall_score']}/100")
    print(f"🔧 自动修复项目: {len(report['auto_fixes'])}")
    print(f"💡 优化建议: {len(report['recommendations'])}")
    
    if report['auto_fixes']:
        print("\n✅ 已自动修复:")
        for fix in report['auto_fixes']:
            print(f"  - {fix}")
    
    if report['recommendations']:
        print("\n💡 优化建议:")
        for rec in report['recommendations'][:10]:  # 只显示前10条
            print(f"  - {rec}")
    
    return 0

if __name__ == "__main__":
    exit(main())