#!/usr/bin/env python3
"""
项目健康度监控脚本
每轮自动执行，监控项目质量指标
"""

import os
import re
import json
import subprocess
from pathlib import Path
from datetime import datetime

def run_command(cmd, capture_output=True):
    """运行shell命令并返回结果"""
    try:
        if capture_output:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=30)
            return result.returncode == 0, result.stdout.strip(), result.stderr.strip()
        else:
            subprocess.run(cmd, shell=True, timeout=30)
            return True, "", ""
    except subprocess.TimeoutExpired:
        return False, "", "Command timeout"
    except Exception as e:
        return False, "", str(e)

def check_root_directory_cleanliness():
    """检查根目录清洁度"""
    issues = []
    root_files = [f for f in os.listdir('.') if os.path.isfile(f)]
    
    # 检查临时脚本文件
    temp_patterns = ['tmp', 'temp', 'test_', 'debug', 'backup']
    for file in root_files:
        if any(pattern in file.lower() for pattern in temp_patterns) and not file.startswith('.'):
            if not file.endswith(('.py', '.md', '.txt', '.json', '.toml', '.lock')):
                issues.append(f"临时文件: {file}")
    
    # 检查备份文件
    backup_files = [f for f in root_files if f.endswith('.backup')]
    if backup_files:
        issues.append(f"备份文件: {len(backup_files)}个")
    
    return len(issues), issues

def test_redundancy_check():
    """检查测试冗余"""
    issues = []
    test_dir = Path('tests')
    
    if test_dir.exists():
        test_files = list(test_dir.glob('test_*.py'))
        
        # 检查轮次编号测试文件
        numbered_tests = [f for f in test_files if re.search(r'_r\d+\.py$', f.name)]
        if numbered_tests:
            issues.append(f"轮次编号测试文件: {len(numbered_tests)}个")
        
        # 简单检查可能的重复功能文件
        test_modules = {}
        for test_file in test_files:
            module_name = test_file.stem.replace('test_', '')
            if module_name in test_modules:
                issues.append(f"可能重复的测试模块: {module_name}")
            test_modules[module_name] = test_file
    
    return len(issues), issues

def dependency_safety_check():
    """检查依赖安全性"""
    issues = []
    
    # 检查pyproject.toml依赖
    success, stdout, stderr = run_command("grep -A 20 'dependencies = \\[' pyproject.toml")
    if success:
        dependencies = []
        lines = stdout.split('\n')
        for line in lines:
            line = line.strip().strip(',').strip('"')
            if line and not line.startswith('['):
                if line != 'dependencies':
                    dependencies.append(line)
        
        # 检查是否有不必要的生产依赖
        prod_deps = [d for d in dependencies if not d.startswith(('pytest-', 'coverage-', 'black-', 'flake8-'))]
        
        # 检查版本约束是否合理
        for dep in prod_deps:
            if '>=' in dep and any(ver in dep for ver in ['0.0', '0.1', '0.01']):
                issues.append(f"宽松版本约束: {dep}")
    
    return len(issues), issues

def document_consistency_check():
    """检查文档一致性"""
    issues = []
    
    # 检查版本文件一致性
    version_files = ['pyproject.toml', 'src/excel_mcp_server_fastmcp/__init__.py']
    versions = {}
    
    for file in version_files:
        if os.path.exists(file):
            success, stdout, stderr = run_command(f"grep -o 'version.*=.*\"[0-9.]*\"' {file}")
            if success:
                match = re.search(r'version.*=.*"([0-9.]+)"', stdout)
                if match:
                    versions[file] = match.group(1)
    
    # 检查版本是否一致
    if len(set(versions.values())) > 1:
        issues.append(f"版本不一致: {versions}")
    
    # 检查文档行数
    doc_files = ['docs/NOW.md', 'docs/DECISIONS.md', 'REQUIREMENTS.md']
    for doc_file in doc_files:
        if os.path.exists(doc_file):
            with open(doc_file, 'r', encoding='utf-8') as f:
                lines = len(f.readlines())
                if doc_file == 'docs/DECISIONS.md' and lines > 40:
                    issues.append(f"DECISIONS.md超限: {lines}行")
                elif doc_file == 'docs/NOW.md' and lines > 30:
                    issues.append(f"NOW.md超限: {lines}行")
    
    return len(issues), issues

def check_branches_and_worktrees():
    """检查分支和worktree状态"""
    issues = []
    
    # 检查是否有过期的feature分支
    success, stdout, stderr = run_command("git branch | grep feature/")
    if success and stdout.strip():
        feature_branches = stdout.strip().split('\n')
        issues.append(f"未清理的feature分支: {len(feature_branches)}个")
    
    # 检查worktree
    success, stdout, stderr = run_command("git worktree list")
    if success and len(stdout.split('\n')) > 2:  # 应该只有1个main分支
        worktree_lines = [line for line in stdout.split('\n') if 'worktree' in line]
        issues.append(f"额外worktree: {len(worktree_lines)}个")
    
    return len(issues), issues

def main():
    """主函数"""
    print("🔍 开始项目健康度监控...")
    
    results = {
        'timestamp': datetime.now().isoformat(),
        'metrics': {}
    }
    
    # 执行各项检查
    metrics = [
        ('root_cleanliness', check_root_directory_cleanliness),
        ('test_redundancy', test_redundancy_check),
        ('dependency_safety', dependency_safety_check),
        ('document_consistency', document_consistency_check),
        ('branches_health', check_branches_and_worktrees),
    ]
    
    total_issues = 0
    all_issues = []
    
    for name, check_func in metrics:
        issue_count, issues = check_func()
        results['metrics'][name] = {
            'issues_count': issue_count,
            'issues': issues
        }
        total_issues += issue_count
        all_issues.extend(issues)
    
    # 输出结果
    print(f"📊 健康度检查完成")
    print(f"⚠️ 发现问题总数: {total_issues}")
    
    if total_issues > 0:
        print("\n🔍 问题详情:")
        for i, issue in enumerate(all_issues[:10], 1):  # 最多显示10个问题
            print(f"  {i}. {issue}")
        if len(all_issues) > 10:
            print(f"  ... 还有 {len(all_issues) - 10} 个问题")
    else:
        print("✅ 项目健康度良好，未发现问题")
    
    # 保存结果到文件
    with open('health-monitor-result.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    
    # 返回问题总数作为退出码
    return total_issues

if __name__ == "__main__":
    exit_code = main()
    exit(exit_code)