#!/usr/bin/env python3
"""
元迭代进度管理工具
支持细粒度进度跟踪和断点续传

用法:
  python3 scripts/progress.py --help
  python3 scripts/progress.py check-step K2
  python3 scripts/progress.py mark-step K2 analyze_source done
  python3 scripts/progress.py get-progress K2
  python3 scripts/progress.py archive-round 21
"""

import json
import os
import sys
import argparse
from datetime import datetime
from pathlib import Path

def get_project_root():
    """获取项目根目录"""
    return Path(__file__).parent.parent

def get_progress_file(step):
    """获取进度文件路径"""
    project_root = get_project_root()
    return project_root / f".progress-{step}.json"

def get_archive_dir(round_num):
    """获取归档目录"""
    project_root = get_project_root()
    return project_root / "docs" / "progress" / f"round-{round_num}"

def init_progress_file(step, round_num):
    """初始化进度文件"""
    progress_file = get_progress_file(step)
    
    # 定义各步骤的子步骤
    sub_steps_config = {
        "K0": [
            {"name": "clear_markers", "description": "清空关键步骤标记"},
            {"name": "read_now", "description": "读docs/NOW.md获取轮次信息"},
            {"name": "collect_instances", "description": "采集其他实例状态"},
            {"name": "run_auto_detect", "description": "运行自动检测脚本"},
            {"name": "identify_improvements", "description": "识别可改进点"},
            {"name": "check_stuck_instances", "description": "检查卡住实例并行动"}
        ],
        "K1": [
            {"name": "read_requirements", "description": "读REQUIREMENTS.md检查OPEN需求"},
            {"name": "extract_lessons", "description": "运行经验库采集脚本"},
            {"name": "query_lessons", "description": "查询相关历史经验"},
            {"name": "health_check", "description": "运行健康检查"},
            {"name": "mark_busy", "description": "标记状态为BUSY/IDLE"},
            {"name": "write_k1_marker", "description": "写入K1标记"}
        ],
        "K2": [
            {"name": "select_requirement", "description": "选择最高优先级需求"},
            {"name": "analyze_source", "description": "分析源码和模板"},
            {"name": "implement_changes", "description": "实施改进"},
            {"name": "test_changes", "description": "测试验证"},
            {"name": "git_commit", "description": "Git提交"}
        ],
        "K3": [
            {"name": "update_now_md", "description": "更新docs/NOW.md"},
            {"name": "run_check_py", "description": "运行check.py验证"},
            {"name": "propagate_changes", "description": "架构文档同步和传播"},
            {"name": "write_k3_marker", "description": "写入K3标记"},
            {"name": "push_reports", "description": "推送报告到飞书和QQ"}
        ]
    }
    
    progress_data = {
        "round": round_num,
        "step": step,
        "started_at": datetime.now().isoformat(),
        "sub_steps": sub_steps_config.get(step, []),
        "context": {},
        "outputs": {}
    }
    
    with open(progress_file, 'w', encoding='utf-8') as f:
        json.dump(progress_data, f, indent=2, ensure_ascii=False)
    
    print(f"✅ 初始化进度文件: {progress_file}")
    return progress_file

def load_progress(step):
    """加载进度文件"""
    progress_file = get_progress_file(step)
    if not progress_file.exists():
        return None
    
    with open(progress_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_progress(progress_data):
    """保存进度数据"""
    progress_file = get_project_root() / f".progress-{progress_data['step']}.json"
    with open(progress_file, 'w', encoding='utf-8') as f:
        json.dump(progress_data, f, indent=2, ensure_ascii=False)

def mark_substep_done(step, substep_name, output_data=None):
    """标记子步骤完成"""
    progress_data = load_progress(step)
    if not progress_data:
        print(f"❌ 进度文件不存在: {step}")
        return False
    
    # 查找并更新子步骤
    for substep in progress_data["sub_steps"]:
        if substep["name"] == substep_name:
            substep["status"] = "done"
            substep["completed_at"] = datetime.now().isoformat()
            if output_data:
                substep["output"] = output_data
            break
    else:
        print(f"❌ 子步骤不存在: {substep_name}")
        return False
    
    # 更新上下文
    if "round" not in progress_data["context"]:
        progress_data["context"]["round"] = progress_data["round"]
    
    save_progress(progress_data)
    print(f"✅ 子步骤完成: {step}.{substep_name}")
    return True

def get_next_substep(step):
    """获取下一个待执行的子步骤"""
    progress_data = load_progress(step)
    if not progress_data:
        # 进度文件不存在，返回第一个子步骤
        project_root = get_project_root()
        init_progress_file(step, get_current_round())
        return progress_data["sub_steps"][0] if progress_data else None
    
    # 找到第一个未完成的子步骤
    for substep in progress_data["sub_steps"]:
        if substep.get("status") != "done":
            return substep
    
    return None

def get_current_round():
    """获取当前轮次"""
    project_root = get_project_root()
    now_file = project_root / "docs" / "NOW.md"
    
    if now_file.exists():
        with open(now_file, 'r', encoding='utf-8') as f:
            content = f.read()
            for line in content.split('\n'):
                if '第' in line and '轮' in line:
                    # 提取轮次数字
                    import re
                    match = re.search(r'第(\d+)轮', line)
                    if match:
                        return int(match.group(1))
    
    return 21  # 默认轮次

def archive_progress(round_num):
    """归档当前轮次的进度文件"""
    project_root = get_project_root()
    archive_dir = get_archive_dir(round_num)
    
    # 创建归档目录
    archive_dir.mkdir(parents=True, exist_ok=True)
    
    # 移动进度文件
    for step in ["K0", "K1", "K2", "K3"]:
        progress_file = get_progress_file(step)
        if progress_file.exists():
            archive_file = archive_dir / f"progress-{step}.json"
            progress_file.rename(archive_file)
            print(f"📁 归档进度文件: {progress_file} -> {archive_file}")
    
    print(f"✅ 轮次 {round_num} 进度已归档到: {archive_dir}")

def check_step_exists(step):
    """检查步骤是否存在进度文件"""
    progress_file = get_progress_file(step)
    return progress_file.exists()

def main():
    parser = argparse.ArgumentParser(description="元迭代进度管理工具")
    subparsers = parser.add_subparsers(dest='command', help='可用命令')
    
    # 检查步骤
    check_parser = subparsers.add_parser('check-step', help='检查步骤状态')
    check_parser.add_argument('step', choices=['K0', 'K1', 'K2', 'K3'], help='步骤名称')
    
    # 标记子步骤完成
    mark_parser = subparsers.add_parser('mark-step', help='标记子步骤完成')
    mark_parser.add_argument('step', choices=['K0', 'K1', 'K2', 'K3'], help='步骤名称')
    mark_parser.add_argument('substep', help='子步骤名称')
    mark_parser.add_argument('--output', help='输出数据（JSON格式）')
    
    # 获取进度
    get_parser = subparsers.add_parser('get-progress', help='获取进度信息')
    get_parser.add_argument('step', choices=['K0', 'K1', 'K2', 'K3'], help='步骤名称')
    
    # 归档轮次
    archive_parser = subparsers.add_parser('archive-round', help='归档轮次进度')
    archive_parser.add_argument('round', type=int, help='轮次号')
    
    # 初始化进度文件
    init_parser = subparsers.add_parser('init', help='初始化指定步骤的进度文件')
    init_parser.add_argument('step', choices=['K0', 'K1', 'K2', 'K3'], help='步骤名称')
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return
    
    if args.command == 'check-step':
        exists = check_step_exists(args.step)
        print(f"步骤 {args.step}: {'存在' if exists else '不存在'}")
        
        if exists:
            progress = load_progress(args.step)
            print(f"开始时间: {progress.get('started_at', 'N/A')}")
            completed = [s for s in progress.get('sub_steps', []) if s.get('status') == 'done']
            print(f"完成子步骤: {len(completed)}/{len(progress.get('sub_steps', []))}")
    
    elif args.command == 'mark-step':
        output_data = None
        if args.output:
            try:
                output_data = json.loads(args.output)
            except json.JSONDecodeError:
                print("❌ 输出数据格式错误，应为JSON")
                return
        
        success = mark_substep_done(args.step, args.substep, output_data)
        if success:
            print(f"✅ {args.step}.{args.substep} 已标记为完成")
        else:
            print(f"❌ 标记失败: {args.step}.{args.substep}")
    
    elif args.command == 'get-progress':
        progress = load_progress(args.step)
        if progress:
            print(json.dumps(progress, indent=2, ensure_ascii=False))
        else:
            print(f"❌ 进度文件不存在: {args.step}")
    
    elif args.command == 'archive-round':
        archive_progress(args.round)
        print(f"✅ 轮次 {args.round} 已归档")
    
    elif args.command == 'init':
        round_num = get_current_round()
        init_progress_file(args.step, round_num)
        print(f"✅ 初始化进度文件 {args.step} for 轮次 {round_num}")

if __name__ == "__main__":
    main()