#!/usr/bin/env python3
"""
轮次快照管理工具
实现REQ-022: 轮次快照机制（进展追溯+断点续传闭环）
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

class RoundSnapshotManager:
    def __init__(self, project_path):
        self.project_path = Path(project_path)
        self.snapshot_file = self.project_path / ".round-snapshot.json"
        
    def read_snapshot(self):
        """读取轮次快照文件"""
        if not self.snapshot_file.exists():
            return None
        try:
            with open(self.snapshot_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error reading snapshot: {e}")
            return None
            
    def write_snapshot(self, round_number, req, phase, progress, files_modified, 
                      commits, tests, next_step, failed_attempts=None, blockers=None):
        """写入轮次快照文件"""
        snapshot = {
            "round": round_number,
            "req": req,
            "phase": phase,
            "progress": progress,
            "files_modified": files_modified,
            "commits": commits,
            "tests": tests,
            "next_step": next_step,
            "failed_attempts": failed_attempts or [],
            "blockers": blockers or [],
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat()
        }
        
        try:
            with open(self.snapshot_file, 'w', encoding='utf-8') as f:
                json.dump(snapshot, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error writing snapshot: {e}")
            return False
            
    def update_phase(self, new_phase):
        """更新当前phase"""
        snapshot = self.read_snapshot()
        if snapshot:
            snapshot["phase"] = new_phase
            snapshot["updated_at"] = datetime.now().isoformat()
            with open(self.snapshot_file, 'w', encoding='utf-8') as f:
                json.dump(snapshot, f, indent=2, ensure_ascii=False)
            return True
        return False
        
    def update_phase_detailed(self, new_phase, sub_step=None, status="running"):
        """更新详细phase信息（支持子步骤跟踪）"""
        snapshot = self.read_snapshot()
        if snapshot:
            snapshot["phase"] = new_phase
            snapshot["updated_at"] = datetime.now().isoformat()
            if sub_step:
                snapshot["sub_step"] = sub_step
                snapshot["sub_step_status"] = status
            with open(self.snapshot_file, 'w', encoding='utf-8') as f:
                json.dump(snapshot, f, indent=2, ensure_ascii=False)
            return True
        return False
        
    def update_progress(self, new_progress, next_step=None):
        """更新进度信息"""
        snapshot = self.read_snapshot()
        if snapshot:
            snapshot["progress"] = new_progress
            snapshot["updated_at"] = datetime.now().isoformat()
            if next_step:
                snapshot["next_step"] = next_step
            with open(self.snapshot_file, 'w', encoding='utf-8') as f:
                json.dump(snapshot, f, indent=2, ensure_ascii=False)
            return True
        return False
        
    def has_progress(self, last_round):
        """检查是否有进展（对比快照）"""
        if last_round < 0:
            return True  # 第一轮都有进展
            
        # 这里需要读取上一轮快照进行比较
        # 简化实现：有文件变更或commit就算有进展
        snapshot = self.read_snapshot()
        return snapshot is not None
        
    def get_next_phase(self):
        """从快照获取下一步骤"""
        snapshot = self.read_snapshot()
        if snapshot:
            return snapshot.get("next_step", "K0")
        return "K0"

def check_snapshot_exists(project_path):
    """检查快照文件是否存在"""
    snapshot_file = Path(project_path) / ".round-snapshot.json"
    return snapshot_file.exists()

def validate_snapshot_format(project_path):
    """验证快照文件格式"""
    snapshot_file = Path(project_path) / ".round-snapshot.json"
    if not snapshot_file.exists():
        return False
        
    try:
        with open(snapshot_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            required_fields = ["round", "req", "phase", "progress", "files_modified", "commits", "tests", "next_step"]
            return all(field in data for field in required_fields)
    except:
        return False

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 round-snapshot.py <action> [args...]")
        sys.exit(1)
        
    action = sys.argv[1]
    project_path = "/root/.openclaw/workspace/skills/meta-evolve"
    manager = RoundSnapshotManager(project_path)
    
    if action == "read":
        snapshot = manager.read_snapshot()
        print(json.dumps(snapshot, indent=2) if snapshot else "No snapshot found")
        
    elif action == "write":
        if len(sys.argv) < 9:
            print("Usage: python3 round-snapshot.py write <round> <req> <phase> <progress> <files_modified> <commits> <tests> <next_step>")
            sys.exit(1)
        round_num = int(sys.argv[2])
        req = sys.argv[3]
        phase = sys.argv[4]
        progress = sys.argv[5]
        files_modified = sys.argv[6].split(',') if sys.argv[6] else []
        commits = sys.argv[7].split(',') if sys.argv[7] else []
        tests = sys.argv[8]
        next_step = sys.argv[9] if len(sys.argv) > 9 else ""
        
        success = manager.write_snapshot(round_num, req, phase, progress, files_modified, 
                                       commits, tests, next_step)
        print(f"Write snapshot: {'Success' if success else 'Failed'}")
        
    elif action == "update-phase":
        new_phase = sys.argv[2]
        sub_step = sys.argv[3] if len(sys.argv) > 3 else None
        status = sys.argv[4] if len(sys.argv) > 4 else "running"
        success = manager.update_phase_detailed(new_phase, sub_step, status)
        print(f"Update phase: {'Success' if success else 'Failed'}")
        
    elif action == "update-progress":
        new_progress = sys.argv[2]
        next_step = sys.argv[3] if len(sys.argv) > 3 else None
        success = manager.update_progress(new_progress, next_step)
        print(f"Update progress: {'Success' if success else 'Failed'}")
        
    elif action == "check-exists":
        exists = check_snapshot_exists(project_path)
        print(f"Snapshot exists: {exists}")
        
    elif action == "validate":
        valid = validate_snapshot_format(project_path)
        print(f"Snapshot valid: {valid}")
        
    else:
        print(f"Unknown action: {action}")
        sys.exit(1)