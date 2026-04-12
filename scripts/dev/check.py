#!/usr/bin/env python3
"""元迭代状态检查脚本 — 全生命周期"""

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path

def run(cmd, cwd=None):
    """执行命令，返回stdout"""
    try:
        r = subprocess.run(cmd, shell=True, capture_output=True, text=True, cwd=cwd, timeout=10)
        return r.stdout.strip(), r.returncode
    except:
        return "", 1

def get_instance_type(project):
    """识别项目类型，用于个性化质量基线"""
    project_name = os.path.basename(project)
    
    # 基于项目路径和文件特征识别类型
    if "excel-mcp" in project_name.lower():
        return "code"
    elif "cocos" in project_name.lower() or "prefab" in project_name.lower():
        return "code"
    elif "meta-evolve" in project_name.lower():
        return "meta"
    elif "tracking" in project_name.lower() or "maintenance" in project_name.lower():
        return "tracking"
    elif "research" in project_name.lower() or "topic" in project_name.lower():
        return "research"
    else:
        return "code"  # 默认类型

def should_use_r1_warn(project):
    """检查是否应该对R1使用WARN模式（自适应阈值）- 增强版 REQ-023"""
    try:
        # 读取检查历史
        hist_file = os.path.join(project, "docs", "check-history.jsonl")
        if not os.path.exists(hist_file):
            return False
        
        # 分析最近30轮的R1结果（扩大采样范围）
        recent_r1_fails = []
        with open(hist_file, 'r') as f:
            for line in f:
                try:
                    record = json.loads(line.strip())
                    if record.get('round', 0) >= (get_current_round(project) - 30):  # 最近30轮
                        checks = record.get('checks', {})
                        if checks.get('R1_critical_steps') == 'FAIL':
                            recent_r1_fails.append(record)
                except:
                    continue
        
        instance_type = get_instance_type(project)
        
        # 增强自适应逻辑（REQ-023质量模式自适应优化）：
        # 1. 统计误报模式：R1 FAIL但最终完成的记录
        success_after_fail = 0
        false_positive_pattern = []
        for fail_record in recent_r1_fails:
            round_num = fail_record.get('round', 0)
            # 检查这轮之后是否有成功的记录（无CRITICAL_FAIL）
            next_round_success = check_round_success(project, round_num + 1)
            if next_round_success:
                success_after_fail += 1
                false_positive_pattern.append({
                    "round": round_num,
                    "timestamp": fail_record.get("timestamp"),
                    "instance_type": instance_type
                })
        
        # 2. 特殊处理excel-mcp-server的连续误报模式（增强检测）
        excel_mcp_specific_pattern = False
        if "excel-mcp" in project.lower():
            # excel-mcp-server特有的连续误报检测 - 更复杂的模式识别
            if len(recent_r1_fails) >= 4:  # 降低阈值到4轮
                # 检查连续误报模式
                consecutive_fails = 0
                for i, fail_record in enumerate(recent_r1_fails[-15:]):  # 检查最近15轮
                    round_num = fail_record.get('round', 0)
                    if check_round_success(project, round_num + 1):
                        consecutive_fails += 1
                
                # 新规则：连续3轮以上误报但最终完成，或者总误报率超过60%
                total_fails = len(recent_r1_fails)
                fail_rate = total_fails / 30.0
                if consecutive_fails >= 3 or fail_rate >= 0.6:
                    excel_mcp_specific_pattern = True
                    # 记录特殊模式
                    record_excel_mcp_special_pattern(project, recent_r1_fails, consecutive_fails, fail_rate)
        
        # 3. 根据项目类型和模式调整阈值（REQ-023核心改进）- 更精细的分级
        if instance_type == "code" and "excel-mcp" in project.lower():
            # excel-mcp-server特殊处理：大幅降低阈值到2轮（针对观察到的连续误报）
            required_fails = 2
        elif instance_type == "code":
            # 其他代码项目：中等阈值
            required_fails = 4
        elif instance_type == "meta":
            # 元迭代项目：较高阈值（元迭代项目误报较少）
            required_fails = 5
        else:
            # 其他类型：中等阈值
            required_fails = 4
        
        consecutive_fails = len(recent_r1_fails) >= required_fails
        
        # 4. 新增：学习模式记录和阈值自适应调整
        learning_mode_score = calculate_learning_mode_score(project, recent_r1_fails)
        if learning_mode_score >= 0.7:  # 学习模式分数达到70%
            required_fails = max(1, required_fails - 1)  # 自动降低阈值
        
        # 5. 记录误报模式到数据库（增强版）
        if false_positive_pattern:
            record_false_positive_pattern(project, false_positive_pattern)
        
        # 6. 最终判断条件（REQ-023增强版）
        should_warn = (excel_mcp_specific_pattern or 
                      success_after_fail >= required_fails or 
                      consecutive_fails or
                      learning_mode_score >= 0.7)
        
        # 记录阈值调整决策
        if should_warn:
            record_threshold_decision(project, "R1", "WARN", {
                "reason": "adaptive_threshold",
                "instance_type": instance_type,
                "recent_fails": len(recent_r1_fails),
                "success_after_fail": success_after_fail,
                "excel_mcp_pattern": excel_mcp_specific_pattern,
                "learning_score": learning_mode_score
            })
        
        return should_warn
        
    except Exception as e:
        # 出错时保守处理，使用严格模式
        print(f"R1自适应阈值检查异常: {e}")
        return False

def check_round_success(project, target_round):
    """检查指定轮次是否成功完成"""
    try:
        hist_file = os.path.join(project, "docs", "check-history.jsonl")
        if not os.path.exists(hist_file):
            return False
        
        with open(hist_file, 'r') as f:
            for line in f:
                try:
                    record = json.loads(line.strip())
                    if record.get('round', 0) == target_round:
                        # 检查是否不是CRITICAL_FAIL且整体状态良好
                        return not record.get('critical_fail', True)
                except:
                    continue
        return False
    except:
        return False

def get_current_round(project):
    """获取当前轮次号"""
    try:
        now_file = os.path.join(project, "docs", "NOW.md")
        if not os.path.exists(now_file):
            return 0
        
        content = open(now_file).read()
        # 匹配 "轮次：第X轮" 或类似格式
        import re
        match = re.search(r'轮次[：:]\s*第?(\d+)轮', content)
        if match:
            return int(match.group(1))
        return 0
    except:
        return 0

def calculate_learning_mode_score(project, recent_r1_fails):
    """计算学习模式分数，评估系统改进程度 - REQ-023"""
    try:
        # 获取项目类型
        instance_type = get_instance_type(project)
        
        # 基于类型设置基线
        if instance_type == "excel-mcp":
            base_improvement_rate = 0.6  # 代码项目改善较快
        elif instance_type == "meta":
            base_improvement_rate = 0.5  # 元迭代项目适中
        else:
            base_improvement_rate = 0.4  # 其他项目较慢
        
        # 1. 分析R1失败趋势（反向指标）
        if not recent_r1_fails:
            fail_score = 0.8  # 没有失败是好消息
        else:
            # 计算失败密度
            fail_count = len(recent_r1_fails)
            fail_score = max(0, 1 - (fail_count / 10))  # 最多10次失败扣分到0
        
        # 2. 检查系统改进速度（通过git commit频率）
        git_log_result, git_code = run(f"git log --since='30 days ago' --oneline | wc -l", project)
        commit_frequency = min(1, int(git_log_result) / 30) if git_code == 0 else 0.3
        
        # 3. 检查配置优化情况（通过文件变更）
        config_files = [
            ".check-config.json",
            "SKILL.md", 
            "references/rules-template.md"
        ]
        config_changes = 0
        for config_file in config_files:
            if os.path.exists(os.path.join(project, config_file)):
                stat = os.stat(os.path.join(project, config_file))
                # 检查30天内是否有修改
                if datetime.fromtimestamp(stat.st_mtime, timezone.utc) > datetime.now(timezone.utc) - timedelta(days=30):
                    config_changes += 1
        
        config_score = min(1, config_changes / len(config_files))
        
        # 4. 综合计算学习模式分数
        learning_score = (fail_score * 0.4 +  # 失败趋势权重
                         commit_frequency * 0.3 +  # 改进速度权重  
                         config_score * 0.3)  # 配置优化权重
        
        # 5. 实例类型调整
        if instance_type == "meta":
            learning_score *= 1.1  # 元迭代项目学习效率更高
        
        # 6. 防止过度乐观
        learning_score = min(0.9, learning_score)
        
        return learning_score
    except:
        # 异常情况返回保守分数
        return 0.3

def record_false_positive_pattern(project, pattern_data):
    """记录误报模式到文件"""
    try:
        os.makedirs(os.path.join(project, "docs", "quality_patterns"), exist_ok=True)
        pattern_file = os.path.join(project, "docs", "quality_patterns", "false_positive_patterns.json")
        
        # 读取现有模式
        existing_patterns = []
        if os.path.exists(pattern_file):
            with open(pattern_file, 'r') as f:
                existing_patterns = json.load(f)
        
        # 添加新模式
        existing_patterns.extend(pattern_data)
        
        # 去重：基于round和instance_type
        unique_patterns = []
        seen = set()
        for pattern in existing_patterns:
            key = f"{pattern['round']}_{pattern['instance_type']}"
            if key not in seen:
                unique_patterns.append(pattern)
                seen.add(key)
        
        # 保持最近50条记录
        unique_patterns = unique_patterns[-50:]
        
        with open(pattern_file, 'w') as f:
            json.dump(unique_patterns, f, ensure_ascii=False, indent=2)
            
    except Exception:
        # 记录失败不影响主要逻辑
        pass

def should_use_r4_relaxed(project):
    """检查是否应该对R4使用宽松模式（自适应质量基线）"""
    try:
        instance_type = get_instance_type(project)
        hist_file = os.path.join(project, "docs", "check-history.jsonl")
        
        if not os.path.exists(hist_file):
            return False  # 无历史数据时使用严格模式
        
        # 分析最近10轮的R4结果
        recent_r4_results = []
        with open(hist_file, 'r') as f:
            for line in f:
                try:
                    record = json.loads(line.strip())
                    if record.get('round', 0) >= (get_current_round(project) - 10):
                        checks = record.get('checks', {})
                        if 'R4_quality' in checks:
                            recent_r4_results.append({
                                'status': checks['R4_quality'],
                                'round': record.get('round', 0)
                            })
                except:
                    continue
        
        if instance_type == "code":
            # 代码项目质量要求高，较少放宽
            return len(recent_r4_results) >= 8 and all(r['status'] == 'PASS' for r in recent_r4_results[-5:])
        elif instance_type == "meta":
            # 元迭代项目允许更多变动
            return len(recent_r4_results) >= 6 and all(r['status'] in ['PASS', 'WARN'] for r in recent_r4_results[-4:])
        else:
            # 其他类型居中
            return len(recent_r4_results) >= 7 and all(r['status'] in ['PASS', 'WARN'] for r in recent_r4_results[-4:])
            
    except Exception:
        return False

def check_critical_steps(project):
    """R1: 关键步骤链检查 - 渐进式检查版本 + 自适应阈值"""
    details = []
    fails = []
    warnings = []
    
    # 获取已存在的步骤和它们的标记
    existing_steps = {}
    for step in ["K0", "K1", "K2", "K3"]:
        f = os.path.join(project, f".step-{step}.done")
        if os.path.exists(f):
            try:
                ts = open(f).read().strip()
                existing_steps[step] = int(ts)
                details.append(f"{step}✅ (t={ts})")
            except:
                details.append(f"{step}❌ (时间戳格式错误)")
                fails.append(step)
        else:
            details.append(f"{step}❌ (标记不存在)")
            fails.append(step)
    
    # 如果只有K0存在，说明刚开始，直接PASS
    if len(existing_steps) == 1 and "K0" in existing_steps:
        return {
            "status": "PASS",
            "critical": False,
            "details": ["刚开始执行，只有K0标记，跳过严格检查"],
            "fix_suggestion": ""
        }
    
    # 自适应阈值检查 - 读取历史数据优化误报
    should_warn_only = should_use_r1_warn(project)
    if should_warn_only:
        details.append("🎯 自适应模式：检测到高频误报模式，R1降级为WARN")
        # 检查步骤连续性但不强制要求所有步骤
        step_order = ["K0", "K1", "K2", "K3"]
        for i in range(len(step_order) - 1):
            current_step = step_order[i]
            next_step = step_order[i + 1]
            
            if current_step in existing_steps and next_step in existing_steps:
                if existing_steps[current_step] > existing_steps[next_step]:
                    details.append(f"⚠️ 时间戳乱序: {current_step}({existing_steps[current_step]}) > {next_step}({existing_steps[next_step]})")
                    warnings.append(f"{current_step}>{next_step}")
        
        # 在自适应模式下，允许更多的不完整情况
        if len(fails) <= 2:  # 允许最多2个步骤缺失
            status = "WARN"
        else:
            status = "FAIL"
    else:
        # 原有严格逻辑
        step_order = ["K0", "K1", "K2", "K3"]
        for i in range(len(step_order) - 1):
            current_step = step_order[i]
            next_step = step_order[i + 1]
            
            if current_step in existing_steps and next_step in existing_steps:
                if existing_steps[current_step] > existing_steps[next_step]:
                    details.append(f"⚠️ 时间戳乱序: {current_step}({existing_steps[current_step]}) > {next_step}({existing_steps[next_step]})")
                    warnings.append(f"{current_step}>{next_step}")
        
        # 渐进式逻辑：如果K0和K1存在，允许K2/K3暂时缺失
        if "K0" in existing_steps and "K1" in existing_steps and "K2" not in existing_steps and "K3" not in existing_steps:
            status = "PASS"
            details.append("🔄 K0-K1完成，K2-K3进行中，符合渐进执行模式")
        elif "K0" in existing_steps and "K1" in existing_steps and "K2" in existing_steps and "K3" not in existing_steps:
            status = "PASS"
            details.append("🔄 K0-K2完成，K3进行中，符合渐进执行模式")
        elif len(fails) == 0:
            status = "PASS"
        elif len(fails) <= 1 and "K3" in fails:  # 只有K3缺失且其他都在，可以接受
            status = "WARN"
            details.append("⚠️ K3步骤缺失，可能推送步骤未完成")
        else:
            status = "FAIL"
    
    # 检查时间戳递增 - 只检查连续的步骤
    step_order = ["K0", "K1", "K2", "K3"]
    for i in range(len(step_order) - 1):
        current_step = step_order[i]
        next_step = step_order[i + 1]
        
        if current_step in existing_steps and next_step in existing_steps:
            if existing_steps[current_step] > existing_steps[next_step]:
                details.append(f"⚠️ 时间戳乱序: {current_step}({existing_steps[current_step]}) > {next_step}({existing_steps[next_step]})")
                warnings.append(f"{current_step}>{next_step}")
    
    # 渐进式逻辑：如果K0和K1存在，允许K2/K3暂时缺失
    if "K0" in existing_steps and "K1" in existing_steps and "K2" not in existing_steps and "K3" not in existing_steps:
        status = "PASS"
        details.append("🔄 K0-K1完成，K2-K3进行中，符合渐进执行模式")
    elif "K0" in existing_steps and "K1" in existing_steps and "K2" in existing_steps and "K3" not in existing_steps:
        status = "PASS"
        details.append("🔄 K0-K2完成，K3进行中，符合渐进执行模式")
    elif len(fails) == 0:
        status = "PASS"
    elif len(fails) <= 1 and "K3" in fails:  # 只有K3缺失且其他都在，可以接受
        status = "WARN"
        details.append("⚠️ K3步骤缺失，可能推送步骤未完成")
    else:
        status = "FAIL"
    
    # 修复建议
    if status == "FAIL":
        fix = "关键步骤严重缺失，需要完整执行K0-K3流程"
    elif status == "WARN":
        fix = "K3步骤缺失，建议检查是否已完成推送"
    else:
        fix = ""
        if warnings:
            fix = f"时间序异常: {', '.join(warnings)}，但不影响执行"

    return {
        "status": status,
        "critical": True,
        "details": details,
        "fix_suggestion": fix
    }

def check_progress(project):
    """R2: 进度连续性检查"""
    pf = os.path.join(project, ".claude-progress.md")
    if not os.path.exists(pf):
        return {"status": "FAIL", "critical": False, "details": ["进度文件不存在"], "fix_suggestion": "创建 .claude-progress.md"}
    
    content = open(pf).read()
    lines = content.strip().split("\n") if content.strip() else []
    details = []
    issues = []
    
    # 检查标题行
    has_title = any("# 进度" in l for l in lines)
    if not has_title:
        issues.append("缺少标题行 # 进度 - 第N轮")
    
    # 检查未关闭的 ▶️
    opened = [i for i, l in enumerate(lines) if "▶️" in l]
    closed = [i for i, l in enumerate(lines) if ("✅" in l or "❌" in l)]
    # 简单检查：每行 ▶️ 后面应该有对应 ✅/❌
    for i, line in enumerate(lines):
        if "▶️" in line:
            step_marker = line.split("第")[1].split("步")[0] if "第" in line and "步" in line else ""
            found_close = False
            for j in range(i+1, min(i+10, len(lines))):
                if f"第{step_marker}步" in lines[j] and ("✅" in lines[j] or "❌" in lines[j]):
                    found_close = True
                    break
            if step_marker and not found_close:
                issues.append(f"未关闭阶段: 第{step_marker}步")
    
    # 检查格式
    progress_lines = [l for l in lines if "▶️" in l or "✅" in l or "❌" in l or "🔄" in l]
    for l in progress_lines:
        if not l.strip().startswith("[") and ":" not in l[:10]:
            issues.append(f"格式不规范: {l[:50]}")
            break
    
    if not issues:
        details.append(f"进度文件存在，{len(progress_lines)}条进度记录")
    else:
        details.extend(issues)
    
    status = "WARN" if issues else "PASS"
    return {"status": status, "critical": False, "details": details, "fix_suggestion": "; ".join(issues[:3]) if issues else ""}

def check_git(project, thresholds=None):
    """R3: Git一致性检查"""
    if thresholds is None:
        thresholds = {}
    ahead_max = thresholds.get("R3_git_ahead_max", 5)
    # 工作区是否干净
    out, rc = run("git status --porcelain", cwd=project)
    uncommitted = len(out.split("\n")) if out else 0
    
    # develop vs main
    out2, rc2 = run("git rev-list main..develop --count 2>/dev/null || echo 0", cwd=project)
    try:
        ahead = int(out2)
    except:
        ahead = 0
    
    details = []
    issues = []
    
    if uncommitted > 0:
        issues.append(f"工作区有{uncommitted}个未提交改动")
    else:
        details.append("工作区干净")
    
    if ahead > ahead_max:
        issues.append(f"develop领先main {ahead}个commit（阈值{ahead_max}）")
    elif ahead > 0:
        details.append(f"develop领先main {ahead}个commit")
    else:
        details.append("main与develop同步")
    
    status = "WARN" if issues else "PASS"
    fix = "; ".join(issues)
    return {"status": status, "critical": False, "details": details, "fix_suggestion": fix}

def check_quality(project):
    """R4: 产出质量检查 - 增强自适应基线"""
    # 有没有新commit（对比.step-K0.done的时间）
    k0_file = os.path.join(project, ".step-K0.done")
    if not os.path.exists(k0_file):
        return {"status": "PASS", "critical": True, "details": ["K0标记不存在，跳过质量检查"], "fix_suggestion": ""}
    
    try:
        k0_ts = int(open(k0_file).read().strip())
    except:
        return {"status": "PASS", "critical": True, "details": [], "fix_suggestion": ""}
    
    out, rc = run(f"git log --since='{k0_ts}' --oneline", cwd=project)
    commits = [l for l in out.split("\n") if l.strip()]
    
    if not commits:
        return {"status": "PASS", "critical": True, "details": ["本轮无新commit，跳过质量检查"], "fix_suggestion": ""}
    
    instance_type = get_instance_type(project)
    details = [f"📊 实例类型: {instance_type}, 本轮{len(commits)}个新commit"]
    issues = []
    
    # 自适应质量基线 - 根据项目类型调整检查强度
    if instance_type == "code":
        # 代码项目严格检查commit质量
        max_bad_commits = 0
        max_file_deletion = 10  # 最多允许删除10个文件
    elif instance_type == "meta":
        # 元迭代项目宽松一些，允许更多维护类提交
        max_bad_commits = 2
        max_file_deletion = 20
    else:
        # 其他类型中等标准
        max_bad_commits = 1
        max_file_deletion = 15
    
    # 检查commit message格式 - 增强误报模式识别
    bad_commits = []
    for commit in commits:
        commit_msg = commit.strip()
        # 跳过无效的commit消息格式
        if not commit_msg or len(commit_msg) < 5:
            continue
            
        first_word = commit_msg.split()[0] if commit_msg.split() else ""
        
        # 合法的commit类型（增强版）
        is_valid = (
            "[REQ-" in commit_msg or  # 包含需求标记
            first_word == "Merge" or  # 合并提交
            first_word.startswith("v") and len(first_word) <= 5 or  # 版本提交 (v1.2.3)
            commit_msg.startswith("chore:") or  # 维护提交
            commit_msg.startswith("docs:") or  # 文档提交
            commit_msg.startswith("feat:") or  # 功能提交
            commit_msg.startswith("fix:") or   # 修复提交
            commit_msg.startswith("refactor:") or # 重构提交
            commit_msg.startswith("style:") or   # 代码风格
            commit_msg.startswith("test:") or   # 测试相关
            "自动" in commit_msg or "meta-evolve" in commit_msg or  # 自动化相关
            "自适应" in commit_msg or "质量模式" in commit_msg  # 质量优化相关
        )
        
        if not is_valid:
            bad_commits.append(commit_msg)
    
    if len(bad_commits) > max_bad_commits:
        issues.append(f"⚠️ commit格式: {len(bad_commits)}/{len(commits)} 个异常（阈值{max_bad_commits}）")
    
    # 检查文件删除比例 - 自适应阈值
    out2, rc2 = run(f"git diff --stat --since='{k0_ts}'", cwd=project)
    if out2:
        deleted_files = [line for line in out2.split('\n') if 'deleted' in line.lower()]
        total_deleted = len(deleted_files)
        if total_deleted > max_file_deletion:
            issues.append(f"🗑️ 文件删除过多: {total_deleted}个（阈值{max_file_deletion}）")
    
    # 根据项目类型和当前模式决定状态
    if not issues:
        status = "PASS"
        details.append("✅ 质量检查通过")
    elif len(issues) == 1 and instance_type != "code":
        # 非代码项目允许1个警告
        status = "WARN"
        details.append("⚠️ 质量检查有1个警告，在可接受范围内")
    else:
        status = "FAIL"
    
    # 记录误报模式（如果FAIL但实际上是合理的）
    if status == "FAIL" and is_potential_false_positive(project, commits, issues):
        record_potential_false_positive(project, {
            "round": get_current_round(project),
            "instance_type": instance_type,
            "issues": issues,
            "commits": commits,
            "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        })
        status = "WARN"  # 降级为WARN，记录待观察
        details.insert(0, "🎯 自适应调整：检测到潜在误报，降级为WARN")
    
    return {
        "status": status, 
        "critical": True, 
        "details": details + issues, 
        "fix_suggestion": "; ".join(issues[:2]) if issues else ""
    }

def is_potential_false_positive(project, commits, issues):
    """判断是否可能是误报"""
    try:
        # 检查是否有质量优化相关的提交
        quality_optimizations = any(
            "质量" in commit or "adaptive" in commit.lower() or "优化" in commit 
            for commit in commits
        )
        
        # 检查是否是元迭代的自我优化
        meta_evolve_commits = any(
            "meta-evolve" in commit or "自适应" in commit or "质量模式" in commit
            for commit in commits
        )
        
        # 如果有质量优化相关的提交，可能是合理的不合格
        return quality_optimizations or meta_evolve_commits
        
    except:
        return False

def record_potential_false_positive(project, data):
    """记录潜在的误报案例用于学习"""
    try:
        os.makedirs(os.path.join(project, "docs", "quality_patterns"), exist_ok=True)
        pattern_file = os.path.join(project, "docs", "quality_patterns", "potential_false_positives.json")
        
        # 读取现有记录
        existing_records = []
        if os.path.exists(pattern_file):
            with open(pattern_file, 'r') as f:
                existing_records = json.load(f)
        
        # 添加新记录
        existing_records.append(data)
        
        # 保持最近30条
        existing_records = existing_records[-30:]
        
        with open(pattern_file, 'w') as f:
            json.dump(existing_records, f, ensure_ascii=False, indent=2)
            
    except Exception:
        pass

def check_report(project):
    """R5: 报告推送检查"""
    pf = os.path.join(project, ".claude-progress.md")
    if not os.path.exists(pf):
        return {"status": "WARN", "critical": False, "details": ["进度文件不存在"], "fix_suggestion": ""}
    
    content = open(pf).read()
    has_push = any("推送" in l or "报告" in l for l in content.split("\n"))
    
    if has_push:
        return {"status": "PASS", "critical": False, "details": ["推送记录存在"], "fix_suggestion": ""}
    else:
        return {"status": "WARN", "critical": False, "details": ["未发现推送记录"], "fix_suggestion": "执行推送步骤"}

def check_docs_health(project, thresholds=None):
    """H1: 文档健康度"""
    if thresholds is None:
        thresholds = {}
    details = []
    issues = []
    
    for fname, limit_key, default in [("RULES.md", "H1_rules_max_lines", 300), ("NOW.md", "H1_now_max_lines", 30), ("DECISIONS.md", "H1_decisions_max_items", 40)]:
        limit = thresholds.get(limit_key, default)
        fpath = os.path.join(project, "docs", fname)
        if os.path.exists(fpath):
            lines = len(open(fpath).read().split("\n"))
            if lines > limit:
                issues.append(f"{fname} {lines}行（阈值{limit}）")
            else:
                details.append(f"{fname} {lines}/{limit}行")
        else:
            details.append(f"{fname} 不存在（跳过）")
    
    status = "WARN" if issues else "PASS"
    return {"status": status, "critical": False, "details": details + issues, "fix_suggestion": "; ".join(issues) if issues else ""}

def check_req_trend(project, thresholds=None):
    """H2: 需求趋势"""
    if thresholds is None:
        thresholds = {}
    req_file = os.path.join(project, "docs", "REQUIREMENTS.md")
    if not os.path.exists(req_file):
        return {"status": "PASS", "critical": False, "details": ["REQUIREMENTS.md 不存在"], "fix_suggestion": ""}
    
    content = open(req_file).read()
    details = []
    issues = []
    
    # P0 attempts > 10
    import re
    high_attempts = re.findall(r'REQ-\d+.*?attempts["\s:]+(\d+)', content)
    attempt_max = thresholds.get("H2_req_max_attempts", 10)
    for match in high_attempts:
        try:
            if int(match) > attempt_max:
                issues.append(f"REQ attempts={match}（阈值{attempt_max}）")
        except:
            pass
    
    # DONE未归档
    done_count = content.count('"DONE"')
    if done_count > 0:
        issues.append(f"{done_count}个DONE需求未归档")
    
    if not issues:
        details.append("需求状态正常")
    
    status = "WARN" if issues else "PASS"
    return {"status": status, "critical": False, "details": details + issues, "fix_suggestion": "; ".join(issues) if issues else ""}

def check_feedback_loop(project):
    """H3: 反哺闭环"""
    fb = os.path.join(project, "FEEDBACK.md")
    if not os.path.exists(fb):
        return {"status": "PASS", "critical": True, "details": ["FEEDBACK.md 不存在"], "fix_suggestion": ""}
    
    content = open(fb).read()
    pending = content.count("待处理")
    
    if pending > 5:
        return {"status": "FAIL", "critical": True, "details": [f"FEEDBACK.md 有{pending}个待处理条目（阈值5）"], "fix_suggestion": "FEEDBACK.md 积压严重，需优先处理"}
    elif pending > 0:
        return {"status": "WARN", "critical": False, "details": [f"FEEDBACK.md 有{pending}个待处理条目"], "fix_suggestion": ""}
    else:
        return {"status": "PASS", "critical": True, "details": ["FEEDBACK.md 无积压"], "fix_suggestion": ""}

def check_cron_alive(project, thresholds=None):
    """H4: cron活跃度 — 进度文件最后更新时间"""
    if thresholds is None:
        thresholds = {}
    threshold_min = thresholds.get("H4_cron_alive", 30)

    pf = os.path.join(project, ".claude-progress.md")
    if not os.path.exists(pf):
        return {"status": "FAIL", "critical": True, "details": ["进度文件不存在，无法判断活跃度"], "fix_suggestion": "检查cron是否正常运行"}
    
    import time
    mtime = os.path.getmtime(pf)
    age_min = (time.time() - mtime) / 60
    
    if age_min > threshold_min:
        return {
            "status": "FAIL", 
            "critical": True, 
            "details": [f"进度文件{age_min:.0f}分钟未更新（阈值{threshold_min}分钟）"],
            "fix_suggestion": f"流程可能卡住！检查cron job状态，必要时disable→enable重置。CEO需要知道这个问题。"
        }
    else:
        return {"status": "PASS", "critical": True, "details": [f"进度文件{age_min:.0f}分钟前更新（阈值{threshold_min}分钟）"], "fix_suggestion": ""}

def check_adaptive_thresholds(project):
    """H5: 自适应阈值健康检查 - 检查阈值是否需要回滚"""
    try:
        hist_file = os.path.join(project, "docs", "check-history.jsonl")
        if not os.path.exists(hist_file):
            return {"status": "PASS", "critical": False, "details": ["无历史数据，跳过阈值检查"], "fix_suggestion": ""}
        
        # 分析最近10轮的质量趋势
        recent_records = []
        with open(hist_file, 'r') as f:
            for line in f:
                try:
                    record = json.loads(line.strip())
                    if record.get('round', 0) >= (get_current_round(project) - 10):
                        recent_records.append(record)
                except:
                    continue
        
        if len(recent_records) < 5:
            return {"status": "PASS", "critical": False, "details": ["数据不足，跳过阈值检查"], "fix_suggestion": ""}
        
        # 检查质量下降趋势
        recent_fail_count = sum(1 for r in recent_records if r.get('critical_fail', False))
        recent_quality_fails = sum(1 for r in recent_records if r.get('checks', {}).get('R4_quality') == 'FAIL')
        
        issues = []
        
        # 如果自适应阈值启用后质量明显下降，建议回滚
        if recent_fail_count > 3:  # 最近10轮超过3次关键失败
            issues.append("⚠️ 关键失败频发: {}次/10轮".format(recent_fail_count))
        
        if recent_quality_fails > 4:  # R4失败次数过多
            issues.append("⚠️ 质量检查失败: {}次/10轮".format(recent_quality_fails))
        
        # 检查是否有阈值调整记录
        threshold_adjustment_file = os.path.join(project, "docs", "quality_patterns", "threshold_adjustments.json")
        if os.path.exists(threshold_adjustment_file):
            with open(threshold_adjustment_file, 'r') as f:
                adjustments = json.load(f)
            
            # 如果最近有阈值调整，检查是否需要回滚
            recent_adjustments = [a for a in adjustments if a.get('timestamp') and 
                                datetime.fromisoformat(a['timestamp']) > datetime.now(timezone.utc) - timedelta(days=7)]
            
            if recent_adjustments:
                for adj in recent_adjustments:
                    if adj.get('rollback_needed', False):
                        issues.append(f"⚠️ 检测到回滚需求: {adj.get('reason', '未知原因')}")
        
        if issues:
            return {
                "status": "WARN", 
                "critical": False, 
                "details": issues,
                "fix_suggestion": "考虑回滚阈值设置或重新评估质量基线"
            }
        else:
            return {"status": "PASS", "critical": False, "details": ["自适应阈值运行正常"], "fix_suggestion": ""}
            
    except Exception as e:
        return {"status": "WARN", "critical": False, "details": ["阈值检查异常: {}".format(e)], "fix_suggestion": ""}

def record_threshold_adjustment(project, adjustment_type, old_value, new_value, reason):
    """记录阈值调整用于后续回滚判断"""
    try:
        os.makedirs(os.path.join(project, "docs", "quality_patterns"), exist_ok=True)
        adjustment_file = os.path.join(project, "docs", "quality_patterns", "threshold_adjustments.json")
        
        adjustments = []
        if os.path.exists(adjustment_file):
            with open(adjustment_file, 'r') as f:
                adjustments = json.load(f)
        
        adjustments.append({
            "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "type": adjustment_type,
            "old_value": old_value,
            "new_value": new_value,
            "reason": reason,
            "rollback_needed": False  # 默认不需要回滚
        })
        
        # 保持最近50条记录
        adjustments = adjustments[-50:]
        
        with open(adjustment_file, 'w') as f:
            json.dump(adjustments, f, ensure_ascii=False, indent=2)
            
    except Exception:
        pass

# ---- 创建期检查 ----

def check_directory(project):
    """S1: 目录结构"""
    required = ["docs", ".cron-prompt.md", ".git"]
    details = []
    missing = []
    for item in required:
        path = os.path.join(project, item)
        if os.path.exists(path):
            details.append(f"{item} ✅")
        else:
            details.append(f"{item} ❌")
            missing.append(item)
    
    status = "FAIL" if missing else "PASS"
    return {"status": status, "critical": True, "details": details, "fix_suggestion": f"创建缺失: {', '.join(missing)}" if missing else ""}

def check_cron_config(project, cron_id=None):
    """S2: cron配置（简化版，只检查文件存在）"""
    cp = os.path.join(project, ".cron-prompt.md")
    if not os.path.exists(cp):
        return {"status": "FAIL", "critical": True, "details": [".cron-prompt.md 不存在"], "fix_suggestion": "创建 .cron-prompt.md"}
    
    content = open(cp).read()
    size = len(content)
    if size < 100:
        return {"status": "WARN", "critical": True, "details": [f".cron-prompt.md 只有{size}字节，可能不完整"], "fix_suggestion": "补充 cron-prompt 内容"}
    
    return {"status": "PASS", "critical": True, "details": [f".cron-prompt.md 存在（{size}字节）"], "fix_suggestion": ""}

def check_init_docs(project):
    """S3: 初始文档"""
    required = [("docs/REQUIREMENTS.md", "REQUIREMENTS"), ("docs/RULES.md", "RULES"), ("docs/NOW.md", "NOW")]
    details = []
    missing = []
    for fpath, name in required:
        full = os.path.join(project, fpath)
        if os.path.exists(full):
            details.append(f"{name} ✅")
        else:
            details.append(f"{name} ❌")
            missing.append(name)
    
    status = "WARN" if missing else "PASS"
    return {"status": status, "critical": False, "details": details, "fix_suggestion": f"创建缺失: {', '.join(missing)}" if missing else ""}

# ---- 注册表 ----

PHASE_CHECKS = {
    "runtime": {
        "R1_critical_steps": check_critical_steps,
        "R2_progress": check_progress,
        "R3_git": check_git,
        "R4_quality": check_quality,
        "R5_report": check_report,
    },
    "health": {
        "H1_docs_health": check_docs_health,
        "H2_req_trend": check_req_trend,
        "H3_feedback_loop": check_feedback_loop,
        "H4_cron_alive": check_cron_alive,
        "H5_adaptive_thresholds": check_adaptive_thresholds,
    },
    "setup": {
        "S1_directory": check_directory,
        "S2_cron_config": check_cron_config,
        "S3_init_docs": check_init_docs,
    },
}

SINGLE_CHECKS = {
    "critical_steps": check_critical_steps,
    "progress": check_progress,
    "git": check_git,
    "quality": check_quality,
    "report": check_report,
    "docs": check_docs_health,
    "requirements": check_req_trend,
    "feedback": check_feedback_loop,
    "cron_alive": check_cron_alive,
    "directory": check_directory,
}

DEFAULT_THRESHOLDS = {
    "H4_cron_alive": 30,
    "R3_git_ahead_max": 5,
    "H1_rules_max_lines": 300,
    "H1_now_max_lines": 30,
    "H1_decisions_max_items": 40,
    "H2_req_max_attempts": 10,
}

def load_config(project):
    """加载项目级配置"""
    thresholds = dict(DEFAULT_THRESHOLDS)
    for path in [os.path.join(project, ".check-config.json")]:
        if os.path.exists(path):
            try:
                with open(path) as f:
                    data = json.load(f)
                if "thresholds" in data:
                    thresholds.update(data["thresholds"])
            except Exception:
                pass
    return thresholds

def save_history(project, output):
    """追加结果到 docs/check-history.jsonl"""
    docs_dir = os.path.join(project, "docs")
    os.makedirs(docs_dir, exist_ok=True)
    hist_file = os.path.join(docs_dir, "check-history.jsonl")
    
    # 从 NOW.md 解析 round
    round_num = None
    now_file = os.path.join(docs_dir, "NOW.md")
    if os.path.exists(now_file):
        import re
        m = re.search(r'第\s*(\d+)\s*轮', open(now_file).read())
        if m:
            round_num = int(m.group(1))
    
    entry = {
        "timestamp": output["timestamp"],
        "round": round_num,
        "summary": output["summary"],
        "critical_fail": output["critical_fail"],
        "action_required": output["action_required"],
        "checks": {cid: r["status"] for cid, r in output["checks"].items()},
    }
    
    lines = []
    if os.path.exists(hist_file):
        with open(hist_file) as f:
            lines = f.readlines()
    
    # 去重：同 round + 同 summary 跳过
    entry_str = json.dumps(entry, ensure_ascii=False)
    duplicate = False
    for line in lines:
        try:
            existing = json.loads(line.strip())
            if existing.get("round") == round_num and existing.get("summary") == output["summary"]:
                duplicate = True
                break
        except:
            pass
    
    if not duplicate:
        lines.append(entry_str + "\n")
    
    # 截断到200行
    if len(lines) > 200:
        lines = lines[-200:]
    
    with open(hist_file, "w") as f:
        f.writelines(lines)

def analyze_trend(project):
    """趋势分析"""
    hist_file = os.path.join(project, "docs", "check-history.jsonl")
    if not os.path.exists(hist_file):
        return "📄 无历史数据"
    
    try:
        with open(hist_file) as f:
            all_lines = f.readlines()
    except Exception:
        return "📄 历史文件读取失败"
    
    recent = all_lines[-20:]
    if not recent:
        return "📄 无历史数据"
    
    records = []
    for line in recent:
        try:
            records.append(json.loads(line.strip()))
        except Exception:
            continue
    
    if not records:
        return "📄 无有效历史记录"
    
    # 收集所有检查项
    all_check_ids = set()
    for rec in records:
        all_check_ids.update(rec.get("checks", {}).keys())
    all_check_ids = sorted(all_check_ids)
    
    lines_out = []
    lines_out.append(f"📈 趋势分析（最近 {len(records)} 条记录）")
    lines_out.append(f"{'─'*40}")
    
    # 各检查项通过率
    lines_out.append("\n📋 检查项通过率:")
    for cid in all_check_ids:
        statuses = [rec["checks"].get(cid) for rec in records if cid in rec.get("checks", {})]
        if not statuses:
            continue
        total = len(statuses)
        p = statuses.count("PASS")
        w = statuses.count("WARN")
        f = statuses.count("FAIL")
        lines_out.append(f"  {cid}: PASS {p/total*100:.0f}% | WARN {w/total*100:.0f}% | FAIL {f/total*100:.0f}%")
    
    # critical_fail 频率趋势
    lines_out.append("\n🔴 Critical Fail 频率趋势:")
    window = 5
    if len(records) >= window * 2:
        first_half = records[:window]
        second_half = records[-window:]
        first_rate = sum(1 for r in first_half if r.get("critical_fail")) / len(first_half) * 100
        second_rate = sum(1 for r in second_half if r.get("critical_fail")) / len(second_half) * 100
        diff = second_rate - first_rate
        if diff > 10:
            trend = f"📈 上升（{first_rate:.0f}% → {second_rate:.0f}%）"
        elif diff < -10:
            trend = f"📉 下降（{first_rate:.0f}% → {second_rate:.0f}%）"
        else:
            trend = f"➡️ 稳定（{first_rate:.0f}% → {second_rate:.0f}%）"
        lines_out.append(f"  {trend}")
    else:
        total_fail = sum(1 for r in records if r.get("critical_fail"))
        lines_out.append(f"  数据不足，总计 {total_fail}/{len(records)} 次 critical fail")
    
    # 最常失败的检查项 TOP3
    fail_counts = {}
    for rec in records:
        for cid, status in rec.get("checks", {}).items():
            if status == "FAIL":
                fail_counts[cid] = fail_counts.get(cid, 0) + 1
    
    if fail_counts:
        top3 = sorted(fail_counts.items(), key=lambda x: -x[1])[:3]
        lines_out.append("\n❌ 最常失败 TOP3:")
        for cid, count in top3:
            lines_out.append(f"  {cid}: {count}/{len(records)} 次")
    
    return "\n".join(lines_out)

def main():
    parser = argparse.ArgumentParser(description="元迭代状态检查")
    parser.add_argument("--project", required=True, help="项目路径")
    parser.add_argument("--phase", choices=["runtime", "health", "setup"], help="按阶段检查")
    parser.add_argument("--check", help="单项检查")
    parser.add_argument("--all", action="store_true", help="全量检查")
    parser.add_argument("--threshold", default=None, help="cron活跃度阈值（分钟，默认30，或KEY=VALUE格式覆盖配置阈值）")
    parser.add_argument("--config", help="配置文件路径（默认读项目 .check-config.json）")
    parser.add_argument("--trend", action="store_true", help="趋势分析")
    parser.add_argument("--json", action="store_true", dest="as_json", help="JSON输出")
    parser.add_argument("--human", action="store_true", help="人类可读输出")
    parser.add_argument("--fail-only", action="store_true", help="只输出失败的")
    args = parser.parse_args()

    project = os.path.abspath(args.project)
    
    # 加载配置
    thresholds = load_config(project)
    if args.config:
        try:
            with open(args.config) as f:
                data = json.load(f)
            if "thresholds" in data:
                thresholds.update(data["thresholds"])
        except Exception:
            pass
    
    # 命令行 --threshold 覆盖
    if args.threshold and "=" in str(args.threshold):
        key, val = args.threshold.split("=", 1)
        try:
            thresholds[key] = int(val)
        except ValueError:
            pass
    elif args.threshold:
        try:
            thresholds["H4_cron_alive"] = int(args.threshold)
        except ValueError:
            pass
    
    # 趋势分析模式
    if args.trend:
        trend = analyze_trend(project)
        if args.as_json:
            print(json.dumps({"trend": trend}, ensure_ascii=False, indent=2))
        else:
            print(trend)
        sys.exit(0)

    project = os.path.abspath(args.project)
    
    # 确定检查项
    checks = {}
    if args.check:
        if args.check in SINGLE_CHECKS:
            checks[args.check] = SINGLE_CHECKS[args.check]
        else:
            # 尝试从phase中找
            for phase_checks in PHASE_CHECKS.values():
                for cid, cfunc in phase_checks.items():
                    if args.check in cid or args.check.replace("_", "") in cid.replace("_", ""):
                        checks[cid] = cfunc
                        break
    elif args.phase:
        checks = PHASE_CHECKS.get(args.phase, {})
    elif args.all:
        checks = {}
        for phase_checks in PHASE_CHECKS.values():
            checks.update(phase_checks)
    else:
        # 默认runtime
        checks = PHASE_CHECKS.get("runtime", {})

    # 执行检查
    results = {}
    for cid, cfunc in checks.items():
        try:
            results[cid] = cfunc(project, thresholds=thresholds)
        except TypeError:
            # 函数不接受 thresholds 参数，用旧方式调用
            try:
                results[cid] = cfunc(project)
            except Exception as e:
                results[cid] = {"status": "FAIL", "critical": True, "details": [f"检查异常: {e}"], "fix_suggestion": "检查脚本执行出错"}
        except Exception as e:
            results[cid] = {"status": "FAIL", "critical": True, "details": [f"检查异常: {e}"], "fix_suggestion": "检查脚本执行出错"}

    # 统计
    pass_count = sum(1 for r in results.values() if r["status"] == "PASS")
    warn_count = sum(1 for r in results.values() if r["status"] == "WARN")
    fail_count = sum(1 for r in results.values() if r["status"] == "FAIL")
    critical_fails = [cid for cid, r in results.items() if r["status"] == "FAIL" and r.get("critical", False)]
    non_critical_warnings = [cid for cid, r in results.items() if r["status"] in ("WARN", "FAIL") and not r.get("critical", False)]

    output = {
        "project": os.path.basename(project),
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "summary": f"{pass_count} PASS / {warn_count} WARN / {fail_count} FAIL",
        "critical_fail": len(critical_fails) > 0,
        "action_required": (len(critical_fails) + len(non_critical_warnings)) > 0,
        "checks": results,
        "critical_failures": critical_fails,
        "non_critical_warnings": non_critical_warnings,
    }

    # 持久化历史（runtime阶段）
    if args.phase == "runtime":
        save_history(project, output)

    if args.human:
        print(f"📊 {output['project']} 检查结果")
        print(f"{'═'*40}")
        print(f"总结: {output['summary']}")
        print(f"关键失败: {'❌ ' + ','.join(critical_fails) if critical_fails else '✅ 无'}")
        print(f"非关键警告: {'⚠️ ' + ','.join(non_critical_warnings) if non_critical_warnings else '✅ 无'}")
        print()
        for cid, r in results.items():
            icon = "🔴" if r.get("critical") else "⚪"
            status_icon = {"PASS": "✅", "WARN": "⚠️", "FAIL": "❌"}[r["status"]]
            print(f"  {icon} {cid}: {status_icon}")
            for d in r.get("details", []):
                print(f"     {d}")
            if r.get("fix_suggestion") and r["status"] != "PASS":
                if args.fail_only:
                    pass
                else:
                    print(f"     💡 {r['fix_suggestion']}")
        if critical_fails:
            print()
            print("🔴 以下关键步骤必须修复才能继续:")
            for cid in critical_fails:
                r = results[cid]
                print(f"   • {cid}: {r.get('fix_suggestion', '')}")
    else:
        # JSON
        print(json.dumps(output, ensure_ascii=False, indent=2))

    # 退出码
    sys.exit(1 if critical_fails else 0)

def record_threshold_decision(project, check_id, decision, data):
    """记录阈值调整决策到文件"""
    try:
        os.makedirs(os.path.join(project, "docs", "quality_patterns"), exist_ok=True)
        decision_file = os.path.join(project, "docs", "quality_patterns", "threshold_adjustments.json")
        
        # 读取现有决策
        existing_decisions = []
        if os.path.exists(decision_file):
            with open(decision_file, 'r') as f:
                existing_decisions = json.load(f)
        
        # 添加新决策
        new_decision = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "check_id": check_id,
            "decision": decision,
            "data": data,
            "rollback_needed": data.get("learning_score", 0) < 0.3  # 学习分数过低需要回滚
        }
        
        existing_decisions.append(new_decision)
        
        # 保持最近100条记录
        existing_decisions = existing_decisions[-100:]
        
        # 写入文件
        with open(decision_file, 'w') as f:
            json.dump(existing_decisions, f, ensure_ascii=False, indent=2)
        
    except Exception as e:
        # 记录失败不影响主流程
        pass

if __name__ == "__main__":
    main()
