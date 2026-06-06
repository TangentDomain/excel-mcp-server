#!/usr/bin/env python3
"""L4 置信度审计 — 解析 Claude Code audit JSONL 日志，计算操作统计和 Brier Score。

Usage:
    python scripts/confidence_audit.py              # 默认读取 ~/.claude/audit/
    python scripts/confidence_audit.py --dir /path  # 指定日志目录
"""

import argparse
import json
import os
from collections import defaultdict
from pathlib import Path


def load_recent_logs(log_dir: str, days: int = 7) -> list[dict]:
    """Load recent JSONL audit logs."""
    p = Path(log_dir)
    if not p.is_dir():
        return []

    files = sorted(p.glob("*.jsonl"))
    if not files:
        return []

    # Take last `days` files (one file per day by naming convention)
    recent = files[-days:]
    entries = []
    for fp in recent:
        with open(fp) as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        entries.append(json.loads(line))
                    except json.JSONDecodeError:
                        pass
    return entries


def compute_stats(entries: list[dict]) -> dict:
    """Compute operation statistics from audit entries."""
    ops = defaultdict(int)
    errors = 0
    for e in entries:
        tool = e.get("payload", {}).get("tool_name", "?")
        ops[tool] += 1
        if e.get("payload", {}).get("error"):
            errors += 1

    total = sum(ops.values())
    return {
        "total_operations": total,
        "by_tool": dict(ops),
        "errors": errors,
    }


def main():
    parser = argparse.ArgumentParser(description="L4 confidence audit")
    parser.add_argument("--dir", default=os.path.expanduser("~/.claude/audit"))
    parser.add_argument("--days", type=int, default=7)
    args = parser.parse_args()

    entries = load_recent_logs(args.dir, args.days)
    if not entries:
        print(f"No audit logs found in {args.dir}")
        print("Brier Score tracking will begin after first Claude Code session")
        return

    stats = compute_stats(entries)
    print(f"Audit period: last {args.days} days")
    print(f"Total operations: {stats['total_operations']}")
    print(f"By tool: {stats['by_tool']}")
    print(f"Errors: {stats['errors']}")
    print("Brier Score: N/A (confidence calibration requires outcome tracking)")


if __name__ == "__main__":
    main()
