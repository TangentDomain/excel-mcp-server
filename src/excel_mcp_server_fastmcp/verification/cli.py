"""命令行入口：运行 baseline 驱动的闭环验证。"""

from __future__ import annotations

import argparse
import json
from typing import Any

from .runner import run_verification


def build_parser() -> argparse.ArgumentParser:
    """构建 CLI 参数解析器。"""
    parser = argparse.ArgumentParser(description="运行 Excel MCP Server baseline 驱动闭环验证")
    parser.add_argument("--case", action="append", dest="cases", help="只运行指定 case_id，可重复传入")
    parser.add_argument("--update-baselines", action="store_true", help="将实际结果写回 baseline")
    return parser


def main(argv: list[str] | None = None) -> int:
    """CLI 入口：运行验证并输出 JSON 结果。"""
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.update_baselines and not args.cases:
        parser.error("--update-baselines 必须同时指定至少一个 --case，避免误更新全部 baseline")
    result: dict[str, Any] = run_verification(case_ids=args.cases, update_baselines=args.update_baselines)
    print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
    return 0 if result.get("success") else 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
