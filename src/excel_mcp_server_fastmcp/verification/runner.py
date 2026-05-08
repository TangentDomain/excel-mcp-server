"""闭环验证运行器：固定 fixture → temp copy → baseline 对比 → artifacts。"""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any

from ..api.advanced_sql_query import (
    execute_advanced_delete_query,
    execute_advanced_insert_query,
    execute_advanced_sql_query,
    execute_advanced_update_query,
)
from .artifacts import create_run_directory, write_json, write_summary
from .diff import (
    compare_structured,
    normalize_select_result,
    normalize_update_result,
    normalize_value,
)
from .scenarios import BASELINE_DIR, ARTIFACT_ROOT, VerificationCase, get_verification_cases


def _load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _save_baseline(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")


def _run_select_case(case: VerificationCase) -> dict[str, Any]:
    result = execute_advanced_sql_query(
        str(case.fixture_path),
        case.sql,
        sheet_name=case.sheet_name,
        include_headers=True,
        output_format="json",
    )
    return normalize_select_result(result)


def _run_update_case(case: VerificationCase, run_dir: Path) -> dict[str, Any]:
    case_dir = run_dir / case.case_id
    case_dir.mkdir(parents=True, exist_ok=True)
    temp_input = case_dir / (case.mutation_copy_name or case.fixture_name)
    shutil.copy2(case.fixture_path, temp_input)

    update_result = execute_advanced_update_query(str(temp_input), case.sql, sheet_name=case.sheet_name, dry_run=False)
    post_query = execute_advanced_sql_query(
        str(temp_input),
        "SELECT skill_id, skill_name, damage, cooldown FROM 技能配置 WHERE skill_id = 'SK001'",
        sheet_name=case.sheet_name,
        include_headers=True,
        output_format="json",
    )
    return {
        "update_result": normalize_update_result(update_result),
        "post_state": normalize_select_result(post_query),
    }


def _run_case(case: VerificationCase, run_dir: Path) -> dict[str, Any]:
    if case.kind == "select":
        return {
            "kind": case.kind,
            "fixture": str(case.fixture_path),
            "sql": case.sql,
            "sheet_name": case.sheet_name,
            "actual": _run_select_case(case),
        }
    if case.kind == "update":
        payload = _run_update_case(case, run_dir)
        return {
            "kind": case.kind,
            "fixture": str(case.fixture_path),
            "sql": case.sql,
            "sheet_name": case.sheet_name,
            "actual": payload,
        }
    raise ValueError(f"不支持的验证场景类型: {case.kind}")


def run_verification(case_ids: list[str] | None = None, update_baselines: bool = False) -> dict[str, Any]:
    """运行闭环验证。"""

    run_dir = create_run_directory(ARTIFACT_ROOT)
    selected_cases = get_verification_cases()
    if case_ids:
        wanted = set(case_ids)
        selected_cases = [case for case in selected_cases if case.case_id in wanted]

    results: list[dict[str, Any]] = []
    all_passed = True

    for case in selected_cases:
        case_result = _run_case(case, run_dir)
        actual = case_result["actual"]
        baseline_path = case.baseline_path
        expected = None
        diffs: list[dict[str, Any]] = []
        baseline_exists = baseline_path.exists()

        if update_baselines:
            _save_baseline(baseline_path, actual)
            expected = actual
        else:
            if not baseline_exists:
                diffs = [{"path": "$", "type": "missing_baseline", "expected": None, "actual": actual}]
                all_passed = False
            else:
                expected = _load_json(baseline_path)
                diffs = compare_structured(expected, actual)
                if diffs:
                    all_passed = False

        case_dir = run_dir / case.case_id
        case_dir.mkdir(parents=True, exist_ok=True)
        write_json(case_dir / "actual.json", actual)
        if expected is not None:
            write_json(case_dir / "expected.json", expected)
        if diffs:
            write_json(case_dir / "diff.json", diffs)

        results.append({
            "case_id": case.case_id,
            "kind": case.kind,
            "fixture": str(case.fixture_path),
            "sql": case.sql,
            "sheet_name": case.sheet_name,
            "baseline_path": str(baseline_path),
            "baseline_exists": baseline_exists,
            "passed": not diffs,
            "diff_count": len(diffs),
            "diffs": diffs,
            "artifact_dir": str(case_dir),
        })

    summary = {
        "success": all_passed,
        "update_baselines": update_baselines,
        "run_dir": str(run_dir),
        "cases": results,
        "baseline_dir": str(BASELINE_DIR),
    }
    summary_path = write_summary(run_dir, summary)
    summary["summary_path"] = str(summary_path)
    return summary


__all__ = ["run_verification"]
