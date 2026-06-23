#!/usr/bin/env python3
"""scripts/verify-output.py — Loop I2 独立验证器

4 项检查，score 归一化到 [0, 1]。输出 JSON。
"""

import json
import subprocess
import sys
from pathlib import Path

PROTECTED_FILES = [
    "src/excel_mcp_server_fastmcp/api/sql_executor.py",
    "src/excel_mcp_server_fastmcp/core/excel_operations.py",
    "src/excel_mcp_server_fastmcp/core/backup_manager.py",
]

CHECKS = [
    {"name": "ruff-format", "cmd": ["ruff", "format", "--check", "src/", "tests/"], "critical": True},
    {"name": "ruff-check", "cmd": ["ruff", "check", "src/", "tests/"], "critical": True},
    {"name": "invariants", "cmd": ["python", "-m", "pytest", "tests/invariants/", "-q", "--timeout=30"], "critical": True},
    {"name": "docstring", "cmd": ["python", "-c", _docstring_check_code()], "critical": False},
]


def _docstring_check_code():
    return (
        "import ast, sys, pathlib;"
        "errors = [];"
        "root = pathlib.Path('src/excel_mcp_server_fastmcp');"
        "for f in root.rglob('*.py'):;"
        "  tree = ast.parse(f.read_text());"
        "  for node in ast.walk(tree):;"
        "    if isinstance(node, ast.FunctionDef) and not node.name.startswith('_'):;"
        "      doc = ast.get_docstring(node);"
        "      if not doc: errors.append(f'{f}:{node.lineno}:{node.name}')"
        "print(len(errors))"
    )


def run_check(check):
    try:
        subprocess.run(check["cmd"], capture_output=True, timeout=120, check=True)
        return {"passed": True}
    except subprocess.CalledProcessError as e:
        return {"passed": False, "reason": e.stderr.decode("utf-8", errors="replace")[-200:]}
    except FileNotFoundError:
        return {"passed": False, "reason": f"command not found: {check['cmd'][0]}"}


def main():
    results = []
    total_weight = 0
    passed_weight = 0

    for check in CHECKS:
        r = run_check(check)
        r["name"] = check["name"]
        r["critical"] = check["critical"]
        results.append(r)

        weight = 2 if check["critical"] else 1
        total_weight += weight
        if r["passed"]:
            passed_weight += weight

    score = passed_weight / total_weight if total_weight > 0 else 0.0

    # protected files check
    protected_ok = all(Path(f).exists() for f in PROTECTED_FILES)

    output = {
        "score": round(score, 4),
        "checks": results,
        "protected_files_exist": protected_ok,
        "verdict": "pass" if score >= 0.75 else "fail",
    }

    print(json.dumps(output, ensure_ascii=False))
    sys.exit(0 if output["verdict"] == "pass" else 2)


if __name__ == "__main__":
    main()
