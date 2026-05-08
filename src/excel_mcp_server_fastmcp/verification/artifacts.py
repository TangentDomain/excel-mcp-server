"""验证运行时的 artifact 管理。"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from .scenarios import ARTIFACT_ROOT


def create_run_directory(base_dir: Path | None = None) -> Path:
    """创建单次验证运行目录。"""

    root = base_dir or ARTIFACT_ROOT
    root.mkdir(parents=True, exist_ok=True)
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    run_dir = root / run_id
    run_dir.mkdir(parents=True, exist_ok=False)
    return run_dir



def write_json(path: Path, payload: Any) -> Path:
    """写出 UTF-8 JSON artifact。"""

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")
    return path



def write_summary(run_dir: Path, summary: dict[str, Any]) -> Path:
    return write_json(run_dir / "summary.json", summary)


__all__ = ["create_run_directory", "write_json", "write_summary"]
