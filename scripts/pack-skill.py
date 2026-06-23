#!/usr/bin/env python3
"""将 Excel skill 打包成 ZIP 包，包含 SKILL.md + PyInstaller 可执行文件。

用法:
  python scripts/pack-skill.py                    # 打包当前平台
  python scripts/pack-skill.py --no-build          # 跳过 PyInstaller，只打包已有可执行文件
  python scripts/pack-skill.py --output dist/      # 指定输出目录

产物结构（ZIP 内）:
  excel-skill/
  ├── SKILL.md              # skill 文档（工具选择决策树）
  └── bin/
      └── excel-cli(.exe)   # 当前平台的可执行文件

ZIP 文件名: excel-skill-{version}-{platform}-{arch}.zip
"""
from __future__ import annotations

import argparse
import os
import platform
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

# 项目根目录
SCRIPT_DIR = Path(__file__).parent.resolve()
ROOT_DIR = SCRIPT_DIR.parent
SRC_DIR = ROOT_DIR / "src"
SKILL_DIR = ROOT_DIR / "scripts" / "excel-skill"
CLI_SOURCE = ROOT_DIR / "scripts" / "excel-cli.py"


def detect_platform() -> tuple[str, str]:
    """检测当前平台和架构，返回 (platform, arch)。"""
    s = sys.platform.lower()
    if s.startswith("win"):
        pf = "windows"
    elif s.startswith("darwin"):
        pf = "macos"
    else:
        pf = "linux"

    machine = platform.machine().lower()
    if machine in ("x86_64", "amd64"):
        arch = "amd64"
    elif machine in ("aarch64", "arm64"):
        arch = "arm64"
    else:
        arch = machine
    return pf, arch


def get_version() -> str:
    """从 excel-cli.py 读取 VERSION。"""
    import re

    content = CLI_SOURCE.read_text(encoding="utf-8")
    match = re.search(r'^VERSION\s*=\s*"([^"]+)"', content, re.MULTILINE)
    return match.group(1) if match else "0.0.0"


def build_executable() -> Path:
    """用 PyInstaller --onedir 构建可执行文件，返回产物目录路径。

    --onedir 模式：产物是一个目录（excel-cli/），包含 exe + _internal/ 依赖。
    启动快（无需每次解压），适合日常 CLI 使用。
    """
    build_dir = ROOT_DIR / "build" / "pyinstaller"
    dist_dir = build_dir / "dist"
    spec_dir = build_dir

    # 清理旧构建
    if dist_dir.exists():
        shutil.rmtree(dist_dir)

    build_dir.mkdir(parents=True, exist_ok=True)

    exe_name = "excel-cli"

    # 全局 Python 安装含大量无关包（torch/onnxruntime/scipy 等），必须显式排除
    excludes = [
        "torch", "torchvision", "torchaudio", "tensorboard",
        "onnxruntime", "scipy", "matplotlib",
        "numba", "yt_dlp", "yt_dlp_ejs", "curl_cffi",
        "PyQt5", "PyQt6", "PySide2", "PySide6", "tkinter",
        "IPython", "jupyter", "notebook", "pytest",
        "PIL", "cv2", "skimage", "sklearn",
    ]
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",
        "--name", exe_name,
        "--noconfirm",
        "--paths", str(SRC_DIR),
        "--collect-data", "openpyxl",
        "--hidden-import", "openpyxl",
        "--hidden-import", "et_xmlfile",
        "--hidden-import", "sqlglot",
        "--hidden-import", "python_calamine",
        "--hidden-import", "xlcalculator",
        *[arg for excl in excludes for arg in ("--exclude-module", excl)],
        "--workpath", str(build_dir / "build"),
        "--distpath", str(dist_dir),
        "--specpath", str(spec_dir),
        "--log-level", "WARN",
        str(CLI_SOURCE),
    ]

    print(f"  🔨 PyInstaller 构建中（--onedir 模式）...")
    result = subprocess.run(
        cmd,
        cwd=str(build_dir),
    )

    if result.returncode != 0:
        print("❌ PyInstaller 构建失败 (见上方日志)")
        sys.exit(1)

    # --onedir 产物是一个目录：dist/excel-cli/
    app_dir = dist_dir / exe_name
    if not app_dir.is_dir():
        print(f"❌ 构建产物目录未找到: {app_dir}")
        sys.exit(1)

    # 统计总大小和文件数
    total_size = sum(f.stat().st_size for f in app_dir.rglob("*") if f.is_file())
    size_mb = total_size / (1024 * 1024)
    file_count = sum(1 for f in app_dir.rglob("*") if f.is_file())
    print(f"  ✅ 构建完成: {app_dir.name}/ ({size_mb:.1f} MB, {file_count} 文件)")
    return app_dir


def package_zip(app_dir: Path, output_dir: Path) -> Path:
    """将 SKILL.md + onedir 产物目录打包成 ZIP。"""
    pf, arch = detect_platform()
    version = get_version()

    zip_name = f"excel-skill-{version}-{pf}-{arch}.zip"
    zip_path = output_dir / zip_name

    output_dir.mkdir(parents=True, exist_ok=True)

    # 删除旧 ZIP
    if zip_path.exists():
        zip_path.unlink()

    print(f"  📦 打包中: {zip_name}")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        # SKILL.md — 放在 excel-skill/ 根
        skill_md = SKILL_DIR / "SKILL.md"
        if skill_md.exists():
            zf.write(skill_md, "excel-skill/SKILL.md")
            print(f"  + excel-skill/SKILL.md")
        else:
            print(f"  ⚠️ SKILL.md 未找到: {skill_md}")

        # 递归写入 onedir 产物目录（exe + _internal/ 依赖）
        for file_path in app_dir.rglob("*"):
            if file_path.is_file():
                arcname = f"excel-skill/bin/{file_path.relative_to(app_dir)}"
                zf.write(file_path, arcname)
                print(f"  + {arcname}")

        # 版本信息
        import json

        version_info = {
            "version": version,
            "platform": pf,
            "arch": arch,
            "executable": "excel-cli.exe" if pf == "windows" else "excel-cli",
            "python_version": sys.version.split()[0],
        }
        zf.writestr(
            "excel-skill/manifest.json",
            json.dumps(version_info, ensure_ascii=False, indent=2),
        )
        print(f"  + excel-skill/manifest.json")

    size_mb = zip_path.stat().st_size / (1024 * 1024)
    print(f"  ✅ ZIP 产物: {zip_path} ({size_mb:.1f} MB)")
    return zip_path


def verify_zip(zip_path: Path) -> bool:
    """验证 ZIP 产物完整性。"""
    print(f"  🔍 验证 ZIP 完整性...")
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            bad = zf.testzip()
            if bad:
                print(f"  ❌ 损坏文件: {bad}")
                return False

            names = zf.namelist()
            required = ["excel-skill/SKILL.md", "excel-skill/manifest.json"]
            exe_found = any(n.startswith("excel-skill/bin/") for n in names)

            for req in required:
                if req not in names:
                    print(f"  ❌ 缺少必需文件: {req}")
                    return False

            if not exe_found:
                print(f"  ❌ 缺少可执行文件")
                return False

            print(f"  ✅ ZIP 完整 ({len(names)} 个文件)")
            for n in names:
                info = zf.getinfo(n)
                print(f"     {n} ({info.file_size} bytes)")
            return True
    except Exception as e:
        print(f"  ❌ 验证失败: {e}")
        return False


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="将 Excel skill 打包成 ZIP（含可执行文件）"
    )
    parser.add_argument(
        "--no-build",
        action="store_true",
        help="跳过 PyInstaller 构建，使用已有可执行文件",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="dist",
        help="输出目录（默认 dist/）",
    )
    parser.add_argument(
        "--exe",
        default=None,
        help="指定已有 onedir 目录路径（配合 --no-build 使用，指向 excel-cli/ 目录）",
    )
    args = parser.parse_args(argv)

    pf, arch = detect_platform()
    version = get_version()

    print(f"=== Excel Skill 打包 ===")
    print(f"  版本: {version}")
    print(f"  平台: {pf}-{arch}")
    print()

    output_dir = ROOT_DIR / args.output

    # 1. 构建可执行文件（或使用已有的）
    if args.no_build:
        if args.exe:
            app_dir = Path(args.exe)
        else:
            # 查找默认 onedir 目录
            exe_name = "excel-cli"
            app_dir = ROOT_DIR / "build" / "pyinstaller" / "dist" / exe_name

        if not app_dir.is_dir():
            print(f"❌ onedir 产物目录未找到: {app_dir}")
            print(f"   去掉 --no-build 重新构建，或用 --exe 指定路径")
            return 1
        print(f"  📎 使用已有产物目录: {app_dir}")
    else:
        app_dir = build_executable()

    print()

    # 2. 打包 ZIP
    zip_path = package_zip(app_dir, output_dir)

    print()

    # 3. 验证
    if not verify_zip(zip_path):
        print("❌ 打包验证失败")
        return 1

    print()
    print(f"✅ 打包完成: {zip_path}")
    print(f"   安装方式: 解压到 ~/.omp/agent/skills/，skill 自动发现")
    return 0
if __name__ == "__main__":
    raise SystemExit(main())
