#!/usr/bin/env python3
"""Excel skill 自举器。自动从 GitHub 获取并安装最新版本。"""
import subprocess, sys, os

GIT_URL = "https://github.com/TangentDomain/excel-mcp-server"
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VENV_DIR = os.path.join(SKILL_DIR, ".venv")
VENV_ENTRY = os.path.join(VENV_DIR, "Scripts", "excel-cli.exe")
VENV_PYTHON = os.path.join(VENV_DIR, "Scripts", "python.exe")


def _ensure_venv():
    if os.path.isfile(VENV_ENTRY):
        return VENV_ENTRY
    sys.stderr.write("[excel] 首次使用，自动安装...\n")
    subprocess.run(["uv", "venv", VENV_DIR, "--quiet"],
                   capture_output=True, check=True, timeout=60)
    subprocess.run(["uv", "pip", "install", "--python", VENV_PYTHON,
                    f"git+{GIT_URL}", "--quiet"],
                   capture_output=True, check=True, timeout=180)
    sys.stderr.write("[excel] 安装完成！\n")
    return VENV_ENTRY


def main():
    entry = _ensure_venv()
    sys.exit(subprocess.run([entry] + sys.argv[1:]).returncode)


if __name__ == "__main__":
    main()
