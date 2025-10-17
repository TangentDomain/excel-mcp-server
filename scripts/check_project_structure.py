#!/usr/bin/env python3
"""
Check Excel MCP Server project structure integrity
"""

import os
import sys
from pathlib import Path


def check_project_structure():
    """Check if the project maintains the required directory structure"""

    required_dirs = [
        'src',
        'src/api',
        'src/core',
        'src/utils',
        'src/models',
        'tests',
        'tests/test_data',
        'scripts',
        'docs'
    ]

    required_files = [
        'src/server.py',
        'src/api/excel_operations.py',
        'src/core/excel_reader.py',
        'src/core/excel_writer.py',
        'src/core/excel_manager.py',
        'src/utils/formatter.py',
        'src/utils/validators.py',
        'src/models/types.py',
        'tests/conftest.py',
        'tests/test_api_excel_operations.py',
        'tests/test_core.py',
        'tests/test_server.py',
        'pyproject.toml',
        'README.md',
        'CLAUDE.md'
    ]

    missing_dirs = []
    missing_files = []

    # Check directories
    for dir_path in required_dirs:
        if not Path(dir_path).exists():
            missing_dirs.append(dir_path)

    # Check files
    for file_path in required_files:
        if not Path(file_path).exists():
            missing_files.append(file_path)

    # Report results
    if missing_dirs:
        print(f"❌ Missing required directories:")
        for dir_path in missing_dirs:
            print(f"   - {dir_path}")
        return False

    if missing_files:
        print(f"❌ Missing required files:")
        for file_path in missing_files:
            print(f"   - {file_path}")
        return False

    print("✅ Project structure is valid")

    # Optional: Check for extra structure consistency
    extra_checks = []

    # Check __init__.py files in Python packages
    for root_dir in ['src', 'src/api', 'src/core', 'src/utils', 'src/models', 'tests']:
        init_file = Path(root_dir) / '__init__.py'
        if not init_file.exists():
            extra_checks.append(f"Missing __init__.py in {root_dir}")

    if extra_checks:
        print("⚠️  Optional issues found:")
        for issue in extra_checks:
            print(f"   - {issue}")

    return True


if __name__ == "__main__":
    success = check_project_structure()
    sys.exit(0 if success else 1)