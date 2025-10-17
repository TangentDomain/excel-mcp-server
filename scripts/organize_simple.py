#!/usr/bin/env python3
"""
Excel MCP Server - 简化目录结构整理脚本
"""

import os
import shutil
from pathlib import Path

def main():
    """主函数"""
    print("Excel MCP Server - Directory Structure Organization")
    print("=" * 60)

    current_dir = Path('.')

    # 创建必要的目录
    directories = [
        'docs/reports',
        'scripts/verification',
        'scripts/security',
        'temp'
    ]

    for dir_path in directories:
        (current_dir / dir_path).mkdir(parents=True, exist_ok=True)
        print(f"OK Created directory: {dir_path}")

    # 移动报告文件
    report_files = [
        'FINAL_STATUS_REPORT.md',
        'FINAL_VERIFICATION_REPORT.md',
        'OPENSPEC_COMPLETION_REPORT.md',
        'PROJECT_COMPLETION_SUMMARY.md',
        'PROJECT_SUMMARY.md',
        'SECURITY_ENHANCEMENT_COMPLETION_REPORT.md',
        'SECURITY_IMPROVEMENTS_SUMMARY.md',
        'SECURITY_TEST_REPORT.md',
        'SAURITY_IMPLEMENTATION_SUMMARY.md'
    ]

    moved_count = 0
    for filename in report_files:
        src = current_dir / filename
        dst = current_dir / 'docs' / 'reports' / filename

        if src.exists():
            try:
                shutil.move(str(src), str(dst))
                print(f"MOVED Report: {filename} -> docs/reports/")
                moved_count += 1
            except Exception as e:
                print(f"ERROR moving {filename}: {e}")

    # 移动验证脚本
    verify_scripts = [
        'verify_cleanup_simple.py',
        'verify_security_features.py',
        'verify_temp_cleanup.py'
    ]

    for filename in verify_scripts:
        src = current_dir / filename
        dst = current_dir / 'scripts' / 'verification' / filename

        if src.exists():
            try:
                shutil.move(str(src), str(dst))
                print(f"MOVED Verification: {filename} -> scripts/verification/")
                moved_count += 1
            except Exception as e:
                print(f"ERROR moving {filename}: {e}")

    # 移动安全脚本
    security_scripts = [
        'run_security_tests.py'
    ]

    for filename in security_scripts:
        src = current_dir / filename
        dst = current_dir / 'scripts' / 'security' / filename

        if src.exists():
            try:
                shutil.move(str(src), str(dst))
                print(f"MOVED Security: {filename} -> scripts/security/")
                moved_count += 1
            except Exception as e:
                print(f"ERROR moving {filename}: {e}")

    # 移动临时脚本
    temp_scripts = [
        'run-all-tests.py'
    ]

    for filename in temp_scripts:
        src = current_dir / filename
        dst = current_dir / 'temp' / filename

        if src.exists():
            try:
                shutil.move(str(src), str(dst))
                print(f"MOVED Temp: {filename} -> temp/")
                moved_count += 1
            except Exception as e:
                print(f"ERROR moving {filename}: {e}")

    # 移动安全文档
    security_docs = [
        'EXCEL_SECURITY_BEST_PRACTICES.md',
        'SECURITY_FOCUSED_LLM_PROMPT.md'
    ]

    for filename in security_docs:
        src = current_dir / filename
        dst = current_dir / 'docs' / filename

        if src.exists():
            try:
                shutil.move(str(src), str(dst))
                print(f"MOVED Security Doc: {filename} -> docs/")
                moved_count += 1
            except Exception as e:
                print(f"ERROR moving {filename}: {e}")

    # 清理临时文件
    temp_patterns = [
        '*temp*.py',
        'test_*template*.py',
        'comprehensive_verification.py'
    ]

    cleaned_count = 0
    for pattern in temp_patterns:
        for file_path in current_dir.glob(pattern):
            if file_path.is_file():
                try:
                    dst = current_dir / 'temp' / file_path.name
                    shutil.move(str(file_path), str(dst))
                    print(f"CLEANUP: {file_path.name} -> temp/")
                    cleaned_count += 1
                except Exception as e:
                    print(f"ERROR cleaning {file_path.name}: {e}")

    # 创建目录索引
    index_content = """# Excel MCP Server - Directory Structure Index

## Directory Organization

### Core Code
- `src/` - Source code directory
- `tests/` - Test directory

### Scripts
- `scripts/` - Script tools directory
  - `verification/` - Verification scripts
  - `security/` - Security related scripts

### Documentation
- `docs/` - Documentation directory
  - `reports/` - Project reports
  - `archive/` - Archived documents

### Temporary Files
- `temp/` - Temporary files directory

## File Categories

### Reports (docs/reports/)
- FINAL_VERIFICATION_REPORT.md - Final verification report
- PROJECT_COMPLETION_SUMMARY.md - Project completion summary
- SECURITY_ENHANCEMENT_COMPLETION_REPORT.md - Security enhancement report

### Verification Scripts (scripts/verification/)
- verify_cleanup_simple.py - Simple cleanup verification
- verify_security_features.py - Security features verification
- verify_temp_cleanup.py - Temp file cleanup verification

### Security Scripts (scripts/security/)
- run_security_tests.py - Run security tests

### Security Documentation (docs/)
- EXCEL_SECURITY_BEST_PRACTICES.md - Security best practices
- SECURITY_FOCUSED_LLM_PROMPT.md - Security focused LLM prompt

---
*Generated by organize_simple.py*
"""

    with open('DIRECTORY_INDEX.md', 'w', encoding='utf-8') as f:
        f.write(index_content)

    print("CREATED: DIRECTORY_INDEX.md")

    # 总结
    print("\n" + "=" * 60)
    print(f"ORGANIZATION COMPLETE!")
    print(f"Moved {moved_count} files")
    print(f"Cleaned {cleaned_count} temporary files")
    print(f"Created directory index")
    print("\nSee DIRECTORY_INDEX.md for new structure")

if __name__ == "__main__":
    main()