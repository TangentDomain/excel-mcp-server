#!/usr/bin/env python3
"""
Excel MCP Server - 项目目录结构整理脚本
整理项目文件到合适的目录结构中
"""

import os
import shutil
from pathlib import Path

def organize_project_files():
    """整理项目文件到合适的目录结构"""

    current_dir = Path('.')

    # 创建必要的目录
    directories_to_create = [
        'docs/archive',
        'docs/reports',
        'scripts/verification',
        'scripts/security',
        'temp'  # 临时脚本目录
    ]

    for dir_path in directories_to_create:
        (current_dir / dir_path).mkdir(parents=True, exist_ok=True)
        print(f"✅ 创建目录: {dir_path}")

    # 需要移动的文件映射
    files_to_move = {
        # 报告类文件移动到 docs/reports
        'FINAL_STATUS_REPORT.md': 'docs/reports/FINAL_STATUS_REPORT.md',
        'FINAL_VERIFICATION_REPORT.md': 'docs/reports/FINAL_VERIFICATION_REPORT.md',
        'OPENSPEC_COMPLETION_REPORT.md': 'docs/reports/OPENSPEC_COMPLETION_REPORT.md',
        'PROJECT_COMPLETION_SUMMARY.md': 'docs/reports/PROJECT_COMPLETION_SUMMARY.md',
        'PROJECT_SUMMARY.md': 'docs/reports/PROJECT_SUMMARY.md',
        'SECURITY_ENHANCEMENT_COMPLETION_REPORT.md': 'docs/reports/SECURITY_ENHANCEMENT_COMPLETION_REPORT.md',
        'SECURITY_IMPROVEMENTS_SUMMARY.md': 'docs/reports/SECURITY_IMPROVEMENTS_SUMMARY.md',
        'SECURITY_TEST_REPORT.md': 'docs/reports/SECURITY_TEST_REPORT.md',
        'SAURITY_IMPLEMENTATION_SUMMARY.md': 'docs/reports/SAURITY_IMPLEMENTATION_SUMMARY.md',  # 保持原名

        # 验证脚本移动到 scripts/verification
        'verify_cleanup_simple.py': 'scripts/verification/verify_cleanup_simple.py',
        'verify_security_features.py': 'scripts/verification/verify_security_features.py',
        'verify_temp_cleanup.py': 'scripts/verification/verify_temp_cleanup.py',

        # 安全相关移动到 scripts/security
        'run_security_tests.py': 'scripts/security/run_security_tests.py',

        # 临时脚本移动到 temp
        'run-all-tests.py': 'temp/run-all-tests.py',

        # 安全文档移动到 docs
        'EXCEL_SECURITY_BEST_PRACTICES.md': 'docs/EXCEL_SECURITY_BEST_PRACTICES.md',
        'SECURITY_FOCUSED_LLM_PROMPT.md': 'docs/SECURITY_FOCUSED_LLM_PROMPT.md'
    }

    moved_count = 0

    for src_file, dst_file in files_to_move.items():
        src_path = current_dir / src_file
        dst_path = current_dir / dst_file

        if src_path.exists():
            try:
                # 确保目标目录存在
                dst_path.parent.mkdir(parents=True, exist_ok=True)

                # 移动文件
                shutil.move(str(src_path), str(dst_path))
                print(f"📁 移动文件: {src_file} -> {dst_file}")
                moved_count += 1

            except Exception as e:
                print(f"❌ 移动文件失败 {src_file}: {e}")
        else:
            print(f"⚠️  文件不存在: {src_file}")

    return moved_count

def cleanup_temp_files():
    """清理临时和测试文件"""

    current_dir = Path('.')

    # 需要清理的文件模式
    temp_patterns = [
        '*temp*.py',
        'test_*template*.py',
        'comprehensive_verification.py',
        'test_*enhanced*.py'
    ]

    cleaned_count = 0

    for pattern in temp_patterns:
        for file_path in current_dir.glob(pattern):
            if file_path.is_file():
                try:
                    # 移动到 temp 目录
                    temp_dir = current_dir / 'temp'
                    temp_dir.mkdir(exist_ok=True)

                    dst_path = temp_dir / file_path.name
                    shutil.move(str(file_path), str(dst_path))
                    print(f"🗂️  清理临时文件: {file_path.name} -> temp/")
                    cleaned_count += 1

                except Exception as e:
                    print(f"❌ 清理文件失败 {file_path.name}: {e}")

    return cleaned_count

def create_directory_index():
    """创建目录索引文件"""

    index_content = """# Excel MCP Server - 目录结构索引

## 📁 目录组织

### 核心代码
- `src/` - 源代码目录
  - `server.py` - MCP 服务器入口
  - `api/` - API 业务逻辑层
  - `core/` - 核心操作层
  - `utils/` - 工具层

### 测试文件
- `tests/` - 测试目录
  - `test_*.py` - 各种测试文件
  - `conftest.py` - 测试配置

### 脚本工具
- `scripts/` - 脚本工具目录
  - `verification/` - 验证脚本
  - `security/` - 安全相关脚本
  - `monitor*.py` - 监控脚本

### 文档
- `docs/` - 文档目录
  - `reports/` - 项目报告
  - `archive/` - 归档文档
  - `*.md` - 各种文档

### 配置文件
- `pyproject.toml` - 项目配置
- `*.json` - 配置文件
- `*.md` - 说明文档

### 临时文件
- `temp/` - 临时文件目录

## 📋 文件分类

### 📊 报告文件 (docs/reports/)
- FINAL_VERIFICATION_REPORT.md - 最终验证报告
- PROJECT_COMPLETION_SUMMARY.md - 项目完成总结
- SECURITY_ENHANCEMENT_COMPLETION_REPORT.md - 安全增强完成报告
- 其他项目报告...

### 🔧 验证脚本 (scripts/verification/)
- verify_cleanup_simple.py - 简化清理验证
- verify_security_features.py - 安全功能验证
- verify_temp_cleanup.py - 临时文件清理验证

### 🛡️ 安全脚本 (scripts/security/)
- run_security_tests.py - 运行安全测试

### 📝 文档文件 (docs/)
- 游戏开发Excel配置表比较指南.md - 游戏开发指南
- EXCEL_SECURITY_BEST_PRACTICES.md - 安全最佳实践
- 其他项目文档...

---
*此文件由 organize_project_structure.py 自动生成*
"""

    with open('DIRECTORY_INDEX.md', 'w', encoding='utf-8') as f:
        f.write(index_content)

    print("📋 创建目录索引: DIRECTORY_INDEX.md")

def main():
    """主函数"""
    print("Excel MCP Server - 目录结构整理")
    print("=" * 50)

    # 整理项目文件
    print("\n1. 整理项目文件...")
    moved_count = organize_project_files()

    # 清理临时文件
    print("\n2. 清理临时文件...")
    cleaned_count = cleanup_temp_files()

    # 创建目录索引
    print("\n3. 创建目录索引...")
    create_directory_index()

    # 总结
    print("\n" + "=" * 50)
    print(f"✅ 整理完成!")
    print(f"📁 移动了 {moved_count} 个文件")
    print(f"🗂️  清理了 {cleaned_count} 个临时文件")
    print(f"📋 创建了目录索引文件")

    print("\n📂 建议查看 DIRECTORY_INDEX.md 了解新的目录结构")

if __name__ == "__main__":
    main()