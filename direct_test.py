#!/usr/bin/env python3
"""
直接运行测试函数
"""
import sys
import os
from pathlib import Path

# 设置路径
work_dir = Path('/root/.openclaw/workspace/excel-mcp-server')
sys.path.insert(0, str(work_dir / 'src'))
os.chdir(str(work_dir))

# 导入测试模块
import test_fixes

if __name__ == '__main__':
    print("\n" + "="*80)
    print("ExcelMCP 修复验证测试 - 直接运行")
    print("="*80)

    try:
        test_fixes.test_same_file_join()
        test_fixes.test_group_concat_complex_expression()

        print("\n" + "="*80)
        print("✅ 所有测试完成")
        print("="*80)

    except Exception as e:
        print(f"\n❌ 测试执行出错: {e}")
        import traceback
        traceback.print_exc()
