#!/usr/bin/env python3
"""
简单的转换器测试
"""
import sys
import os
from pathlib import Path

# 添加src到路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

# 测试导入
try:
    from core.excel_converter import ExcelConverter
    print("✅ ExcelConverter导入成功")

    # 测试类是否可以实例化（不需要真实文件）
    print(f"ExcelConverter类: {ExcelConverter}")
    print(f"方法列表: {[m for m in dir(ExcelConverter) if not m.startswith('_')]}")

except ImportError as e:
    print(f"❌ 导入失败: {e}")

except Exception as e:
    print(f"❌ 其他错误: {e}")
