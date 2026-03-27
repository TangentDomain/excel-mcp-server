#!/usr/bin/env python3
"""
REQ-029 最简单验证脚本
"""

import sys
import tempfile
import pandas as pd
from pathlib import Path

# 添加项目路径到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src"))

from excel_mcp_server_fastmcp.server import excel_describe_table

def simple_test():
    print("开始简单测试...")
    
    # 创建临时文件
    temp_dir = tempfile.mkdtemp()
    test_file = f"{temp_dir}/simple.xlsx"
    
    # 创建测试数据
    data = {'id': [1, 2, 3], 'name': ['A', 'B', 'C']}
    df = pd.DataFrame(data)
    df.to_excel(test_file, sheet_name='Sheet1', index=False)
    
    print(f"创建测试文件: {test_file}")
    
    # 测试describe_table
    result = excel_describe_table(test_file, 'Sheet1')
    print(f"测试结果: {result}")
    
    # 清理
    import os
    os.remove(test_file)
    os.rmdir(temp_dir)
    
    if result.get('success') == True:
        print("✅ 测试通过")
        return True
    else:
        print("❌ 测试失败")
        return False

if __name__ == "__main__":
    success = simple_test()
    print(f"最终结果: {'成功' if success else '失败'}")
    sys.exit(0 if success else 1)