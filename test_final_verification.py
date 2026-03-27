#!/usr/bin/env python3
"""
REQ-029 最终验证脚本
专注于验证修复效果
"""

import os
import sys
import tempfile
import pandas as pd
from pathlib import Path

# 添加项目路径到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src"))

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
from excel_mcp_server_fastmcp.server import excel_describe_table
import logging

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_join_alias():
    """测试JOIN表别名修复"""
    logger.info("测试JOIN表别名修复...")
    
    temp_dir = tempfile.mkdtemp(prefix="join_test_")
    test_file = os.path.join(temp_dir, "join.xlsx")
    
    try:
        # 创建测试数据
        with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
            # 玩家表
            pd.DataFrame({
                'player_id': [1, 2, 3],
                'name': ['Alice', 'Bob', 'Charlie']
            }).to_excel(writer, sheet_name='players', index=False)
            
            # 技能表
            pd.DataFrame({
                'skill_id': [101, 102, 103],
                'name': ['Fire', 'Ice', 'Lightning'],
                'player_id': [1, 2, 3]
            }).to_excel(writer, sheet_name='skills', index=False)
        
        # 执行JOIN查询
        query_engine = AdvancedSQLQueryEngine()
        query = """
        SELECT r.name as player_name, s.name as skill_name
        FROM players r
        JOIN skills s ON r.player_id = s.player_id
        """
        
        result = query_engine.execute_sql_query(test_file, query)
        logger.info(f"JOIN查询结果: {result}")
        
        if result.get('success') == True or "SQL查询成功执行" in result.get('message', ''):
            logger.info("✅ JOIN表别名修复成功")
            return True
        else:
            logger.error(f"❌ JOIN查询失败: {result.get('message')}")
            return False
            
    except Exception as e:
        logger.error(f"❌ JOIN测试异常: {e}")
        return False
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)
        os.rmdir(temp_dir)

def test_describe_table_fix():
    """测试describe_table修复"""
    logger.info("测试describe_table修复...")
    
    temp_dir = tempfile.mkdtemp(prefix="describe_test_")
    test_file = os.path.join(temp_dir, "describe.xlsx")
    
    try:
        # 创建一个有大量数据的Excel文件
        data = {
            'id': list(range(1, 1005)),  # 1004行数据
            'name': [f"Item{i}" for i in range(1, 1005)],
            'value': [float(i * 1.1) for i in range(1, 1005)]
        }
        
        with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
            pd.DataFrame(data).to_excel(writer, sheet_name='Sheet1', index=False)
        
        logger.info("创建1004行数据")
        
        # 测试describe_table
        result = excel_describe_table(test_file, 'Sheet1')
        logger.info(f"describe_table结果: {result}")
        
        if result.get('success') == True:
            data_info = result.get('data', {})
            row_count = data_info.get('row_count', 0)
            logger.info(f"返回行数: {row_count}")
            
            # 期望1004行数据（我们创建了1004行）
            if row_count == 1004:
                logger.info("✅ describe_table修复成功")
                return True
            else:
                logger.error(f"❌ 行数不正确，期望1004，实际{row_count}")
                return False
        else:
            logger.error(f"❌ describe_table失败: {result.get('message')}")
            return False
            
    except Exception as e:
        logger.error(f"❌ describe_table测试异常: {e}")
        return False
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)
        os.rmdir(temp_dir)

def main():
    logger.info("开始REQ-029最终验证...")
    
    # 测试JOIN修复
    join_ok = test_join_alias()
    logger.info(f"JOIN修复结果: {'✅ 通过' if join_ok else '❌ 失败'}")
    
    # 测试describe_table修复
    describe_ok = test_describe_table_fix()
    logger.info(f"describe_table修复结果: {'✅ 通过' if describe_ok else '❌ 失败'}")
    
    if join_ok and describe_ok:
        logger.info("🎉 所有修复验证通过！")
        return 0
    else:
        logger.error("💥 部分修复验证失败！")
        return 1

if __name__ == "__main__":
    sys.exit(main())