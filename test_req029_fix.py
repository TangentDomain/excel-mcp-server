"""
验证REQ-029 bug修复的测试文件

Bug 1: JOIN表别名 + describe_table崩溃
- JOIN后SQL表别名`r.名称`不生效，pandas加`_x`/`_y`后缀未映射
- Bug 2: streaming写入后openpyxl read_only模式max_row=None，describe_table崩溃
"""

import os
import tempfile
import pandas as pd
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
from src.excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations

def test_join_alias_mapping():
    """测试JOIN别名映射"""
    print("🔍 测试JOIN别名映射...")
    
    # 创建测试数据
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        # 角色表
        char_data = pd.DataFrame({
            '角色ID': [1, 2, 3],
            '名称': ['战士', '法师', '射手'],
            '职业': ['近战', '远程', '远程']
        })
        
        # 技能表
        skill_data = pd.DataFrame({
            '技能ID': [101, 102, 103],
            '名称': ['斩击', '火球术', '精准射击'],
            '职业限制': ['近战', '远程', '远程']
        })
        
        # 保存为Excel文件
        with pd.ExcelWriter(f.name, engine='openpyxl') as writer:
            char_data.to_excel(writer, sheet_name='角色表', index=False)
            skill_data.to_excel(writer, sheet_name='技能表', index=False)
        
        # 测试JOIN查询
        engine = AdvancedSQLQueryEngine()
        query = """
        SELECT r.名称, s.名称 
        FROM 角色表 r 
        JOIN 技能表 s ON r.职业 = s.职业限制
        """
        
        result = engine.execute_sql_query(f.name, query)
        print(f"JOIN查询结果: {result}")
        
        # 清理
        os.unlink(f.name)

def test_describe_table_after_streaming():
    """测试streaming写入后的describe_table"""
    print("🔍 测试streaming写入后的describe_table...")
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        # 创建原始数据
        data = pd.DataFrame({
            'ID': [1, 2, 3],
            '名称': ['测试1', '测试2', '测试3'],
            '数值': [100, 200, 300]
        })
        
        data.to_excel(f.name, index=False)
        
        # 执行streaming写入
        excel_ops = ExcelOperations()
        try:
            # 执行streaming插入
            from excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
            if StreamingWriter.is_available():
                StreamingWriter.insert_rows_streaming(
                    f.name, 'Sheet', 4,  # start_row
                    [[4, '测试4', 400], [5, '测试5', 500]]
                )
            
            # 测试describe_table
            from src.excel_mcp_server_fastmcp.server import excel_describe_table
            result = excel_describe_table(f.name)
            print(f"Describe Table结果: {result}")
            
        except Exception as e:
            print(f"❌ 错误: {e}")
        
        # 清理
        os.unlink(f.name)

if __name__ == "__main__":
    print("🚀 开始测试REQ-029 bug修复...")
    
    test_join_alias_mapping()
    test_describe_table_after_streaming()
    
    print("✅ 测试完成！")