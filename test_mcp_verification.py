import json
import tempfile
import os
from src.excel_mcp_server_fastmcp.server import excel_describe_table, excel_query
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

def test_mcp_verification():
    """MCP真实验证"""
    
    # 创建测试文件
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
        
        print(f"📁 创建测试文件: {f.name}")
        
        # 测试1: describe_table
        print("\n🔍 测试1: describe_table")
        result1 = excel_describe_table(f.name)
        print(f"describe_table结果: {json.dumps(result1, indent=2, ensure_ascii=False)}")
        
        # 测试2: JOIN查询
        print("\n🔍 测试2: JOIN查询")
        engine = AdvancedSQLQueryEngine()
        query = """
        SELECT r.名称, s.名称 
        FROM 角色表 r 
        JOIN 技能表 s ON r.职业 = s.职业限制
        """
        result2 = engine.execute_sql_query(f.name, query)
        print(f"JOIN查询结果: {json.dumps(result2, indent=2, ensure_ascii=False)}")
        
        # 测试3: streaming写入测试
        print("\n🔍 测试3: streaming写入后describe_table")
        # 先插入一些数据
        new_data = pd.DataFrame({
            '角色ID': [4, 5],
            '名称': ['牧师', '刺客'],
            '职业': ['治疗', '潜行']
        })
        with pd.ExcelWriter(f.name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            new_data.to_excel(writer, sheet_name='角色表', startrow=4, index=False, header=False)
        
        # 再调用describe_table
        result3 = excel_describe_table(f.name, '角色表')
        print(f"streaming后describe_table: {json.dumps(result3, indent=2, ensure_ascii=False)}")
        
        # 清理
        os.unlink(f.name)
        print("\n✅ 所有测试完成！")

if __name__ == "__main__":
    import pandas as pd
    test_mcp_verification()