import sys
import os
sys.path.insert(0, 'src')

import tempfile
import pandas as pd
from excel_mcp_server_fastmcp.server import (
    excel_query,
    excel_insert_rows,
    excel_describe_table
)

# 创建测试Excel文件
def create_test_excel():
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        temp_file = f.name
    
    # 创建测试数据
    data = {
        '角色名称': ['战士', '法师', '牧师', '刺客'],
        '职业': ['近战', '远程', '治疗', '近战'],
        '等级': [10, 8, 12, 9],
        '生命值': [500, 300, 400, 350],
        '技能': ['斩击', '火球', '治疗术', '背刺']
    }
    
    df = pd.DataFrame(data)
    df.to_excel(temp_file, sheet_name='角色', index=False)
    
    # 第二个工作表：技能
    skill_data = {
        '技能名称': ['斩击', '火球', '治疗术', '背刺', '闪电'],
        '职业限制': ['近战', '远程', '治疗', '近战', '远程'],
        '伤害': [100, 80, 0, 90, 70],
        '冷却时间': [5, 3, 8, 6, 4]
    }
    
    skill_df = pd.DataFrame(skill_data)
    with pd.ExcelWriter(temp_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        skill_df.to_excel(writer, sheet_name='技能', index=False)
    
    return temp_file

def test_req029_bugs():
    """测试REQ-029的两个bug"""
    test_file = create_test_excel()
    print(f"测试文件创建: {test_file}")
    
    try:
        # 测试Bug 1: JOIN后SQL表别名不生效
        print("\n=== 测试Bug 1: JOIN表别名引用 ===")
        result1 = excel_query(
            file_path=test_file,
            query_expression="SELECT r.角色名称 FROM 角色 r JOIN 技能 s ON r.职业 = s.职业限制",
            include_headers=True
        )
        print(f"JOIN别名查询结果: {result1}")
        
        # 测试Bug 2: streaming写入后describe_table崩溃
        print("\n=== 测试Bug 2: streaming写入后describe_table ===")
        
        # 先插入空行来测试
        result2 = excel_insert_rows(
            file_path=test_file,
            sheet_name="角色",
            row_index=10,
            count=1,
            streaming=True
        )
        print(f"插入空行结果: {result2}")
        
        # 然后调用describe_table
        result3 = excel_describe_table(
            file_path=test_file,
            sheet_name="角色"
        )
        print(f"describe_table结果: {result3}")
        
        # 检查结果
        bug1_fixed = ("r.角色名称" in str(result1) or "战士" in str(result1)) and len(result1.get("data", [])) > 0
        bug2_fixed = result3.get("success") is not None and "error" not in str(result3).lower()
        
        print(f"\n=== 测试结果 ===")
        print(f"Bug 1 (JOIN别名) 修复状态: {'✅ 通过' if bug1_fixed else '❌ 失败'}")
        print(f"Bug 2 (describe_table崩溃) 修复状态: {'✅ 通过' if bug2_fixed else '❌ 失败'}")
        
        return bug1_fixed and bug2_fixed
        
    except Exception as e:
        print(f"❌ 测试过程中出现异常: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        # 清理测试文件
        if os.path.exists(test_file):
            os.unlink(test_file)
            print(f"测试文件已清理: {test_file}")

if __name__ == "__main__":
    result = test_req029_bugs()
    print(f"\nREQ-029 最终验证结果: {'✅ 全部通过' if result else '❌ 仍有问题'}")