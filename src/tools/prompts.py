# MCP Prompts 模板模块
"""预定义的提示模板，方便用户快速执行常用操作"""

from typing import List, Dict, Any


def register_prompts(mcp) -> None:
    """注册 MCP Prompts - 预定义常用操作模板"""
    
    @mcp.prompt()
    def compare_versions() -> str:
        """版本对比提示模板"""
        return """请帮我比较两个版本的Excel配置表差异：

        ## 输入
        - 旧版本文件: {old_file}
        - 新版本文件: {new_file}
        - 工作表名称: {sheet_name}
        - ID列（用于跟踪对象）: {id_column}

        ## 输出要求
        1. 列出新增的配置项（原表中没有的）
        2. 列出被删除的配置项（新表中没有的）
        3. 列出被修改的配置项（数值或属性发生变化的）
        4. 对于数值变化，计算变化百分比
        5. 给出总体评估和建议

        ## 建议使用的工具
        - excel_compare_sheets: 对比两个工作表
        - excel_check_duplicate_ids: 验证ID唯一性"""

    @mcp.prompt()
    def skill_balance_analysis() -> str:
        """技能平衡分析提示模板"""
        return """请帮我分析游戏技能配置表的数值平衡：

        ## 输入
        - 技能配置文件: {file_path}
        - 技能工作表: {sheet_name}
        - 技能类型列: {type_column}
        - 伤害/数值列: {damage_column}

        ## 分析要求
        1. 按技能类型统计：
           - 平均伤害值
           - 最高/最低伤害值
           - 伤害分布情况
        2. 找出异常值（过高或过低的技能）
        3. 识别可能需要调整的技能
        4. 给出具体的数值调整建议

        ## 建议使用的工具
        - excel_query: SQL聚合分析
        - excel_search: 定位特定技能
        - excel_update_range: 批量调整数值"""

    @mcp.prompt()
    def equipment_quality_check() -> str:
        """装备品质检查提示模板"""
        return """请帮我检查装备配置表的品质分布：

        ## 输入
        - 装备配置文件: {file_path}
        - 装备工作表: {sheet_name}
        - 品质列: {quality_column}
        - 数值列: {value_column}

        ## 检查要求
        1. 统计各品质装备的数量和占比
        2. 分析各品质装备的数值分布
        3. 检查是否存在异常数据（数值与品质不匹配）
        4. 验证装备ID是否有重复

        ## 建议使用的工具
        - excel_query: 品质统计
        - excel_check_duplicate_ids: ID重复检查
        - excel_format_cells: 标记异常数据"""

    @mcp.prompt()
    def data_import_workflow() -> str:
        """数据导入工作流提示模板"""
        return """请帮我完成数据导入任务：

        ## 输入
        - 源文件: {source_file}
        - 目标Excel: {target_file}
        - 目标工作表: {target_sheet}

        ## 工作流步骤
        1. 首先查看源文件的结构（表头和数据格式）
        2. 检查目标文件是否存在，如不存在则创建
        3. 确认目标工作表，如不存在则创建
        4. 导入数据（处理编码问题）
        5. 验证导入结果
        6. 美化表格格式

        ## 建议使用的工具
        - excel_get_headers: 查看源文件表头
        - excel_create_file/create_sheet: 创建目标文件
        - excel_import_from_csv: 导入数据
        - excel_format_cells: 美化格式"""

    @mcp.prompt()
    def batch_update_workflow() -> str:
        """批量更新工作流提示模板"""
        return """请帮我完成批量数据更新任务：

        ## 输入
        - 文件: {file_path}
        - 工作表: {sheet_name}
        - 更新条件: {condition}
        - 更新内容: {update_value}

        ## 工作流步骤
        1. 先查询当前数据，了解数据范围
        2. 使用搜索工具定位需要更新的单元格
        3. 读取现有数据
        4. 执行批量更新
        5. 验证更新结果

        ## 示例场景
        - "将所有传说品质装备的攻击力提升20%"
        - "将火系技能的冷却时间减少30%"

        ## 建议使用的工具
        - excel_find_last_row: 确定数据范围
        - excel_search: 定位目标单元格
        - excel_get_range: 读取现有数据
        - excel_update_range: 执行更新"""
