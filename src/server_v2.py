#!/usr/bin/env python3
"""
Excel MCP Server - 模块化重构版本

特点：
- MCP Resources: 暴露Excel文件/工作表为可订阅资源
- MCP Prompts: 预定义常用操作提示模板
- 模块化结构: 按功能分类的工具集合

技术栈：
- FastMCP: MCP服务器框架
- openpyxl: Excel文件操作
"""

import logging

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as e:
    print(f"错误: 缺少必要的依赖包: {e}")
    print("请运行: pip install fastmcp openpyxl")
    exit(1)

# 导入工具模块
from .tools import (
    register_resources,
    register_file_tools,
    register_data_tools,
    register_search_tools,
    register_format_tools,
    register_compare_tools,
    register_sql_tools,
    register_prompts,
)

# ==================== 日志配置 ====================
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# ==================== FastMCP 服务器 ====================
mcp = FastMCP(
    name="excel-mcp",
    instructions=r"""🎮 游戏开发Excel配置表管理专家

## 🔥 核心原则：SQL优先
**优先使用 `excel_query`** - 所有数据查询分析任务
- 复杂条件筛选 ✅ WHERE, LIKE, IN, BETWEEN
- 聚合统计分析 ✅ COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
- 排序限制 ✅ ORDER BY, LIMIT, OFFSET

## 📊 工具选择决策树
```
需要数据分析/查询？ → excel_query (SQL引擎)
需要定位单元格？   → excel_search (返回row/column)
需要数据修改？     → excel_update_range
需要格式调整？     → excel_format_cells
```

## ✅ SQL已支持功能 (27项)
基础查询: SELECT, DISTINCT, 别名(AS)
条件筛选: WHERE, =/>/</<, LIKE, IN, BETWEEN, AND/OR
聚合统计: COUNT(*), COUNT(col), SUM, AVG, MAX, MIN, GROUP BY, HAVING
排序限制: ORDER BY, LIMIT, OFFSET

## ❌ SQL不支持功能
子查询, CTE(WITH), JOIN, UNION, 窗口函数, CASE WHEN, INSERT/UPDATE/DELETE

## ⚠️ 重要原则
- 1-based索引: 第1行=1, 第1列=1
- 范围格式: 必须包含工作表名 "技能表!A1:Z100"
- 默认覆盖: update_range默认覆盖，需保留数据用insert_mode=True

## 🎮 游戏配置表示例
技能: SELECT 技能类型, AVG(伤害), COUNT(*) FROM 技能表 GROUP BY 技能类型
装备: SELECT 品质, AVG(价格) FROM 装备表 GROUP BY 品质 ORDER BY AVG(价格) DESC

## ⚡ 常用流程
1. excel_list_sheets - 列出工作表
2. excel_get_headers - 查看表头
3. excel_query - SQL查询
4. excel_update_range - 数据更新
5. excel_format_cells - 格式美化
""",
    debug=True,
    log_level="DEBUG"
)


# ==================== 注册 MCP 组件 ====================
def setup_server() -> FastMCP:
    """设置并注册所有MCP组件"""
    logger.info("正在注册 MCP Resources...")
    register_resources(mcp)
    
    logger.info("正在注册 MCP Prompts...")
    register_prompts(mcp)
    
    logger.info("正在注册文件操作工具...")
    register_file_tools(mcp)
    
    logger.info("正在注册数据操作工具...")
    register_data_tools(mcp)
    
    logger.info("正在注册搜索工具...")
    register_search_tools(mcp)
    
    logger.info("正在注册格式工具...")
    register_format_tools(mcp)
    
    logger.info("正在注册对比工具...")
    register_compare_tools(mcp)
    
    logger.info("正在注册SQL工具...")
    register_sql_tools(mcp)
    
    logger.info("MCP 服务器组件注册完成!")
    return mcp


# 初始化服务器
setup_server()


# ==================== 主程序 ====================
if __name__ == "__main__":
    mcp.run()
