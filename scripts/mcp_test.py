#!/usr/bin/env python3
"""
MCP 工具真实验证脚本
通过 MCP 协议直接调用工具，模拟真实客户端行为
"""
import asyncio, json, sys, os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

async def call_tool(tool_name: str, arguments: dict = None) -> dict:
    """调用单个MCP工具，返回结果"""
    from mcp.client.stdio import stdio_client, StdioServerParameters
    from mcp import ClientSession

    server_params = StdioServerParameters(
        command="python3",
        args=["-m", "src.server"],
        cwd=os.path.join(os.path.dirname(__file__), '..'),
        env={"PYTHONPATH": os.path.join(os.path.dirname(__file__), '..')}
    )

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            result = await session.call_tool(tool_name, arguments or {})
            # 解析结果
            if result.content and hasattr(result.content[0], 'text'):
                return json.loads(result.content[0].text)
            return {"error": "empty response"}

async def run_scenario(name: str, steps: list) -> list:
    """运行一个测试场景（多步操作）"""
    results = []
    for step in steps:
        tool = step["tool"]
        args = step.get("args", {})
        try:
            result = await call_tool(tool, args)
            success = result.get("success", False)
            results.append({
                "step": step.get("desc", tool),
                "tool": tool,
                "success": success,
                "result": {k: v for k, v in result.items() if k != "data"}  # 精简输出
            })
            if not success and step.get("required", True):
                results.append({"error": f"步骤失败: {step.get('desc', tool)}", "detail": result})
                break
        except Exception as e:
            results.append({"step": step.get("desc", tool), "tool": tool, "success": False, "error": str(e)})
            break
    return results

# ========== 预设游戏场景 ==========

SCENARIOS = {
    "基础CRUD": [
        {"desc": "创建文件", "tool": "excel_create_file", "args": {"file_path": "/tmp/mcp_test_basic.xlsx"}},
        {"desc": "创建工作表", "tool": "excel_create_sheet", "args": {"file_path": "/tmp/mcp_test_basic.xlsx", "sheet_name": "Skills"}},
        {"desc": "更新表头", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_basic.xlsx", "range": "Skills!A1:D1", "data": [["技能ID", "名称", "CD", "伤害"]]}},
        {"desc": "插入数据", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_basic.xlsx", "range": "Skills!A2:D4", "data": [[1, "火球术", 5, 1.5], [2, "冰冻术", 8, 1.2], [3, "治疗术", 10, 0]]}},
        {"desc": "读取数据", "tool": "excel_get_range", "args": {"file_path": "/tmp/mcp_test_basic.xlsx", "range": "Skills!A1:D4"}},
        {"desc": "SQL查询", "tool": "excel_query", "args": {"file_path": "/tmp/mcp_test_basic.xlsx", "query": "SELECT * FROM Skills WHERE CD < 8"}},
        {"desc": "删除行", "tool": "excel_delete_rows", "args": {"file_path": "/tmp/mcp_test_basic.xlsx", "sheet_name": "Skills", "start_row": 4, "count": 1}},
        {"desc": "列出工作表", "tool": "excel_list_sheets", "args": {"file_path": "/tmp/mcp_test_basic.xlsx"}},
    ],
    "双行表头": [
        {"desc": "创建装备表", "tool": "excel_create_file", "args": {"file_path": "/tmp/mcp_test_header.xlsx"}},
        {"desc": "中文描述行", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_header.xlsx", "range": "Sheet!A1:E1", "data": [["装备ID", "装备名称", "攻击力", "防御力", "稀有度"]]}},
        {"desc": "字段名行", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_header.xlsx", "range": "Sheet!A2:E2", "data": [["id", "name", "atk", "def", "rarity"]]}},
        {"desc": "装备数据", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_header.xlsx", "range": "Sheet!A3:E6", "data": [[1, "烈焰剑", 100, 0, 5], [2, "冰霜甲", 0, 200, 4], [3, "暗影匕首", 80, 0, 4], [4, "治疗戒指", 0, 50, 3]]}},
        {"desc": "获取表头", "tool": "excel_get_sheet_headers", "args": {"file_path": "/tmp/mcp_test_header.xlsx", "sheet_name": "Sheet"}},
        {"desc": "中文列名查询", "tool": "excel_query", "args": {"file_path": "/tmp/mcp_test_header.xlsx", "query": "SELECT 稀有度, AVG(攻击力) as avg_atk FROM Sheet GROUP BY 稀有度"}},
        {"desc": "DESCRIBE表结构", "tool": "excel_describe_table", "args": {"file_path": "/tmp/mcp_test_header.xlsx", "sheet_name": "Sheet"}},
    ],
    "搜索和对比": [
        {"desc": "创建测试文件A", "tool": "excel_create_file", "args": {"file_path": "/tmp/mcp_test_a.xlsx"}},
        {"desc": "写入数据A", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_a.xlsx", "range": "Sheet!A1:B3", "data": [["id", "value"], [1, "hello"], [2, "world"]]}},
        {"desc": "创建测试文件B", "tool": "excel_create_file", "args": {"file_path": "/tmp/mcp_test_b.xlsx"}},
        {"desc": "写入数据B", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_b.xlsx", "range": "Sheet!A1:B3", "data": [["id", "value"], [1, "hello"], [2, "changed"]]}},
        {"desc": "搜索内容", "tool": "excel_search", "args": {"file_path": "/tmp/mcp_test_a.xlsx", "query": "hello"}},
        {"desc": "对比文件", "tool": "excel_compare_files", "args": {"file_path1": "/tmp/mcp_test_a.xlsx", "file_path2": "/tmp/mcp_test_b.xlsx"}},
        {"desc": "文件信息", "tool": "excel_get_file_info", "args": {"file_path": "/tmp/mcp_test_a.xlsx"}},
    ],
    "格式和样式": [
        {"desc": "创建文件", "tool": "excel_create_file", "args": {"file_path": "/tmp/mcp_test_format.xlsx"}},
        {"desc": "写入数据", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_format.xlsx", "range": "Sheet!A1:B3", "data": [["标题", "内容"], ["测试1", "值1"], ["测试2", "值2"]]}},
        {"desc": "合并单元格", "tool": "excel_merge_cells", "args": {"file_path": "/tmp/mcp_test_format.xlsx", "range": "A1:B1"}},
        {"desc": "设置列宽", "tool": "excel_set_column_width", "args": {"file_path": "/tmp/mcp_test_format.xlsx", "columns": {"A": 20, "B": 30}}},
        {"desc": "设置行高", "tool": "excel_set_row_height", "args": {"file_path": "/tmp/mcp_test_format.xlsx", "rows": {1: 30}}},
    ],
    "备份和恢复": [
        {"desc": "创建文件", "tool": "excel_create_file", "args": {"file_path": "/tmp/mcp_test_backup.xlsx"}},
        {"desc": "写入数据", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_backup.xlsx", "range": "Sheet!A1:B2", "data": [["id", "val"], [1, "original"]]}},
        {"desc": "创建备份", "tool": "excel_create_backup", "args": {"file_path": "/tmp/mcp_test_backup.xlsx"}},
        {"desc": "修改数据", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_backup.xlsx", "range": "Sheet!B2", "data": [["modified"]]}},
        {"desc": "列出备份", "tool": "excel_list_backups", "args": {"file_path": "/tmp/mcp_test_backup.xlsx"}},
        {"desc": "获取历史", "tool": "excel_get_operation_history", "args": {"file_path": "/tmp/mcp_test_backup.xlsx"}},
    ],
    "SQL高级查询": [
        {"desc": "创建怪物表", "tool": "excel_create_file", "args": {"file_path": "/tmp/mcp_test_sql.xlsx"}},
        {"desc": "写入怪物数据", "tool": "excel_update_range", "args": {"file_path": "/tmp/mcp_test_sql.xlsx", "range": "Sheet!A1:F6", "data": [
            ["id", "name", "hp", "atk", "type", "level"],
            [1, "哥布林", 100, 10, "normal", 1],
            [2, "骷髅兵", 200, 20, "normal", 3],
            [3, "火龙", 1000, 100, "boss", 10],
            [4, "冰龙", 1200, 90, "boss", 10],
            [5, "史莱姆", 50, 5, "normal", 1]
        ]}},
        {"desc": "GROUP BY+聚合", "tool": "excel_query", "args": {"file_path": "/tmp/mcp_test_sql.xlsx", "query": "SELECT type, COUNT(*) as cnt, AVG(atk) as avg_atk, MAX(hp) as max_hp FROM Sheet GROUP BY type"}},
        {"desc": "ORDER BY排序", "tool": "excel_query", "args": {"file_path": "/tmp/mcp_test_sql.xlsx", "query": "SELECT name, hp, atk FROM Sheet ORDER BY hp DESC"}},
        {"desc": "WHERE条件", "tool": "excel_query", "args": {"file_path": "/tmp/mcp_test_sql.xlsx", "query": "SELECT name, atk FROM Sheet WHERE type = 'boss'"}},
        {"desc": "LIKE模糊搜索", "tool": "excel_query", "args": {"file_path": "/tmp/mcp_test_sql.xlsx", "query": "SELECT name FROM Sheet WHERE name LIKE '%龙%'"}},
        {"desc": "BETWEEN范围", "tool": "excel_query", "args": {"file_path": "/tmp/mcp_test_sql.xlsx", "query": "SELECT name, hp FROM Sheet WHERE level BETWEEN 1 AND 3"}},
        {"desc": "COUNT DISTINCT", "tool": "excel_query", "args": {"file_path": "/tmp/mcp_test_sql.xlsx", "query": "SELECT COUNT(DISTINCT type) as type_count FROM Sheet"}},
    ],
    "边界异常": [
        {"desc": "查询不存在的文件", "tool": "excel_query", "args": {"file_path": "/tmp/nonexist.xlsx", "query": "SELECT 1"}, "required": False},
        {"desc": "读取不存在的范围", "tool": "excel_get_range", "args": {"file_path": "/tmp/mcp_test_basic.xlsx", "range": "ZZ999:ZZZ999"}, "required": False},
        {"desc": "创建已存在的工作表", "tool": "excel_create_sheet", "args": {"file_path": "/tmp/mcp_test_basic.xlsx", "sheet_name": "Sheet"}, "required": False},
        {"desc": "删除不存在的工作表", "tool": "excel_delete_sheet", "args": {"file_path": "/tmp/mcp_test_basic.xlsx", "sheet_name": "NonExist"}, "required": False},
    ],
}

async def main():
    scenario_name = sys.argv[1] if len(sys.argv) > 1 else None
    
    if scenario_name and scenario_name in SCENARIOS:
        # 运行指定场景
        print(f"🧪 场景: {scenario_name}")
        print("=" * 50)
        results = await run_scenario(scenario_name, SCENARIOS[scenario_name])
        passed = sum(1 for r in results if r.get("success"))
        failed = sum(1 for r in results if not r.get("success"))
        for r in results:
            status = "✅" if r.get("success") else "❌"
            detail = r.get("result", {}).get("message", r.get("error", ""))
            print(f"  {status} {r.get('step', '?')}: {detail[:80]}")
        print(f"\n结果: {passed}通过 / {failed}失败 / 共{len(results)}步")
        return 0 if failed == 0 else 1
    
    elif scenario_name == "list":
        # 列出所有场景
        for name in SCENARIOS:
            steps = SCENARIOS[name]
            tools = [s["tool"] for s in steps]
            print(f"  {name}: {len(steps)}步, 工具: {', '.join(set(tools))}")
        return 0
    
    elif scenario_name == "all":
        # 运行所有场景
        total_pass, total_fail = 0, 0
        for name, steps in SCENARIOS.items():
            print(f"\n🧪 场景: {name}")
            print("-" * 40)
            results = await run_scenario(name, steps)
            passed = sum(1 for r in results if r.get("success"))
            failed = sum(1 for r in results if not r.get("success"))
            total_pass += passed
            total_fail += failed
            for r in results:
                status = "✅" if r.get("success") else "❌"
                detail = r.get("result", {}).get("message", r.get("error", ""))[:60]
                print(f"  {status} {r.get('step', '?')}: {detail}")
        print(f"\n{'='*50}")
        print(f"总计: {total_pass}通过 / {total_fail}失败 / 共{total_pass+total_fail}步")
        return 0 if total_fail == 0 else 1
    
    else:
        print(f"用法: python3 {sys.argv[0]} [场景名|list|all]")
        print(f"可用场景: {', '.join(SCENARIOS.keys())}")
        return 1

if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
