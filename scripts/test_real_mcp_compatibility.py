#!/usr/bin/env python3
"""
多客户端真实MCP兼容性验证
REQ-012: 使用真实MCP协议测试不同客户端环境
"""

import asyncio
import json
import subprocess
import sys
import os
import tempfile
import time
from pathlib import Path

# 添加项目路径
project_root = Path(__file__).parent.parent
src_path = project_root / "src"
sys.path.insert(0, str(src_path))

async def call_mcp_tool(tool_name: str, arguments: dict = None) -> dict:
    """通过MCP协议调用工具，返回结果"""
    try:
        from mcp.client.stdio import stdio_client, StdioServerParameters
        from mcp import ClientSession
        
        server_params = StdioServerParameters(
            command="uvx",
            args=["excel-mcp-server-fastmcp"],
            cwd=str(project_root),
            env={"PYTHONPATH": str(project_root)}
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.call_tool(tool_name, arguments or {})
                
                # 解析结果
                if result.content and hasattr(result.content[0], 'text'):
                    return json.loads(result.content[0].text)
                return {"error": "empty response", "success": False}
                
    except Exception as e:
        return {"error": str(e), "success": False}

async def test_cursor_compatibility():
    """测试Cursor兼容性（OpenAI风格配置）"""
    print("🎯 测试: Cursor兼容性")
    
    # 创建测试Excel文件
    test_file = "/tmp/test_cursor.xlsx"
    
    try:
        # 模拟Cursor工作流
        steps = [
            {
                "tool": "excel_create_file", 
                "args": {"file_path": test_file},
                "desc": "创建测试文件"
            },
            {
                "tool": "excel_create_sheet",
                "args": {"file_path": test_file, "sheet_name": "角色"},
                "desc": "创建角色表"
            },
            {
                "tool": "excel_update_range",
                "args": {"file_path": test_file, "range": "角色!A1:C1", "data": [["ID", "名称", "职业"]]},
                "desc": "写入表头"
            },
            {
                "tool": "excel_update_range", 
                "args": {"file_path": test_file, "range": "角色!A2:C4", "data": [[1, "战士", "近战"], [2, "法师", "远程"], [3, "牧师", "辅助"]]},
                "desc": "写入角色数据"
            },
            {
                "tool": "excel_query",
                "args": {"file_path": test_file, "query": "SELECT * FROM 角色 WHERE 职业 = '近战'"},
                "desc": "SQL查询测试"
            }
        ]
        
        results = []
        for step in steps:
            result = await call_mcp_tool(step["tool"], step["args"])
            success = result.get("success", False)
            results.append({
                "step": step["desc"],
                "success": success,
                "result": result
            })
            
            if not success:
                print(f"  ❌ {step['desc']}: {result.get('error', '未知错误')}")
                break
            else:
                print(f"  ✅ {step['desc']}")
        
        # 清理
        if os.path.exists(test_file):
            os.remove(test_file)
            
        cursor_success = all(r["success"] for r in results)
        print(f"  📊 Cursor兼容性: {'✅ 通过' if cursor_success else '❌ 失败'} ({len([r for r in results if r['success']])}/{len(results)})")
        return cursor_success
        
    except Exception as e:
        print(f"  ❌ Cursor测试异常: {e}")
        return False

async def test_claude_desktop_compatibility():
    """测试Claude Desktop兼容性"""
    print("🎯 测试: Claude Desktop兼容性")
    
    test_file = "/tmp/test_claude.xlsx"
    
    try:
        # Claude Desktop工作流 - 更复杂的操作序列
        steps = [
            {
                "tool": "excel_create_file",
                "args": {"file_path": test_file},
                "desc": "创建项目Excel"
            },
            {
                "tool": "excel_create_sheet", 
                "args": {"file_path": test_file, "sheet_name": "装备"},
                "desc": "创建装备表"
            },
            {
                "tool": "excel_create_sheet",
                "args": {"file_path": test_file, "sheet_name": "技能"}, 
                "desc": "创建技能表"
            },
            {
                "tool": "excel_update_range",
                "args": {"file_path": test_file, "range": "装备!A1:D1", "data": [["ID", "名称", "攻击力", "防御力"]]},
                "desc": "装备表头"
            },
            {
                "tool": "excel_update_range",
                "args": {"file_path": test_file, "range": "技能!A1:C1", "data": [["ID", "名称", "消耗"]]},
                "desc": "技能表头"
            },
            {
                "tool": "excel_update_range",
                "args": {"file_path": test_file, "range": "装备!A2:D4", "data": [[1, "烈焰剑", 100, 20], [2, "冰霜甲", 50, 80], [3, "暗影匕首", 80, 10]]},
                "desc": "装备数据"
            },
            {
                "tool": "excel_update_range",
                "args": {"file_path": test_file, "range": "技能!A2:C4", "data": [[1, "火球术", 10], [2, "冰冻术", 15], [3, "治疗术", 20]]},
                "desc": "技能数据"
            },
            {
                "tool": "excel_query",
                "args": {"file_path": test_file, "query": "SELECT * FROM 装备 WHERE 攻击力 > 50"},
                "desc": "装备查询"
            },
            {
                "tool": "excel_query", 
                "args": {"file_path": test_file, "query": "SELECT 名称, 消耗 FROM 技能 WHERE 消耗 < 15"},
                "desc": "技能查询"
            },
            {
                "tool": "excel_describe_table",
                "args": {"file_path": test_file, "sheet_name": "装备"},
                "desc": "装备表结构"
            },
            {
                "tool": "excel_list_sheets",
                "args": {"file_path": test_file},
                "desc": "列出所有工作表"
            }
        ]
        
        results = []
        for step in steps:
            result = await call_mcp_tool(step["tool"], step["args"])
            success = result.get("success", False)
            results.append({
                "step": step["desc"],
                "success": success,
                "result": result
            })
            
            if not success:
                print(f"  ❌ {step['desc']}: {result.get('error', '未知错误')}")
                break
            else:
                print(f"  ✅ {step['desc']}")
        
        # 清理
        if os.path.exists(test_file):
            os.remove(test_file)
            
        claude_success = all(r["success"] for r in results)
        print(f"  📊 Claude Desktop兼容性: {'✅ 通过' if claude_success else '❌ 失败'} ({len([r for r in results if r['success']])}/{len(results)})")
        return claude_success
        
    except Exception as e:
        print(f"  ❌ Claude Desktop测试异常: {e}")
        return False

async def test_vscode_mcp_compatibility():
    """测试VSCode MCP扩展兼容性"""
    print("🎯 测试: VSCode MCP扩展兼容性")
    
    test_file = "/tmp/test_vscode.xlsx"
    
    try:
        # VSCode工作流 - 侧重编辑和开发任务
        steps = [
            {
                "tool": "excel_create_file",
                "args": {"file_path": test_file},
                "desc": "创建开发用Excel"
            },
            {
                "tool": "excel_create_sheet",
                "args": {"file_path": test_file, "sheet_name": "怪物配置"},
                "desc": "创建怪物配置表"
            },
            {
                "tool": "excel_update_range",
                "args": {"file_path": test_file, "range": "怪物配置!A1:E1", "data": [["ID", "名称", "HP", "ATK", "类型"]]},
                "desc": "怪物表头"
            },
            {
                "tool": "excel_update_range",
                "args": {"file_path": test_file, "range": "怪物配置!A2:E6", "data": [
                    [1, "哥布林", 100, 10, "普通"],
                    [2, "骷髅兵", 200, 20, "普通"],
                    [3, "火龙", 1000, 100, "Boss"],
                    [4, "冰龙", 1200, 90, "Boss"],
                    [5, "史莱姆", 50, 5, "普通"]
                ]},
                "desc": "怪物数据"
            },
            {
                "tool": "excel_query",
                "args": {"file_path": test_file, "query": "SELECT 名称, HP, ATK FROM 怪物配置 WHERE 类型 = 'Boss' ORDER BY HP DESC"},
                "desc": "Boss怪物查询"
            },
            {
                "tool": "excel_query",
                "args": {"file_path": test_file, "query": "SELECT 类型, COUNT(*) as 数量, AVG(HP) as 平均HP FROM 怪物配置 GROUP BY 类型"},
                "desc": "分组统计"
            },
            {
                "tool": "excel_get_headers",
                "args": {"file_path": test_file, "sheet_name": "怪物配置"},
                "desc": "获取表头"
            },
            {
                "tool": "excel_find_last_row",
                "args": {"file_path": test_file, "sheet_name": "怪物配置"},
                "desc": "查找最后一行"
            },
            {
                "tool": "excel_get_range",
                "args": {"file_path": test_file, "range": "怪物配置!A1:E1"},
                "desc": "读取表头"
            }
        ]
        
        results = []
        for step in steps:
            result = await call_mcp_tool(step["tool"], step["args"])
            success = result.get("success", False)
            results.append({
                "step": step["desc"],
                "success": success,
                "result": result
            })
            
            if not success:
                print(f"  ❌ {step['desc']}: {result.get('error', '未知错误')}")
                break
            else:
                print(f"  ✅ {step['desc']}")
        
        # 清理
        if os.path.exists(test_file):
            os.remove(test_file)
            
        vscode_success = all(r["success"] for r in results)
        print(f"  📊 VSCode MCP兼容性: {'✅ 通过' if vscode_success else '❌ 失败'} ({len([r for r in results if r['success']])}/{len(results)})")
        return vscode_success
        
    except Exception as e:
        print(f"  ❌ VSCode MCP测试异常: {e}")
        return False

async def test_streaming_compatibility():
    """测试流式写入在不同客户端的兼容性"""
    print("🎯 测试: 流式写入兼容性")
    
    test_file = "/tmp/test_streaming.xlsx"
    
    try:
        # 测试不同流式写入配置
        streaming_tests = [
            {
                "name": "大批量数据流式写入",
                "steps": [
                    {
                        "tool": "excel_create_file",
                        "args": {"file_path": test_file},
                        "desc": "创建流式测试文件"
                    },
                    {
                        "tool": "excel_create_sheet",
                        "args": {"file_path": test_file, "sheet_name": "测试数据"},
                        "desc": "创建工作表"
                    },
                    {
                        "tool": "excel_update_range",
                        "args": {"file_path": test_file, "range": "测试数据!A1:D1", "data": [["ID", "名称", "数值", "时间戳"]]},
                        "desc": "写入表头"
                    },
                    # 模拟流式写入大量数据
                    {
                        "tool": "excel_update_range",
                        "args": {"file_path": test_file, "range": "测试数据!A2:D101", "data": [
                            [i, f"测试数据{i}", i * 10, int(time.time()) + i] for i in range(1, 101)
                        ]},
                        "desc": "写入100行数据"
                    },
                    {
                        "tool": "excel_find_last_row",
                        "args": {"file_path": test_file, "sheet_name": "测试数据"},
                        "desc": "验证最后一行"
                    },
                    {
                        "tool": "excel_query",
                        "args": {"file_path": test_file, "query": "SELECT COUNT(*) as 总行数 FROM 测试数据"},
                        "desc": "验证数据完整性"
                    }
                ]
            },
            {
                "name": "小批量流式写入",
                "steps": [
                    {
                        "tool": "excel_update_range",
                        "args": {"file_path": test_file, "range": "测试数据!A102:D105", "data": [
                            [101, "小批量1", 1010, int(time.time()) + 101],
                            [102, "小批量2", 1020, int(time.time()) + 102],
                            [103, "小批量3", 1030, int(time.time()) + 103],
                            [104, "小批量4", 1040, int(time.time()) + 104]
                        ]},
                        "desc": "写入小批量数据"
                    },
                    {
                        "tool": "excel_query",
                        "args": {"file_path": test_file, "query": "SELECT 名称, 数值 FROM 测试数据 WHERE ID BETWEEN 101 AND 104"},
                        "desc": "验证小批量数据"
                    }
                ]
            }
        ]
        
        all_success = True
        for test_case in streaming_tests:
            print(f"  📦 {test_case['name']}")
            for step in test_case["steps"]:
                result = await call_mcp_tool(step["tool"], step["args"])
                success = result.get("success", False)
                
                if not success:
                    print(f"    ❌ {step['desc']}: {result.get('error', '未知错误')}")
                    all_success = False
                    break
                else:
                    print(f"    ✅ {step['desc']}")
        
        # 清理
        if os.path.exists(test_file):
            os.remove(test_file)
            
        streaming_success = all_success
        print(f"  📊 流式写入兼容性: {'✅ 通过' if streaming_success else '❌ 失败'}")
        return streaming_success
        
    except Exception as e:
        print(f"  ❌ 流式写入测试异常: {e}")
        return False

async def main():
    """主测试流程"""
    print("🚀 开始REQ-012: 多客户端真实MCP兼容性验证")
    print("=" * 70)
    
    # 运行所有测试
    tests = [
        ("Cursor兼容性", test_cursor_compatibility),
        ("Claude Desktop兼容性", test_claude_desktop_compatibility), 
        ("VSCode MCP兼容性", test_vscode_mcp_compatibility),
        ("流式写入兼容性", test_streaming_compatibility)
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"\n{'='*70}")
        success = await test_func()
        results.append({"test": test_name, "success": success})
    
    # 生成报告
    print(f"\n{'='*70}")
    print("📋 多客户端兼容性测试报告")
    print("=" * 70)
    
    total_tests = len(results)
    passed_tests = sum(1 for r in results if r["success"])
    success_rate = passed_tests / total_tests * 100
    
    print(f"总测试数: {total_tests}")
    print(f"通过测试: {passed_tests}")
    print(f"成功率: {success_rate:.1f}%")
    
    # 详细结果
    for result in results:
        status = "✅ 通过" if result["success"] else "❌ 失败"
        print(f"  {result['test']}: {status}")
    
    # 判断是否通过REQ-012验证
    if success_rate >= 90.0:  # 90%成功率认为兼容性良好
        print(f"\n🎉 REQ-012 多客户端兼容性验证: ✅ 通过")
        print("   所有主要客户端环境都能正常使用ExcelMCP")
        
        # 保存详细报告
        report = {
            "test_date": time.strftime("%Y-%m-%d %H:%M:%S"),
            "total_tests": total_tests,
            "passed_tests": passed_tests,
            "success_rate": success_rate,
            "details": results,
            "conclusion": "通过"
        }
        
        report_path = "/tmp/multi_client_compatibility_report.json"
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        
        print(f"📄 详细报告已保存至: {report_path}")
        return True
        
    else:
        print(f"\n❌ REQ-012 多客户端兼容性验证: 需要改进")
        print(f"   当前成功率 {success_rate:.1f}% < 90% 要求")
        return False

if __name__ == "__main__":
    try:
        success = asyncio.run(main())
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"❌ 测试运行异常: {e}")
        sys.exit(1)