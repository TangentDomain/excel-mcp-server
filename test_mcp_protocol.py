#!/usr/bin/env python3
"""
MCP协议测试脚本 - 测试Excel MCP服务器
通过stdio模式与MCP服务器通信
"""

import json
import subprocess
import sys
from pathlib import Path

def send_mcp_request(server_process, request):
    """发送MCP请求并获取响应"""
    try:
        # 发送请求
        request_json = json.dumps(request) + '\n'
        server_process.stdin.write(request_json.encode())
        server_process.stdin.flush()

        # 读取响应
        response_line = server_process.stdout.readline()
        if response_line:
            return json.loads(response_line.decode().strip())
        return None
    except Exception as e:
        print(f"❌ MCP通信错误: {e}")
        return None

def test_excel_mcp_via_protocol():
    """通过MCP协议测试Excel服务器"""
    print("🧪 Excel MCP 协议测试")
    print("=" * 40)

    # 启动MCP服务器
    server_cmd = [
        str(Path("D:/mcp/excel-mcp-server-fastmcp/venv/Scripts/python.exe")),
        "D:/mcp/excel-mcp-server-fastmcp/server.py"
    ]

    try:
        server_process = subprocess.Popen(
            server_cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=False,
            cwd="D:/mcp/excel-mcp-server-fastmcp"
        )

        print("🚀 MCP服务器已启动")

        # 1. 初始化连接
        print("\n📡 测试1: 初始化MCP连接")
        init_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {
                    "name": "test-client",
                    "version": "1.0.0"
                }
            }
        }

        response = send_mcp_request(server_process, init_request)
        if response and response.get('result'):
            print("  ✅ MCP连接初始化成功")
            print(f"    🔧 服务器信息: {response['result'].get('serverInfo', {}).get('name', 'unknown')}")
        else:
            print("  ❌ MCP连接初始化失败")
            return

        # 2. 获取工具列表
        print("\n📋 测试2: 获取可用工具列表")
        tools_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list"
        }

        response = send_mcp_request(server_process, tools_request)
        if response and response.get('result'):
            tools = response['result'].get('tools', [])
            print(f"  ✅ 找到 {len(tools)} 个可用工具:")
            for tool in tools:
                print(f"    🔧 {tool['name']}")
        else:
            print("  ❌ 获取工具列表失败")

        # 3. 测试excel_list_sheets工具
        print("\n📊 测试3: excel_list_sheets工具")
        if Path("D:/mcp/excel-mcp-server-fastmcp/TrSkill.xlsx").exists():
            list_sheets_request = {
                "jsonrpc": "2.0",
                "id": 3,
                "method": "tools/call",
                "params": {
                    "name": "excel_list_sheets",
                    "arguments": {
                        "file_path": "D:/mcp/excel-mcp-server-fastmcp/TrSkill.xlsx"
                    }
                }
            }

            response = send_mcp_request(server_process, list_sheets_request)
            if response and response.get('result'):
                result = json.loads(response['result']['content'][0]['text'])
                if result['success']:
                    print(f"  ✅ 成功获取 {result['total_sheets']} 个工作表")
                    for sheet in result['sheets'][:3]:  # 显示前3个
                        active = "🎯" if sheet['is_active'] else "📄"
                        print(f"    {active} {sheet['name']}")
                else:
                    print(f"  ❌ 获取失败: {result['error']}")
            else:
                print("  ❌ MCP调用失败")
        else:
            print("  ⚠️  TrSkill.xlsx文件不存在，跳过测试")

        # 4. 测试行列访问功能
        print("\n🔢 测试4: 行列访问功能")
        if Path("D:/mcp/excel-mcp-server-fastmcp/TrSkill.xlsx").exists():
            row_access_request = {
                "jsonrpc": "2.0",
                "id": 4,
                "method": "tools/call",
                "params": {
                    "name": "excel_get_range",
                    "arguments": {
                        "file_path": "D:/mcp/excel-mcp-server-fastmcp/TrSkill.xlsx",
                        "range_expression": "1:1"
                    }
                }
            }

            response = send_mcp_request(server_process, row_access_request)
            if response and response.get('result'):
                result = json.loads(response['result']['content'][0]['text'])
                if result['success']:
                    print(f"  ✅ 第1行访问成功，类型: {result['range_type']}")
                    print(f"    📊 维度: {result['dimensions']['rows']}x{result['dimensions']['columns']}")
                else:
                    print(f"  ❌ 访问失败: {result['error']}")
            else:
                print("  ❌ MCP调用失败")

        print("\n🎉 MCP协议测试完成")

    except Exception as e:
        print(f"❌ 测试过程错误: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # 清理服务器进程
        if 'server_process' in locals():
            server_process.terminate()
            server_process.wait()

if __name__ == "__main__":
    test_excel_mcp_via_protocol()
