#!/usr/bin/env python3
"""
测试通过MCP协议调用Excel服务器的新功能
"""

import asyncio
import json
import subprocess
import sys
from pathlib import Path
import signal
import time

class MCPClient:
    def __init__(self, server_path: str):
        self.server_path = server_path
        self.process = None

    async def start_server(self):
        """启动MCP服务器"""
        print(f"🚀 启动MCP服务器: {self.server_path}")

        # 启动服务器进程
        self.process = await asyncio.create_subprocess_exec(
            sys.executable, self.server_path,
            stdin=asyncio.subprocess.PIPE,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE
        )

        # 初始化MCP连接
        init_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {
                    "tools": {}
                },
                "clientInfo": {
                    "name": "test-client",
                    "version": "1.0.0"
                }
            }
        }

        await self.send_request(init_request)

        # 接收初始化响应
        response = await self.receive_response()
        print(f"✅ 服务器初始化响应: {response}")

        return True

    async def send_request(self, request: dict):
        """发送MCP请求"""
        if not self.process:
            raise RuntimeError("服务器未启动")

        request_str = json.dumps(request) + '\n'
        self.process.stdin.write(request_str.encode())
        await self.process.stdin.drain()

    async def receive_response(self):
        """接收MCP响应"""
        if not self.process:
            raise RuntimeError("服务器未启动")

        line = await self.process.stdout.readline()
        if not line:
            return None

        try:
            # 尝试多种编码解码
            text = None
            for encoding in ['utf-8', 'gbk', 'cp1252', 'latin1']:
                try:
                    text = line.decode(encoding).strip()
                    break
                except UnicodeDecodeError:
                    continue

            if text is None:
                print(f"❌ 无法解码响应: {line}")
                return None

            # 如果不是JSON格式，可能是服务器输出的调试信息
            if not text.startswith('{'):
                print(f"📝 服务器输出: {text}")
                return await self.receive_response()  # 继续读取下一行

            return json.loads(text)

        except json.JSONDecodeError as e:
            print(f"❌ JSON解析错误: {e}")
            print(f"原始数据: {line}")
            return None

    async def call_tool(self, name: str, arguments: dict):
        """调用MCP工具"""
        request = {
            "jsonrpc": "2.0",
            "id": int(time.time() * 1000),  # 使用时间戳作为ID
            "method": "tools/call",
            "params": {
                "name": name,
                "arguments": arguments
            }
        }

        print(f"📤 调用工具: {name}")
        print(f"   参数: {json.dumps(arguments, ensure_ascii=False, indent=2)}")

        await self.send_request(request)
        response = await self.receive_response()

        if response and "result" in response:
            print(f"✅ 工具调用成功")
            return response["result"]
        elif response and "error" in response:
            print(f"❌ 工具调用失败: {response['error']}")
            return None
        else:
            print(f"⚠️  未知响应: {response}")
            return None

    async def list_tools(self):
        """列出可用工具"""
        request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list",
            "params": {}
        }

        await self.send_request(request)
        response = await self.receive_response()

        if response and "result" in response:
            tools = response["result"]["tools"]
            print(f"📋 可用工具列表 ({len(tools)}个):")
            for tool in tools:
                print(f"   • {tool['name']}: {tool.get('description', '无描述')}")
            return tools
        else:
            print(f"❌ 获取工具列表失败: {response}")
            return []

    async def stop_server(self):
        """停止MCP服务器"""
        if self.process:
            print("🛑 停止MCP服务器...")
            self.process.terminate()
            try:
                await asyncio.wait_for(self.process.wait(), timeout=5.0)
            except asyncio.TimeoutError:
                print("⚠️  强制终止服务器进程")
                self.process.kill()
                await self.process.wait()

async def test_excel_mcp():
    """测试Excel MCP服务器的新功能"""

    # 确定服务器路径
    server_path = Path(__file__).parent / "server.py"
    if not server_path.exists():
        print(f"❌ 服务器文件不存在: {server_path}")
        return False

    # 确定测试文件路径
    test_file = Path(__file__).parent / "TrSkill.xlsx"
    if not test_file.exists():
        print(f"❌ 测试文件不存在: {test_file}")
        return False

    client = MCPClient(str(server_path))

    try:
        # 启动服务器
        await client.start_server()

        # 列出可用工具
        tools = await client.list_tools()

        # 测试1: 列出Excel工作表
        print("\n" + "="*50)
        print("🧪 测试1: 列出Excel工作表")
        print("="*50)

        result1 = await client.call_tool("excel_list_sheets", {
            "file_path": str(test_file)
        })

        if result1:
            print(f"📊 工作表列表: {result1}")

        # 测试2: 获取特定范围数据（行列访问）
        print("\n" + "="*50)
        print("🧪 测试2: 获取特定范围数据（行列访问）")
        print("="*50)

        result2 = await client.call_tool("excel_get_range", {
            "file_path": str(test_file),
            "range_ref": "A1:C5",
            "sheet_name": "Sheet1"
        })

        if result2:
            print(f"📈 范围数据: {json.dumps(result2, ensure_ascii=False, indent=2)}")

        # 测试3: 获取整行数据
        print("\n" + "="*50)
        print("🧪 测试3: 获取整行数据")
        print("="*50)

        result3 = await client.call_tool("excel_get_range", {
            "file_path": str(test_file),
            "range_ref": "3:3",  # 第3行
            "sheet_name": "Sheet1"
        })

        if result3:
            print(f"📈 行数据: {json.dumps(result3, ensure_ascii=False, indent=2)}")

        # 测试4: 获取整列数据
        print("\n" + "="*50)
        print("🧪 测试4: 获取整列数据")
        print("="*50)

        result4 = await client.call_tool("excel_get_range", {
            "file_path": str(test_file),
            "range_ref": "B:B",  # B列
            "sheet_name": "Sheet1"
        })

        if result4:
            print(f"📈 列数据 (前10行): {json.dumps(result4['data'][:10] if result4.get('data') else None, ensure_ascii=False, indent=2)}")
            print(f"📊 范围类型: {result4.get('range_type', 'unknown')}")

        print("\n" + "="*50)
        print("🎉 MCP协议测试完成！")
        print("="*50)

        return True

    except Exception as e:
        print(f"❌ 测试过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        await client.stop_server()

if __name__ == "__main__":
    try:
        # 运行测试
        success = asyncio.run(test_excel_mcp())

        if success:
            print("\n✅ 所有测试通过！")
            sys.exit(0)
        else:
            print("\n❌ 测试失败！")
            sys.exit(1)

    except KeyboardInterrupt:
        print("\n⚠️  用户中断测试")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 测试异常: {e}")
        sys.exit(1)
