#!/usr/bin/env python3
"""
直接测试Excel服务器是否能正常启动和响应MCP协议
"""

import asyncio
import json
import subprocess
import sys
from pathlib import Path

async def test_server_startup():
    """测试服务器启动和基本MCP协议"""

    server_path = Path(__file__).parent / "server.py"
    print(f"🚀 测试服务器: {server_path}")

    # 启动服务器
    process = await asyncio.create_subprocess_exec(
        sys.executable, str(server_path),
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE
    )

    try:
        # 发送初始化请求
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

        request_str = json.dumps(init_request) + '\n'
        process.stdin.write(request_str.encode())
        await process.stdin.drain()

        # 读取响应
        line = await asyncio.wait_for(process.stdout.readline(), timeout=10.0)
        response = json.loads(line.decode().strip())

        print(f"✅ 初始化响应: {json.dumps(response, ensure_ascii=False, indent=2)}")

        # 列出工具
        tools_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list",
            "params": {}
        }

        request_str = json.dumps(tools_request) + '\n'
        process.stdin.write(request_str.encode())
        await process.stdin.drain()

        # 读取工具列表响应
        line = await asyncio.wait_for(process.stdout.readline(), timeout=10.0)
        response = json.loads(line.decode().strip())

        print(f"✅ 工具列表响应: {json.dumps(response, ensure_ascii=False, indent=2)}")

        return True

    except Exception as e:
        print(f"❌ 测试失败: {e}")
        # 读取错误输出
        stderr_data = await process.stderr.read()
        if stderr_data:
            print(f"服务器错误: {stderr_data.decode()}")
        return False

    finally:
        process.terminate()
        await process.wait()

if __name__ == "__main__":
    asyncio.run(test_server_startup())
