#!/usr/bin/env python3
"""
Excel MCP Server 测试运行器

使用uv运行完整的测试套件
"""

import subprocess
import sys
from pathlib import Path


def run_command(command: str, description: str):
    """运行命令并处理结果"""
    print(f"\n🔄 {description}")
    print(f"命令: {command}")

    try:
        result = subprocess.run(
            command,
            shell=True,
            check=True,
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent
        )
        print(f"✅ {description} 成功")
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} 失败")
        if e.stdout:
            print("标准输出:", e.stdout)
        if e.stderr:
            print("标准错误:", e.stderr)
        return False


def main():
    """主函数"""
    print("🚀 开始运行Excel MCP Server完整测试套件")

    tests = [
        ("uv run python tests/test_runner.py", "运行基础功能测试"),
        ("uv run pytest tests/test_parsers.py -v", "运行解析器单元测试"),
        ("uv run pytest tests/test_validators.py -v", "运行验证器单元测试"),
    ]

    passed = 0
    total = len(tests)

    for command, description in tests:
        if run_command(command, description):
            passed += 1

    print(f"\n📊 测试总结:")
    print(f"✅ 通过: {passed}/{total}")

    if passed == total:
        print("🎉 所有测试都通过！")
        print("💡 现在可以使用以下命令启动MCP服务器：")
        print("   uv run python src/excel_mcp/server_new.py")
        return 0
    else:
        print("😞 部分测试失败，请检查上面的错误信息")
        return 1


if __name__ == "__main__":
    sys.exit(main())
