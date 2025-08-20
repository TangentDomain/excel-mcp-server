#!/usr/bin/env python3
"""
Excel MCP Server - 虚拟环境测试运行器

使用虚拟环境运行单元测试
"""

import subprocess
import sys
import os
from pathlib import Path

def main():
    """主函数 - 使用虚拟环境运行测试"""

    print("🧪 Excel MCP Server - 虚拟环境测试运行器")
    print("=" * 60)

    # 获取虚拟环境路径
    venv_path = Path('.venv')
    if not venv_path.exists():
        print("❌ 未找到虚拟环境 (.venv)")
        print("请先创建虚拟环境:")
        print("  python -m venv .venv")
        print("  source .venv/bin/activate  # macOS/Linux")
        print("  .venv\\Scripts\\activate     # Windows")
        return False

    # 确定Python路径
    if sys.platform.startswith('win'):
        python_path = venv_path / 'Scripts' / 'python.exe'
        pip_path = venv_path / 'Scripts' / 'pip.exe'
    else:
        python_path = venv_path / 'bin' / 'python'
        pip_path = venv_path / 'bin' / 'pip'

    if not python_path.exists():
        print(f"❌ 虚拟环境Python未找到: {python_path}")
        return False

    print(f"🐍 使用虚拟环境Python: {python_path}")

    # 检查是否安装了pytest
    try:
        result = subprocess.run([str(python_path), '-m', 'pytest', '--version'],
                               capture_output=True, text=True)
        if result.returncode != 0:
            print("📦 正在安装pytest...")
            subprocess.run([str(pip_path), 'install', 'pytest', 'pytest-cov'], check=True)
    except subprocess.CalledProcessError:
        print("❌ 无法安装pytest")
        return False

    # 检查测试文件
    tests_dir = Path('tests')
    if not tests_dir.exists():
        print("❌ 未找到tests目录")
        return False

    test_files = list(tests_dir.glob('test_*.py'))
    print(f"📋 找到 {len(test_files)} 个测试文件:")
    for test_file in test_files:
        print(f"  ✅ {test_file.name}")

    # 运行测试
    print(f"\n🚀 开始运行测试...")
    print("-" * 60)

    cmd = [
        str(python_path), '-m', 'pytest',
        'tests/',
        '-v',
        '--tb=short',
        '--durations=10',
    ]

    # 检查是否有pytest-cov
    try:
        subprocess.run([str(python_path), '-m', 'pytest_cov', '--version'],
                      capture_output=True, check=True)
        cmd.extend(['--cov=src', '--cov-report=term-missing'])
        print("📈 启用覆盖率分析")
    except subprocess.CalledProcessError:
        print("ℹ️  跳过覆盖率分析（pytest-cov未安装）")

    # 执行测试
    result = subprocess.run(cmd)

    print("\n" + "=" * 60)
    if result.returncode == 0:
        print("🎉 所有测试通过!")
        return True
    else:
        print("❌ 部分测试失败")
        return False

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
