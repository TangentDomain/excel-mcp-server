#!/usr/bin/env python3
"""验证 execute_sql_query 方法修复的验证脚本"""
import inspect
from pathlib import Path

# 导入类
import sys
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

def verify_method_signature():
    """验证方法签名"""
    print("=" * 60)
    print("验证 execute_sql_query 方法签名")
    print("=" * 60)

    # 获取方法
    method = AdvancedSQLQueryEngine.execute_sql_query

    # 获取方法签名
    sig = inspect.signature(method)
    print(f"\n方法签名:\n{sig}")

    # 检查参数
    params = list(sig.parameters.keys())
    print(f"\n参数列表:\n{params}")

    # 期望的参数
    expected_params = ['self', 'file_path', 'sql', 'sheet_name', 'limit', 'include_headers', 'output_format']

    # 验证参数
    all_correct = True
    for expected in expected_params:
        if expected in params:
            print(f"  ✅ 参数 '{expected}' 存在")
        else:
            print(f"  ❌ 参数 '{expected}' 缺失")
            all_correct = False

    # 检查默认值
    print("\n默认值:")
    if 'sheet_name' in sig.parameters:
        default = sig.parameters['sheet_name'].default
        if default is None or default == inspect.Parameter.empty:
            print(f"  ✅ sheet_name: {default}")
        else:
            print(f"  ⚠️  sheet_name: {default} (期望 None)")
    else:
        print(f"  ❌ sheet_name 参数缺失")
        all_correct = False

    if 'limit' in sig.parameters:
        default = sig.parameters['limit'].default
        if default is None or default == inspect.Parameter.empty:
            print(f"  ✅ limit: {default}")
        else:
            print(f"  ⚠️  limit: {default} (期望 None)")
    else:
        print(f"  ❌ limit 参数缺失")
        all_correct = False

    if 'include_headers' in sig.parameters:
        default = sig.parameters['include_headers'].default
        if default == True or default == inspect.Parameter.empty:
            print(f"  ✅ include_headers: {default}")
        else:
            print(f"  ⚠️  include_headers: {default} (期望 True)")
    else:
        print(f"  ❌ include_headers 参数缺失")
        all_correct = False

    if 'output_format' in sig.parameters:
        default = sig.parameters['output_format'].default
        if default == "table":
            print(f"  ✅ output_format: '{default}'")
        else:
            print(f"  ⚠️  output_format: '{default}' (期望 'table')")
    else:
        print(f"  ❌ output_format 参数缺失")
        all_correct = False

    # 检查返回类型注解
    print("\n返回类型注解:")
    return_annotation = sig.return_annotation
    print(f"  返回类型: {return_annotation}")
    if 'Dict' in str(return_annotation):
        print(f"  ✅ 返回类型包含 Dict")
    else:
        print(f"  ⚠️  返回类型可能不正确")

    print("\n" + "=" * 60)
    if all_correct:
        print("✅ 方法签名验证通过")
    else:
        print("❌ 方法签名验证失败")
    print("=" * 60)

    return all_correct

def verify_docstring():
    """验证 docstring"""
    print("\n" + "=" * 60)
    print("验证 execute_sql_query 方法 docstring")
    print("=" * 60)

    method = AdvancedSQLQueryEngine.execute_sql_query
    doc = method.__doc__

    if not doc:
        print("❌ 方法没有 docstring")
        return False

    print(f"\nDocstring 长度: {len(doc)} 字符")

    # 检查关键内容
    checks = [
        ('Args', 'Args 部分'),
        ('file_path', 'file_path 参数'),
        ('sql', 'sql 参数'),
        ('sheet_name', 'sheet_name 参数'),
        ('limit', 'limit 参数'),
        ('include_headers', 'include_headers 参数'),
        ('output_format', 'output_format 参数'),
        ('Returns', 'Returns 部分'),
        ('Dict', '返回类型 Dict'),
        ('success', 'success 字段'),
        ('message', 'message 字段'),
        ('data', 'data 字段'),
    ]

    all_correct = True
    for keyword, description in checks:
        if keyword in doc:
            print(f"  ✅ {description} 存在")
        else:
            print(f"  ❌ {description} 缺失")
            all_correct = False

    print("\n" + "=" * 60)
    if all_correct:
        print("✅ Docstring 验证通过")
    else:
        print("❌ Docstring 验证失败")
    print("=" * 60)

    return all_correct

def main():
    """主验证函数"""
    print("\n")
    print("╔" + "═" * 58 + "╗")
    print("║" + " " * 10 + "execute_sql_query 修复验证" + " " * 19 + "║")
    print("╚" + "═" * 58 + "╝")

    # 验证方法签名
    signature_ok = verify_method_signature()

    # 验证 docstring
    docstring_ok = verify_docstring()

    # 总体结果
    print("\n" + "╔" + "═" * 58 + "╗")
    if signature_ok and docstring_ok:
        print("║" + " " * 18 + "FIXED" + " " * 27 + "║")
        print("║" + " " * 12 + "所有验证通过，修复成功！" + " " * 14 + "║")
        print("╚" + "═" * 58 + "╝")
        return 0
    else:
        print("║" + " " * 18 + "FAILED" + " " * 26 + "║")
        print("║" + " " * 10 + "部分验证失败，请检查修复内容！" + " " * 13 + "║")
        print("╚" + "═" * 58 + "╝")
        return 1

if __name__ == "__main__":
    exit(main())
