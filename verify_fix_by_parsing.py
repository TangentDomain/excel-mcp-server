#!/usr/bin/env python3
"""验证 execute_sql_query 方法修复的验证脚本（基于文件内容解析）"""
import re
from pathlib import Path

def extract_method_signature_and_docstring(file_path, method_name):
    """从文件中提取方法的签名和 docstring"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 查找方法定义的开始
    pattern = rf'def {method_name}\('
    match = re.search(pattern, content)

    if not match:
        return None, None

    # 提取方法签名（多行支持）
    signature_start = match.start()
    # 找到方法签名的结束位置（: 之前，处理多行和类型注解中的冒号）
    signature_end = signature_start
    paren_count = 1  # 初始为 1，因为我们已经匹配了第一个 '('
    i = match.end()  # 从 '(' 之后开始

    while i < len(content) and paren_count > 0:
        char = content[i]
        if char == '(':
            paren_count += 1
        elif char == ')':
            paren_count -= 1
            if paren_count == 0:
                # 找到匹配的 ')'
                signature_end = i + 1
                break
        i += 1

    if signature_end == signature_start:
        return None, None

    # 查找返回类型注解和冒号
    # 从 signature_end 开始查找 -> 和 :
    remaining = content[signature_end:]
    match_return = re.search(r'\s*(?:->[^:]+)?\s*:', remaining)
    if match_return:
        signature_end += match_return.end()

    signature = content[signature_start:signature_end]

    # 提取 docstring
    docstring_start = signature_end
    # 跳过空行和缩进
    while docstring_start < len(content) and content[docstring_start] in ' \t\n':
        docstring_start += 1

    # 检查 docstring 是否开始
    if not content[docstring_start:docstring_start+3] == '"""':
        return signature, None

    # 找到 docstring 结束
    docstring_end = content.find('"""', docstring_start + 3)
    if docstring_end == -1:
        return signature, None

    docstring = content[docstring_start:docstring_end + 3]

    return signature, docstring

def verify_method_signature(signature):
    """验证方法签名"""
    print("=" * 60)
    print("验证 execute_sql_query 方法签名")
    print("=" * 60)

    if not signature:
        print("❌ 未找到方法签名")
        return False

    print(f"\n方法签名:\n{signature}")

    # 检查必需的参数
    required_params = ['self', 'file_path', 'sql']
    optional_params = ['sheet_name', 'limit', 'include_headers', 'output_format']

    all_correct = True

    # 检查必需参数
    for param in required_params:
        if param in signature:
            print(f"  ✅ 参数 '{param}' 存在")
        else:
            print(f"  ❌ 参数 '{param}' 缺失")
            all_correct = False

    # 检查可选参数
    for param in optional_params:
        if param in signature:
            print(f"  ✅ 参数 '{param}' 存在")
        else:
            print(f"  ❌ 参数 '{param}' 缺失")
            all_correct = False

    # 检查默认值
    print("\n默认值:")
    if 'sheet_name: Optional[str] = None' in signature or 'sheet_name=None' in signature:
        print(f"  ✅ sheet_name: Optional[str] = None")
    else:
        print(f"  ⚠️  sheet_name 默认值可能不正确")

    if 'limit: Optional[int] = None' in signature or 'limit=None' in signature:
        print(f"  ✅ limit: Optional[int] = None")
    else:
        print(f"  ⚠️  limit 默认值可能不正确")

    if 'include_headers: bool = True' in signature or 'include_headers=True' in signature:
        print(f"  ✅ include_headers: bool = True")
    else:
        print(f"  ⚠️  include_headers 默认值可能不正确")

    if 'output_format: str = "table"' in signature or 'output_format="table"' in signature:
        print(f"  ✅ output_format: str = 'table'")
    else:
        print(f"  ⚠️  output_format 默认值可能不正确")

    # 检查返回类型注解
    print("\n返回类型注解:")
    if '-> Dict[str, Any]' in signature or '-> Dict[' in signature:
        print(f"  ✅ 返回类型包含 Dict[str, Any]")
    else:
        print(f"  ⚠️  返回类型可能不正确")

    print("\n" + "=" * 60)
    if all_correct:
        print("✅ 方法签名验证通过")
    else:
        print("❌ 方法签名验证失败")
    print("=" * 60)

    return all_correct

def verify_docstring(docstring):
    """验证 docstring"""
    print("\n" + "=" * 60)
    print("验证 execute_sql_query 方法 docstring")
    print("=" * 60)

    if not docstring:
        print("❌ 方法没有 docstring")
        return False

    print(f"\nDocstring 前 200 字符:\n{docstring[:200]}...")

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
        if keyword in docstring:
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

    file_path = 'src/excel_mcp_server_fastmcp/api/advanced_sql_query.py'

    # 提取方法签名和 docstring
    signature, docstring = extract_method_signature_and_docstring(file_path, 'execute_sql_query')

    if not signature:
        print("❌ 无法提取方法签名")
        return 1

    # 验证方法签名
    signature_ok = verify_method_signature(signature)

    # 验证 docstring
    docstring_ok = verify_docstring(docstring)

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
