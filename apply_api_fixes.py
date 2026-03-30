#!/usr/bin/env python3
"""
自动应用 API 问题的修复

修复内容：
1. apply_formula 缺少 formula 参数时的处理
2. write_data_to_excel 数据格式不匹配时的处理
"""

import sys
from pathlib import Path

def fix_excel_set_formula():
    """修复 excel_set_formula 函数"""
    print("🔧 修复 excel_set_formula...")

    file_path = Path("src/excel_mcp_server_fastmcp/server.py")

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 查找函数定义
    if 'def excel_set_formula(' not in content:
        print("❌ 未找到 excel_set_formula 函数")
        return False

    # 检查是否已经修复
    if '# 参数验证：formula 不能为空' in content:
        print("✅ excel_set_formula 已经修复")
        return True

    # 查找插入点（在函数定义后，第一个 _path_err 之前）
    old_code = '''def excel_set_formula(
    file_path: str,
    sheet_name: str,
    cell_address: str,
    formula: str
) -> Dict[str, Any]:
    """在单元格写入Excel公式。"""
    _path_err = _validate_path(file_path)'''

    new_code = '''def excel_set_formula(
    file_path: str,
    sheet_name: str,
    cell_address: str,
    formula: str
) -> Dict[str, Any]:
    """在单元格写入Excel公式。"""
    # 参数验证：formula 不能为空
    if not formula or not formula.strip():
        return _fail(
            '公式不能为空，请提供有效的Excel公式（如 "=A1+B1"）',
            meta={"error_code": "MISSING_FORMULA"}
        )

    _path_err = _validate_path(file_path)'''

    if old_code not in content:
        print("❌ 无法找到匹配的代码片段")
        return False

    content = content.replace(old_code, new_code, 1)

    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)

    print("✅ excel_set_formula 修复完成")
    return True


def fix_update_range():
    """修复 update_range 函数"""
    print("🔧 修复 update_range...")

    file_path = Path("src/excel_mcp_server_fastmcp/api/excel_operations.py")

    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 检查是否已经修复
    if any('# 新增：数据格式验证' in line for line in lines):
        print("✅ update_range 已经修复")
        return True

    # 查找函数定义位置
    func_start = None
    for i, line in enumerate(lines):
        if 'def update_range(' in line and 'file_path: str' in lines[i+1]:
            func_start = i
            break

    if func_start is None:
        print("❌ 未找到 update_range 函数")
        return False

    # 查找函数 docstring 结束位置
    indent = 0
    docstring_end = None
    for i in range(func_start, min(func_start + 50, len(lines))):
        if lines[i].strip().startswith('"""'):
            if docstring_end is not None:
                docstring_end = i + 1
                break
            else:
                docstring_end = i + 1  # 找到开始

    if docstring_end is None or docstring_end >= len(lines):
        print("❌ 无法找到 docstring 结束位置")
        return False

    # 获取缩进
    for line in lines[docstring_end:]:
        if line.strip():
            indent = len(line) - len(line.lstrip())
            break

    # 插入验证代码
    validation_code = f"""{' ' * indent}    # 新增：数据格式验证
{' ' * indent}    if not data:
{' ' * indent}        return cls._format_error_result("数据不能为空")
{' ' * indent}
{' ' * indent}    if not isinstance(data, list):
{' ' * indent}        return cls._format_error_result(
{' ' * indent}            f"数据格式错误：data 应该是二维数组 [[row1], [row2], ...]，"
{' ' * indent}            f"实际收到类型: {{type(data).__name__}}"
{' ' * indent}        )
{' ' * indent}
{' ' * indent}    # 验证每一行是否都是列表
{' ' * indent}    for i, row in enumerate(data):
{' ' * indent}        if not isinstance(row, list):
{' ' * indent}            return cls._format_error_result(
{' ' * indent}                f"数据格式错误：第 {{i+1}} 行应该是列表，实际收到类型: {{type(row).__name__}}。"
{' ' * indent}                "data 应该是二维数组 [[row1], [row2], ...]"
{' ' * indent}            )
{' ' * indent}
"""

    # 插入代码
    lines.insert(docstring_end, validation_code)

    with open(file_path, 'w', encoding='utf-8') as f:
        f.writelines(lines)

    print("✅ update_range 修复完成")
    return True


def verify_fixes():
    """验证修复是否成功"""
    print("\n🔍 验证修复...")

    # 验证 excel_set_formula
    with open("src/excel_mcp_server_fastmcp/server.py", 'r', encoding='utf-8') as f:
        content = f.read()

    if '# 参数验证：formula 不能为空' in content:
        print("✅ excel_set_formula 验证通过")
    else:
        print("❌ excel_set_formula 验证失败")
        return False

    # 验证 update_range
    with open("src/excel_mcp_server_fastmcp/api/excel_operations.py", 'r', encoding='utf-8') as f:
        content = f.read()

    if '# 新增：数据格式验证' in content:
        print("✅ update_range 验证通过")
    else:
        print("❌ update_range 验证失败")
        return False

    return True


def main():
    """主函数"""
    print("=" * 60)
    print("Excel MCP API 问题自动修复")
    print("=" * 60)
    print()

    success = True

    # 修复 excel_set_formula
    if not fix_excel_set_formula():
        success = False

    # 修复 update_range
    if not fix_update_range():
        success = False

    # 验证修复
    if not verify_fixes():
        success = False

    print()
    print("=" * 60)
    if success:
        print("✅ 所有修复完成")
        print("=" * 60)
        print()
        print("📝 下一步：")
        print("1. 运行测试验证修复效果：")
        print("   python3 test_api_issues_v2.py")
        print()
        print("2. 检查代码修改：")
        print("   git diff src/excel_mcp_server_fastmcp/server.py")
        print("   git diff src/excel_mcp_server_fastmcp/api/excel_operations.py")
        print()
        print("3. 如果需要回滚：")
        print("   git checkout src/excel_mcp_server_fastmcp/server.py")
        print("   git checkout src/excel_mcp_server_fastmcp/api/excel_operations.py")
    else:
        print("❌ 修复失败，请检查错误信息")
        print("=" * 60)
        sys.exit(1)


if __name__ == "__main__":
    main()
