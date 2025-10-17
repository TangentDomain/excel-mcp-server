#!/usr/bin/env python3
"""
Validate that all MCP tools are properly defined in server.py
"""

import ast
import sys
from pathlib import Path


def extract_mcp_tools_from_server():
    """Extract MCP tool definitions from server.py"""

    server_file = Path("src/server.py")
    if not server_file.exists():
        print("❌ src/server.py not found")
        return False

    try:
        with open(server_file, 'r', encoding='utf-8') as f:
            content = f.read()

        # Parse the AST
        tree = ast.parse(content)

        # Find MCP tool decorators and functions
        mcp_tools = []

        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef):
                # Check if function has MCP tool decorator
                if hasattr(node, 'decorator_list'):
                    for decorator in node.decorator_list:
                        if (isinstance(decorator, ast.Name) and
                            decorator.id in ['mcp_tool', 'tool']):
                            mcp_tools.append(node.name)
                        elif (isinstance(decorator, ast.Attribute) and
                              hasattr(decorator, 'attr') and
                              decorator.attr in ['tool', 'mcp_tool']):
                            mcp_tools.append(node.name)

        return mcp_tools

    except SyntaxError as e:
        print(f"❌ Syntax error in src/server.py: {e}")
        return False
    except Exception as e:
        print(f"❌ Error parsing src/server.py: {e}")
        return False


def check_expected_tools(found_tools):
    """Check if expected tools are present"""

    # Expected tools based on the project documentation
    expected_tools = {
        # File and worksheet management
        'excel_list_sheets',
        'excel_get_file_info',
        'excel_create_file',
        'excel_create_sheet',
        'excel_delete_sheet',
        'excel_rename_sheet',
        'excel_get_sheet_headers',

        # Data operations
        'excel_get_range',
        'excel_update_range',
        'excel_get_headers',
        'excel_insert_rows',
        'excel_delete_rows',
        'excel_insert_columns',
        'excel_delete_columns',
        'excel_find_last_row',

        # Search and analysis
        'excel_search',
        'excel_search_directory',
        'excel_check_duplicate_ids',
        'excel_compare_sheets',

        # Formatting and styling
        'excel_format_cells',
        'excel_merge_cells',
        'excel_unmerge_cells',
        'excel_set_borders',
        'excel_set_row_height',
        'excel_set_column_width',

        # Import/Export and conversion
        'excel_export_to_csv',
        'excel_import_from_csv',
        'excel_convert_format',
        'excel_merge_files',
    }

    found_set = set(found_tools)
    missing_tools = expected_tools - found_set
    extra_tools = found_set - expected_tools

    success = True

    if missing_tools:
        print(f"❌ Missing expected MCP tools ({len(missing_tools)}):")
        for tool in sorted(missing_tools):
            print(f"   - {tool}")
        success = False

    if extra_tools:
        print(f"⚠️  Additional MCP tools found ({len(extra_tools)}):")
        for tool in sorted(extra_tools):
            print(f"   - {tool}")

    if len(found_tools) >= len(expected_tools) * 0.8:  # At least 80% coverage
        print(f"✅ Found {len(found_tools)} MCP tools")
        return success
    else:
        print(f"❌ Too few MCP tools found: {len(found_tools)} (expected at least {int(len(expected_tools) * 0.8)})")
        return False


def validate_mcp_tools():
    """Main validation function"""

    print("🔍 Validating MCP tools in src/server.py...")

    mcp_tools = extract_mcp_tools_from_server()
    if not mcp_tools:
        return False

    print(f"📋 Found {len(mcp_tools)} MCP tool functions:")
    for tool in sorted(mcp_tools):
        print(f"   - {tool}")

    return check_expected_tools(mcp_tools)


if __name__ == "__main__":
    success = validate_mcp_tools()
    sys.exit(0 if success else 1)