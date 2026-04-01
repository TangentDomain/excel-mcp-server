import asyncio
import json
import sys
import os
sys.path.append('/root/.openclaw/workspace/excel-mcp-server')

# Import the Excel MCP server directly
from src.excel_mcp_server_fastmcp.server import ExcelMCPServer

async def test_api_issues():
    """Test the 5 API issues mentioned in .cron-focus.md"""
    
    # Create a test Excel file
    test_file = "/tmp/test_api_issues.xlsx"
    
    # Initialize server
    server = ExcelMCPServer()
    
    print("=== Testing API Issues ===")
    
    # Issue 1: read_data_from_excel parameter order
    print("\n1. Testing read_data_from_excel parameter order...")
    try:
        result = await server.handle_read_data_from_excel(
            filepath=test_file,
            sheet_name="Sheet1",
            start_cell="B2",
            end_cell="D5"
        )
        print("✅ Parameter order test passed")
    except Exception as e:
        print(f"❌ Parameter order test failed: {e}")
    
    # Issue 2: format_range missing parameters
    print("\n2. Testing format_range missing parameters...")
    try:
        result = await server.handle_format_range(
            filepath=test_file,
            sheet_name="Sheet1",
            start_cell="A1",
            # Missing bold parameter - should handle gracefully
        )
        print("✅ Missing parameters test passed")
    except Exception as e:
        print(f"❌ Missing parameters test failed: {e}")
    
    # Issue 3: apply_formula missing formula
    print("\n3. Testing apply_formula missing formula...")
    try:
        result = await server.handle_apply_formula(
            filepath=test_file,
            sheet_name="Sheet1",
            cell="A1",
            # Missing formula parameter - should handle gracefully
        )
        print("✅ Missing formula test passed")
    except Exception as e:
        print(f"❌ Missing formula test failed: {e}")
    
    # Issue 4: read_data_from_excel search logic
    print("\n4. Testing read_data_from_excel search logic...")
    try:
        result = await server.handle_read_data_from_excel(
            filepath=test_file,
            sheet_name="Sheet1",
            start_cell="A1",
            end_cell="Z10",
            search="test"  # This might conflict with other parameters
        )
        print("✅ Search logic test passed")
    except Exception as e:
        print(f"❌ Search logic test failed: {e}")
    
    # Issue 5: write_data_to_excel data format
    print("\n5. Testing write_data_to_excel data format...")
    try:
        result = await server.handle_write_data_to_excel(
            filepath=test_file,
            sheet_name="Sheet1",
            data="invalid_format",  # Should be list of lists, not string
            start_cell="A1"
        )
        print("✅ Data format test passed")
    except Exception as e:
        print(f"❌ Data format test failed: {e}")
    
    print("\n=== API Test Complete ===")

if __name__ == "__main__":
    asyncio.run(test_api_issues())