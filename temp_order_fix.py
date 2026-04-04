# Fix for ORDER BY mixed data types
import re

# Read the file
with open('src/excel_mcp_server_fastmcp/api/advanced_sql_query.py', 'r') as f:
    content = f.read()

# Find the _apply_order_by method and add mixed type handling
old_sort_logic = '''        if sort_columns:
            return df.sort_values(by=sort_columns, ascending=ascending)

        return df'''

new_sort_logic = '''        if sort_columns:
            # Handle mixed data types in ORDER BY columns
            for col in sort_columns:
                if col in df.columns:
                    # Convert to string type for mixed data to avoid sorting errors
                    # This ensures consistent ordering of NULLs, numbers, and strings
                    df[f'_temp_sort_{col}'] = df[col].astype(str)
                    # Replace the sort column with the temp string version
                    sort_columns = [f'_temp_sort_{c}' if c == col else c for c in sort_columns]
            
            return df.sort_values(by=sort_columns, ascending=ascending)
            
            # Clean up temporary columns
            for col in sort_columns:
                if col.startswith('_temp_sort_'):
                    df.drop(columns=[col], inplace=True)

        return df'''

# Replace the sorting logic
content = content.replace(old_sort_logic, new_sort_logic)

# Write back
with open('src/excel_mcp_server_fastmcp/api/advanced_sql_query.py', 'w') as f:
    f.write(content)

print("Fixed ORDER BY mixed data types handling")
