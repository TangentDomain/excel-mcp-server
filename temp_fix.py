# Fix for _build_total_row method
import re

# Read the file
with open('src/excel_mcp_server_fastmcp/api/advanced_sql_query.py', 'r') as f:
    content = f.read()

# Find and replace the _build_total_row method
old_method = '''    def _build_total_row(self, result_df: pd.DataFrame, group_by_columns: List[str] = []) -> Optional[List]:
        """构建GROUP BY聚合结果的TOTAL汇总行

        Args:
            result_df (pd.DataFrame): 结果DataFrame
            group_by_columns (List[str]): GROUP BY列名列表，跳过这些列的求和

        Returns:
            Optional[List]: TOTAL汇总行，如果没有数值列则返回None
        """
        if result_df.empty or len(result_df) <= 1:
            return None
        total_row = [''] * len(result_df.columns)
        has_numeric = False
        for i, col in enumerate(result_df.columns):
            if col in group_by_columns:
                continue
            series = pd.to_numeric(result_df[col], errors='coerce')
            if series.notna().sum() > len(result_df) * 0.5:
                total_row[i] = self._serialize_value(series.sum())
                has_numeric = True
        if has_numeric:
            total_row[0] = 'TOTAL'
        return total_row if has_numeric else None'''

new_method = '''    def _build_total_row(self, result_df: pd.DataFrame, group_by_columns: List[str] = []) -> Optional[List]:
        """构建GROUP BY聚合结果的TOTAL汇总行

        Args:
            result_df (pd.DataFrame): 结果DataFrame
            group_by_columns (List[str]): GROUP BY列名列表，跳过这些列的求和

        Returns:
            Optional[List]: TOTAL汇总行，如果没有数值列则返回None
        """
        if result_df.empty or len(result_df) <= 1:
            return None
        
        # Create a copy to avoid modifying original
        total_row = [''] * len(result_df.columns)
        
        # First pass: ensure GROUP BY columns are preserved correctly
        for i, col in enumerate(result_df.columns):
            if col in group_by_columns:
                # Copy the first value from each group for GROUP BY columns
                total_row[i] = result_df.iloc[0, i]
        
        # Second pass: calculate sums for non-GROUP BY numeric columns
        has_numeric = False
        for i, col in enumerate(result_df.columns):
            if col in group_by_columns:
                continue
                
            # Try to convert to numeric
            series = pd.to_numeric(result_df[col], errors='coerce')
            
            # Check if this column is numeric enough
            if series.notna().sum() > 0:
                # Calculate sum and serialize
                col_sum = series.sum()
                total_row[i] = self._serialize_value(col_sum)
                has_numeric = True
        
        # Mark as TOTAL row only if we have numeric aggregations
        if has_numeric:
            total_row[0] = 'TOTAL'
        
        return total_row if has_numeric else None'''

# Replace the method
content = content.replace(old_method, new_method)

# Write back
with open('src/excel_mcp_server_fastmcp/api/advanced_sql_query.py', 'w') as f:
    f.write(content)

print("Fixed _build_total_row method")
