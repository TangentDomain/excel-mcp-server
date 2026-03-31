import pandas as pd
import os

# Create a test Excel file
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'Score': [85, 90, 78]
}

df = pd.DataFrame(data)
df.to_excel('test_data.xlsx', index=False)
print("Created test_data.xlsx")