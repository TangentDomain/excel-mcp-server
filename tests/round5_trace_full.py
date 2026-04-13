"""完整追踪: openpyxl写入 → calamine读取 → _optimize_dtypes → 查询结果"""

import sys

sys.path.insert(0, "/root/workspace/excel-mcp-server/src")
import pandas as pd
from openpyxl import Workbook

# Step 1: 用 openpyxl 创建文件 (和 round5_test.py 完全一样)
wb = Workbook()
ws = wb.active
ws.title = "装备"
ws.append(["ID", "Name", "BaseAtk", "AtkBonus", "Price", "Rarity", "Desc"])
equip_data = [
    (1, "⚔️ 圣剑·Excalibur", 100, 25.5, 9999.99, "Legendary", "传说中的圣剑" * 5),
    (2, "🛡️ 龙鳞盾", 80, 20.0, 5999.50, "Epic", "龙鳞打造的坚固盾牌"),
    (3, "法杖·星辰", 120, 33.33, 1.5e4, "Rare", "A" * 600),
    (4, "匕首·影", 45, 10.10, 199.99, "Common", 'It\'s a "sharp" dagger'),
    (5, "🔥 炎之戒指", 30, 5.55, 899.00, "Rare", "燃烧的戒指"),
    (6, "冰霜项链", 35, 7.77, 1299.00, "Epic", "I'm cold's necklace"),
    (7, "时空之靴", 50, 12.12, 2499.99, "Legendary", ""),
    (8, "破旧木剑", 5, 0.0, 9.99, "Common", None),
    (9, "黄金战锤", 90, 22.22, 4999.00, "Epic", "Heavy!!@#$%^&*()"),
    (10, "暗影斗篷", 55, 14.44, 3499.50, "Rare", "🎮👾🚀💎 emoji test"),
    (11, "边界测试-大数", 999999999999999, 0.001, 1e10, "Common", "max int"),
    (12, "边界测试-小数", 1, 0.000001, 0.01, "Common", "min float"),  # <-- 目标行
    (13, "边界测试-负数", 70, -15.5, -100.0, "Rare", "negative values"),
    (14, "边界测试-零", 0, 0.0, 0.0, "Common", "zero row"),
    (15, "科学计数-价格", 60, 8.88e-2, 2.5e3, "Epic", "sci notation"),
]
for row in equip_data:
    ws.append(list(row))
fp = "/tmp/r5_trace_test.xlsx"
wb.save(fp)

print("=== Step 1: openpyxl 写入完成 ===")

# Step 2: calamine 读取 (项目实际使用)
print("\n=== Step 2: calamine 读取 ===")
df = pd.read_excel(fp, sheet_name="装备", engine="calamine")
print(f"  列: {list(df.columns)}")
print(f"  dtypes:\n{df.dtypes}")
print("\n  ID=12 行:")
row12 = df[df["ID"] == 12]
print(f"    {row12.to_dict('records')}")
print("\n  AtkBonus 列所有值:")
for i, v in enumerate(df["AtkBonus"].values):
    print(f"    row {i}: {v!r} (type={type(v).__name__})")

# Step 3: _optimize_dtypes 后
print("\n=== Step 3: 模拟 _optimize_dtypes ===")
from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

engine = AdvancedSQLQueryEngine()
df_opt = engine._optimize_dtypes(df.copy())
print(f"  优化后 dtypes:\n{df_opt.dtypes}")
print("\n  ID=12 行优化后:")
row12_opt = df_opt[df_opt["ID"] == 12]
print(f"    {row12_opt.to_dict('records')}")
print("\n  AtkBonus 优化后所有值:")
for i, v in enumerate(df_opt["AtkBonus"].values):
    print(f"    row {i}: {v!r}")

# Step 4: 实际查询
print("\n=== Step 4: 实际查询 ===")
from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

r = execute_advanced_sql_query(fp, "SELECT ID, Name, AtkBonus FROM 装备 WHERE ID = 12")
print(f"  result: {r['data']}")
