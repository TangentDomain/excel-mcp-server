"""验证 float32 转换是否导致极小浮点数丢失"""

import numpy as np
import pandas as pd

# 模拟装备表的 AtkBonus 列
values = [
    25.5,
    20.0,
    33.33,
    10.10,
    5.55,
    7.77,
    12.12,
    0.0,
    22.22,
    14.44,
    0.001,
    0.000001,
    -15.5,
    0.0,
    8.88e-2,
]

print("=== 原始值 ===")
for v in values:
    print(f"  {v!r}")

print("\n=== 转 float32 后 ===")
s = pd.Series(values, dtype="float64")
s32 = s.astype("float32").astype("float64")  # 转回 float64 显示
for orig, conv in zip(values, s32):
    status = "✅" if abs(orig - conv) < 1e-6 else f"❌ 差异={abs(orig - conv)}"
    print(f"  {orig!r} -> {conv!r} {status}")

# 特别关注 0.000001
val = 0.000001
f32 = np.float32(val)
f64_back = float(f32)
print("\n=== 重点: 0.000001 ===")
print(f"  原始:     {val!r}")
print(f"  float32:  {f32!r}")
print(f"  转回64:   {f64_back!r}")
print(f"  是否相等: {val == f64_back}")

# 也测试更大的列（含极大整数）
print("\n=== 含极大整数的混合列 ===")
mixed_vals = [100, 80, 120, 45, 30, 35, 50, 5, 90, 55, 999999999999999, 1, 70, 0, 60]
mixed_bonus = [
    25.5,
    20.0,
    33.33,
    10.10,
    5.55,
    7.77,
    12.12,
    0.0,
    22.22,
    14.44,
    0.001,
    0.000001,
    -15.5,
    0.0,
    8.88e-2,
]
df = pd.DataFrame({"BaseAtk": mixed_vals, "AtkBonus": mixed_bonus})
print(f"  转换前 AtkBonus:\n{df['AtkBonus'].values}")
df["AtkBonus"] = df["AtkBonus"].astype("float32")
print(f"  转换后 AtkBonus (float32):\n{df['AtkBonus'].values}")
# 再转回 float64 看实际存储
df["AtkBonus"] = df["AtkBonus"].astype("float64")
print(f"  显示转回后:\n{df['AtkBonus'].values}")
