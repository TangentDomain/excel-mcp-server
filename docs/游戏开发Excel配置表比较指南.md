# 🎮 游戏开发Excel配置表比较使用指南

## 📋 功能概览
针对游戏开发场景优化的Excel配置表比较功能，默认启用结构化数据比较模式，提供直观的配置变化展示。

## 🎯 典型使用场景

### 1. 装备配置表比较
```python
# 比较两个版本的装备配置表
result = excel_compare_sheets(
    "config_v1.xlsx", "装备表",
    "config_v2.xlsx", "装备表"
    # 默认参数已经优化为游戏开发场景
    # structured_comparison=True,  # 默认启用
    # header_row=1,               # 第一行为表头
    # id_column=1,                # 第一列为装备ID
    # show_numeric_changes=True,   # 显示数值变化
    # game_friendly_format=True    # 游戏友好格式
)
```

### 2. 技能配置表比较（使用技能名作为ID）
```python
result = excel_compare_sheets(
    "skills_old.xlsx", "技能表",
    "skills_new.xlsx", "技能表",
    id_column="技能名"  # 使用技能名列作为唯一标识
)
```

### 3. 怪物配置表比较（表头在第2行）
```python
result = excel_compare_sheets(
    "monsters_v1.xlsx", "怪物表",
    "monsters_v2.xlsx", "怪物表",
    header_row=2,  # 第2行为表头
    id_column="怪物ID"
)
```

## 📊 输出格式示例

### 数值属性变化
- 🔺 攻击力: 100.0 → 120.0 (+20.0, +20.0%)
- 🔻 防御力: 80.0 → 70.0 (-10.0, -12.5%)
- 🔺 血量: 1000 → 1200 (+200, +20.0%)

### 非数值属性变化
- 🔄 名称: '火焰剑' → '冰霜剑'
- 🔄 品质: '稀有' → '传说'
- 🔄 类型: '单手剑' → '双手剑'

### 行级变化
- ➕ 新增装备: 雷霆法杖 (ID: 1005)
- ➖ 删除装备: 破损木剑 (ID: 1001)
- 🔄 修改装备: 火焰剑 (ID: 1003) - 2处属性变化

## 🛠️ 高级配置

### 数值敏感度调整
```python
# 不区分大小写的比较（适合名称字段）
result = excel_compare_sheets(
    "config1.xlsx", "道具表",
    "config2.xlsx", "道具表",
    case_sensitive=False
)

# 忽略空单元格
result = excel_compare_sheets(
    "config1.xlsx", "任务表",
    "config2.xlsx", "任务表",
    ignore_empty_cells=True
)
```

### 传统单元格模式（特殊需求）
```python
# 如果需要传统的单元格级比较
result = excel_compare_sheets(
    "raw_data1.xlsx", "Sheet1",
    "raw_data2.xlsx", "Sheet1",
    structured_comparison=False,  # 关闭结构化比较
    game_friendly_format=False   # 使用标准格式
)
```

## 🎯 最佳实践

### 1. 配置表结构建议
- **第一行或第二行作为表头**，包含清晰的列名
- **第一列作为唯一ID**，如装备ID、技能ID、怪物ID等
- **使用描述性的列名**，如"攻击力"而不是"ATK"

### 2. 命名规范
- 数值类型字段使用中文：攻击力、防御力、血量、魔法值
- 枚举类型字段：品质、类型、分类、等级
- 标识字段：ID、名称、编号

### 3. 版本管理工作流
```python
# 1. 策划修改配置表后的验证
before_result = excel_compare_sheets("config_prod.xlsx", "装备表", "config_dev.xlsx", "装备表")

# 2. 生成变更报告给程序和测试
print(f"发现 {before_result['data']['total_differences']} 处配置变化")
print(f"新增装备: {before_result['data']['added_rows']} 个")
print(f"修改装备: {before_result['data']['modified_rows']} 个")
```

## 💡 游戏开发专用功能
1. **智能数值分析**: 自动识别数值字段，显示变化量和百分比
2. **表情符号标识**: 🔺增加、🔻减少、🔄修改，一目了然
3. **游戏术语识别**: 自动识别攻击力、防御力等游戏常用字段
4. **配置表友好**: 默认参数针对游戏配置表优化
5. **版本对比**: 支持不同版本配置表的快速对比分析
