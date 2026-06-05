# 🎮 游戏开发使用示例

本文档提供ExcelMCP在游戏开发中的实际使用场景，包含技能系统、装备管理、怪物配置等典型游戏配置操作。

## 📁 示例文件结构

```
examples/
├── 基础操作/
│   ├── 01_创建技能表.py
│   ├── 02_装备配置管理.py
│   └── 03_怪物属性设置.py
├── 进阶操作/
│   ├── 04_跨文件JOIN查询.py
│   ├── 05_批量数据更新.py
│   └── 06_版本对比与回滚.py
└── 实战案例/
    ├── 07_技能系统完整设计.py
    └── 08_数值平衡调整.py
```

---

## 🎯 基础操作

### 01_创建技能表.py

```python
"""
创建技能配置表示例
演示如何创建和管理技能表数据
"""

# MCP配置连接
from mcp.server import Server
import asyncio

async def create_skills_sheet():
    """创建技能表示例"""
    
    # 1. 创建新的技能表
    await excel_create_worksheet(
        filepath="game_data/skills.xlsx",
        sheet_name="skills",
        headers=["skill_id", "skill_name", "damage", "cooldown", "target_type"]
    )
    
    # 2. 批量插入技能数据
    skills_data = [
        [1, "火球术", 150, 3.0, "单体"],
        [2, "冰冻术", 120, 4.0, "范围"],
        [3, "治疗术", 0, 5.0, "单体"],
        [4, "雷电术", 200, 6.0, "单体"]
    ]
    
    await excel_write_rows(
        filepath="game_data/skills.xlsx",
        sheet_name="skills",
        data=skills_data,
        start_cell="A2"
    )
    
    # 3. 设置数据验证
    await excel_set_validation(
        filepath="game_data/skills.xlsx",
        sheet_name="skills",
        range="B2:B100",
        validation_type="list",
        formula="=技能列表!A2:A20"
    )

if __name__ == "__main__":
    asyncio.run(create_skills_sheet())
```

### 02_装备配置管理.py

```python
"""
装备配置管理示例
演示如何管理装备数据和属性加成
"""

async def manage_equipment():
    """管理装备配置"""
    
    # 1. 创建装备表
    await excel_create_worksheet(
        filepath="game_data/equipment.xlsx",
        sheet_name="weapons",
        headers=["item_id", "name", "rarity", "attack", "defense", "speed_bonus"]
    )
    
    # 2. 插入装备数据
    weapons_data = [
        [1001, "龙鳞剑", "史诗", 85, 20, 15],
        [1002, "魔法法杖", "稀有", 45, 35, 25],
        [1003, "守护盾牌", "传说", 30, 80, 10]
    ]
    
    await excel_write_rows(
        filepath="game_data/equipment.xlsx",
        sheet_name="weapons",
        data=weapons_data,
        start_cell="A2"
    )
    
    # 3. 使用公式计算装备评分
    await excel_apply_formula(
        filepath="game_data/equipment.xlsx",
        sheet_name="weapons",
        cell="G2",
        formula="=SUM(C2:F2)*IF(B2=\"传说\",1.5,IF(B2=\"史诗\",1.2,1))"
    )
    
    # 4. 复制装备到不同稀有度分类
    await excel_copy_range(
        filepath="game_data/equipment.xlsx",
        sheet_name="weapons",
        source_start="A1",
        source_end="F100",
        target_start="A1",
        target_sheet="epic_weapons"
    )

if __name__ == "__main__":
    asyncio.run(manage_equipment())
```

### 03_怪物属性设置.py

```python
"""
怪物属性设置示例
演示如何配置怪物数据和AI行为
"""

async def configure_monsters():
    """配置怪物属性"""
    
    # 1. 创建怪物配置表
    await excel_create_worksheet(
        filepath="game_data/monsters.xlsx",
        sheet_name="monster_stats",
        headers=["monster_id", "name", "level", "hp", "attack", "defense", "ai_type"]
    )
    
    # 2. 插入怪物数据
    monsters_data = [
        [1, "史莱姆", 1, 100, 15, 5, "passive"],
        [2, "哥布林", 3, 150, 25, 8, "aggressive"],
        [3, "龙", 10, 1000, 150, 50, "boss"],
        [4, "蝙蝠", 2, 80, 20, 3, "flying"]
    ]
    
    await excel_write_rows(
        filepath="game_data/monsters.xlsx",
        sheet_name="monster_stats",
        data=monsters_data,
        start_cell="A2"
    )
    
    # 3. 根据等级计算属性缩放
    await excel_apply_formula(
        filepath="game_data/monsters.xlsx",
        sheet_name="monster_stats",
        cell="H2",
        formula="=C2*100+D2*0.5"
    )
    
    # 4. 设置条件格式，高亮Boss级怪物
    await excel_format_range(
        filepath="game_data/monsters.xlsx",
        sheet_name="monster_stats",
        start_cell="G2",
        end_cell="G100",
        bg_color="#ffeb3b",
        conditional_format={
            "type": "cell",
            "criteria": "equal to",
            "value": "\"boss\"",
            "format": {"bg_color": "#ff5722"}
        }
    )

if __name__ == "__main__":
    asyncio.run(configure_monsters())
```

---

## 🚀 进阶操作

### 04_跨文件JOIN查询.py

```python
"""
跨文件JOIN查询示例
演示如何关联多个配置表进行复杂查询
"""

async def cross_file_join():
    """跨文件JOIN查询"""
    
    # 1. 定义技能效果表
    skill_effects = [
        [1, "火球术", "火焰伤害", "单体", "2倍伤害"],
        [2, "冰冻术", "控制效果", "范围", "减速50%"],
        [3, "治疗术", "恢复效果", "单体", "恢复HP"]
    ]
    
    await excel_write_rows(
        filepath="game_data/skills.xlsx",
        sheet_name="skill_effects",
        data=skill_effects,
        start_cell="A1"
    )
    
    # 2. 在技能表中关联效果查询
    result = await excel_sql_query(
        query="""
        SELECT s.skill_id, s.skill_name, s.damage, e.effect_type, e.description
        FROM skills s
        JOIN skill_effects e ON s.skill_id = e.skill_id
        WHERE s.damage > 100
        ORDER BY s.damage DESC
        """,
        filepath="game_data/skills.xlsx"
    )
    
    print("高伤害技能列表：")
    for row in result['data']:
        print(f"{row[1]}: {row[2]}伤害 - {row[4]}")
    
    # 3. 查询特定角色可用的技能
    character_skills = await excel_sql_query(
        query="""
        SELECT m.name, s.skill_name, s.cooldown
        FROM monster_stats m
        LEFT JOIN skills s ON m.level >= 5 AND s.damage <= m.attack * 2
        WHERE m.level > 5
        ORDER BY m.level, s.cooldown
        """,
        filepath="game_data/monsters.xlsx"
    )

if __name__ == "__main__":
    asyncio.run(cross_file_join())
```

### 05_批量数据更新.py

```python
"""
批量数据更新示例
演示如何批量更新游戏配置数据
"""

async def batch_update_data():
    """批量数据更新"""
    
    # 1. 更新所有技能的冷却时间
    await excel_update_rows(
        filepath="game_data/skills.xlsx",
        sheet_name="skills",
        updates=[
            {"row": 2, "column": 4, "value": 2.5, "formula": None},  # 火球术
            {"row": 3, "column": 4, "value": 3.5, "formula": None},  # 冰冻术
            {"row": 4, "column": 4, "value": 4.5, "formula": None}   # 治疗术
        ]
    )
    
    # 2. 使用流式写入更新大量装备数据
    equipment_updates = []
    for i in range(1, 100):
        equipment_updates.append([1000+i, f"装备{i}", "普通", 10+i, 5+i, 2+i])
    
    await excel_write_rows(
        filepath="game_data/equipment.xlsx",
        sheet_name="weapons",
        data=equipment_updates,
        start_cell="A100",
        streaming=True  # 启用流式写入，大数据量性能更优
    )
    
    # 3. 批量更新怪物等级和属性
    await excel_update_rows(
        filepath="game_data/monsters.xlsx",
        sheet_name="monster_stats",
        updates=[
            {"row": 2, "column": 3, "value": 5, "formula": None},   # 史莱姆升级
            {"row": 3, "column": 3, "value": 8, "formula": None},   # 哥布林升级
            {"row": 4, "column": 3, "value": 15, "formula": None}   # 龙升级
        ]
    )
    
    # 4. 批量插入新怪物
    new_monsters = [
        [5, "恶魔", 12, 1200, 180, 60, "boss"],
        [6, "幽灵", 8, 200, 40, 15, "undead"],
        [7, "巨魔", 6, 300, 60, 25, "aggressive"]
    ]
    
    await excel_append_rows(
        filepath="game_data/monsters.xlsx",
        sheet_name="monster_stats",
        data=new_monsters
    )

if __name__ == "__main__":
    asyncio.run(batch_update_data())
```

### 06_版本对比与回滚.py

```python
"""
版本对比与回滚示例
演示如何管理游戏配置的版本控制
"""

async def version_management():
    """版本对比与回滚"""
    
    # 1. 创建配置备份
    await excel_copy_worksheet(
        filepath="game_data/skills.xlsx",
        source_sheet="skills",
        target_sheet="skills_backup_v1"
    )
    
    # 2. 模拟技能平衡调整
    balance_changes = [
        {"row": 2, "column": 3, "value": 130},  # 火球术伤害下调
        {"row": 3, "column": 3, "value": 110},  # 冰冻术伤害下调
        {"row": 4, "column": 4, "value": 4.5}   # 治疗术冷却增加
    ]
    
    await excel_update_rows(
        filepath="game_data/skills.xlsx",
        sheet_name="skills",
        updates=balance_changes
    )
    
    # 3. 对比版本差异
    diff_result = await excel_compare_sheets(
        filepath="game_data/skills.xlsx",
        sheet1="skills",
        sheet2="skills_backup_v1",
        compare_by="row"
    )
    
    print("版本差异：")
    for change in diff_result['changes']:
        print(f"行{change['row']} 列{change['column']}: {change['old_value']} → {change['new_value']}")
    
    # 4. 如果调整效果不好，可以回滚到备份版本
    # await excel_copy_worksheet(
    #     filepath="game_data/skills.xlsx",
    #     source_sheet="skills_backup_v1", 
    #     target_sheet="skills"
    # )

if __name__ == "__main__":
    asyncio.run(version_management())
```

---

## 🎮 实战案例

### 07_技能系统完整设计.py

```python
"""
技能系统完整设计示例
演示从零开始构建完整的技能系统
"""

async def complete_skill_system():
    """技能系统完整设计"""
    
    # 1. 创建技能表结构
    skill_headers = [
        "skill_id", "skill_name", "skill_type", "element", 
        "damage", "mana_cost", "cooldown", "target_type",
        "range", "duration", "is_passive"
    ]
    
    await excel_create_worksheet(
        filepath="game_data/skills_complete.xlsx",
        sheet_name="skills_master",
        headers=skill_headers
    )
    
    # 2. 批量插入技能数据
    all_skills = [
        # 主动技能
        [1, "火球术", "攻击", "火", 150, 50, 3.0, "单体", "中等", 0.0, False],
        [2, "冰冻术", "控制", "冰", 80, 30, 4.0, "范围", "远", 3.0, False],
        [3, "治疗术", "辅助", "光", 0, 20, 5.0, "单体", "接触", 0.0, False],
        [4, "雷电术", "攻击", "雷", 200, 80, 6.0, "单体", "远", 0.0, False],
        [5, "护盾术", "防御", "光", 0, 40, 8.0, "单体", "接触", 10.0, False],
        # 被动技能
        [6, "狂暴", "被动", "无", 0, 0, 0.0, "自身", "无", 0.0, True],
        [7, "闪避", "被动", "无", 0, 0, 0.0, "自身", "无", 0.0, True],
        [8, "经验加成", "被动", "无", 0, 0, 0.0, "自身", "无", 0.0, True]
    ]
    
    await excel_write_rows(
        filepath="game_data/skills_complete.xlsx",
        sheet_name="skills_master",
        data=all_skills,
        start_cell="A2"
    )
    
    # 3. 创建技能学习表（角色-技能关联）
    character_skill_headers = ["character_id", "character_name", "skill_id", "level_required", "unlocked"]
    character_skills_data = [
        [1, "战士", 1, 1, True],
        [1, "战士", 6, 5, True],
        [1, "战士", 7, 10, False],
        [2, "法师", 2, 1, True],
        [2, "法师", 4, 5, True],
        [2, "法师", 8, 15, False],
        [3, "牧师", 3, 1, True],
        [3, "牧师", 5, 3, True]
    ]
    
    await excel_create_worksheet(
        filepath="game_data/skills_complete.xlsx",
        sheet_name="character_skills",
        headers=character_skill_headers
    )
    
    await excel_write_rows(
        filepath="game_data/skills_complete.xlsx",
        sheet_name="character_skills",
        data=character_skills_data,
        start_cell="A2"
    )
    
    # 4. 创建技能效果表
    effect_headers = ["effect_id", "skill_id", "effect_type", "target_attribute", "value", "duration"]
    effects_data = [
        [1, 1, "damage", "hp", -150, 0.0],
        [2, 2, "control", "speed", -50, 3.0],
        [3, 3, "heal", "hp", 200, 0.0],
        [4, 4, "damage", "hp", -200, 0.0],
        [5, 5, "shield", "defense", 50, 10.0]
    ]
    
    await excel_create_worksheet(
        filepath="game_data/skills_complete.xlsx",
        sheet_name="skill_effects",
        headers=effect_headers
    )
    
    await excel_write_rows(
        filepath="game_data/skills_complete.xlsx",
        sheet_name="skill_effects",
        data=effects_data,
        start_cell="A2"
    )
    
    # 5. 生成技能组合查询
    skill_combinations = await excel_sql_query(
        query="""
        SELECT c.character_name, s.skill_name, s.damage, s.mana_cost, s.cooldown
        FROM character_skills cs
        JOIN character_master c ON cs.character_id = c.character_id
        JOIN skills_master s ON cs.skill_id = s.skill_id
        WHERE cs.unlocked = TRUE
        ORDER BY c.character_name, cs.level_required
        """,
        filepath="game_data/skills_complete.xlsx"
    )
    
    print("角色可用技能组合：")
    for row in skill_combinations['data']:
        print(f"{row[0]}: {row[1]} ({row[2]}伤害, {row[3]}法力, {row[4]}冷却)")

if __name__ == "__main__":
    asyncio.run(complete_skill_system())
```

### 08_数值平衡调整.py

```python
"""
数值平衡调整示例
演示如何使用ExcelMCP进行游戏数值平衡和平衡分析
"""

async def numerical_balance():
    """数值平衡调整"""
    
    # 1. 创建战斗平衡分析表
    balance_headers = [
        "character_id", "character_name", "total_damage", "total_defense", 
        "speed", "hp", "dps_ratio", "survivability", "balance_score"
    ]
    
    await excel_create_worksheet(
        filepath="game_data/balance_analysis.xlsx",
        sheet_name="character_balance",
        headers=balance_headers
    )
    
    # 2. 插入角色数据
    character_balance_data = [
        [1, "战士", 150, 80, 100, 500, 1.5, 6.25, 75.0],
        [2, "法师", 200, 40, 120, 300, 1.67, 7.5, 83.3],
        [3, "牧师", 100, 60, 90, 400, 1.11, 6.67, 70.0],
        [4, "刺客", 180, 30, 150, 250, 1.2, 8.33, 90.0]
    ]
    
    await excel_write_rows(
        filepath="game_data/balance_analysis.xlsx",
        sheet_name="character_balance",
        data=character_balance_data,
        start_cell="A2"
    )
    
    # 3. 计算平衡评分
    await excel_apply_formula(
        filepath="game_data/balance_analysis.xlsx",
        sheet_name="character_balance",
        cell="I2",
        formula="=(G2*0.4 + H2*0.6) * 10"
    )
    
    # 4. 设置条件格式识别不平衡角色
    await excel_format_range(
        filepath="game_data/balance_analysis.xlsx",
        sheet_name="character_balance",
        start_cell="I2",
        end_cell="I100",
        conditional_format={
            "type": "cell",
            "criteria": "less than",
            "value": "75",
            "format": {"bg_color": "#ffcdd2", "font_color": "#d32f2f"}
        }
    )
    
    # 5. 生成平衡建议
    imbalance_query = await excel_sql_query(
        query="""
        SELECT character_name, balance_score, 
               CASE 
                 WHEN balance_score < 70 THEN '严重不平衡'
                 WHEN balance_score < 80 THEN '轻微不平衡' 
                 ELSE '平衡良好'
               END as balance_status
        FROM character_balance
        WHERE balance_score < 80
        ORDER BY balance_score ASC
        """,
        filepath="game_data/balance_analysis.xlsx"
    )
    
    print("需要平衡的角色：")
    for row in imbalance_query['data']:
        print(f"{row[0]}: {row[2]} (评分: {row[1]})")
    
    # 6. 执行平衡调整
    balance_adjustments = [
        {"row": 3, "column": 4, "value": 100},  # 牧师速度提升
        {"row": 4, "column": 5, "value": 300},  # 刺客HP提升
        {"row": 1, "column": 3, "value": 70}   # 战士伤害下调
    ]
    
    await excel_update_rows(
        filepath="game_data/balance_analysis.xlsx",
        sheet_name="character_balance",
        updates=balance_adjustments
    )
    
    # 7. 重新计算平衡评分
    await excel_apply_formula(
        filepath="game_data/balance_analysis.xlsx",
        sheet_name="character_balance",
        range="I2:I100",
        formula="=(G2*0.4 + H2*0.6) * 10"
    )
    
    # 8. 导出平衡报告
    balance_report = await excel_sql_query(
        query="""
        SELECT character_name, balance_score, 
               ROUND((balance_score - 80) / 80 * 100, 1) as deviation_percentage
        FROM character_balance
        ORDER BY balance_score DESC
        """,
        filepath="game_data/balance_analysis.xlsx"
    )
    
    print("最终平衡报告：")
    for row in balance_report['data']:
        deviation = row[2]
        status = "平衡" if abs(deviation) <= 5 else "需要调整"
        print(f"{row[0]}: {row[1]}分 ({status}, 偏差{abs(deviation)}%)")

if __name__ == "__main__":
    asyncio.run(numerical_balance())
```

---

## 📝 使用建议

1. **配置文件管理**：建议将游戏配置文件按功能分类，便于维护
2. **版本控制**：重要配置修改前先创建备份，避免数据丢失
3. **性能优化**：大数据量操作时使用streaming参数提升性能
4. **数据验证**：设置数据验证规则确保数据质量
5. **权限管理**：敏感配置数据建议加密存储

## 🔄 示例更新

随着ExcelMCP功能的持续更新，本示例库也会不断完善。欢迎提交使用示例和改进建议。