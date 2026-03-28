# 🎮 Game Development Scenarios

ExcelMCP is specifically optimized for game development workflows. This guide covers common game development scenarios and how to use ExcelMCP to streamline your workflow.

## Typical Game Configuration Tables

### 1. Skills Table (技能表)
```excel
| 技能ID | 技能名 | 技能类型 | 伤害 | 消耗蓝 | 冷却时间 | 描述 |
|--------|--------|----------|------|--------|----------|------|
| SK001  | 火球术 | 火系攻击 | 120  | 50     | 3s       | 发射火球攻击敌人 |
| SK002  | 冰箭   | 冰系攻击 | 80   | 30     | 2s       | 发射冰箭造成减速 |
```

### 2. Equipment Table (装备表)
```excel
| 装备ID | 装备名 | 装备类型 | 攻击力 | 防御力 | 生命值 | 品质 |
|--------|--------|----------|--------|--------|--------|------|
| EQ001  | 钢剑   | 武器     | 45     | 0      | 0      | 稀有 |
| EQ002  | 皮甲   | 防具     | 0      | 25     | 100    | 普通 |
```

### 3. Monsters Table (怪物表)
```excel
| 怪物ID | 怪物名 | 类型 | 生命值 | 攻击力 | 防御力 | 掉落率 |
|--------|--------|------|--------|--------|--------|--------|
| MN001  | 史莱姆 | 元素 | 100    | 15     | 5      | 80%    |
| MN002  | 哥布林 | 人形 | 150    | 25     | 10     | 60%    |
```

## Common Game Development Workflows

### 1. Skill Balancing
```python
# Find all skills with damage > 100
skills = excel_query_sql("skills.xlsx", "SELECT * FROM skills WHERE 伤害 > 100")

# Increase damage of all fire skills by 20%
excel_update_rows("skills.xlsx", "skills", 
                  "技能类型 = '火系攻击'", 
                  {"伤害": "伤害 * 1.2"})

# Create balance report
balance_report = excel_query_sql("skills.xlsx", """
    SELECT 
        技能类型,
        AVG(伤害) as avg_damage,
        COUNT(*) as skill_count
    FROM skills 
    GROUP BY 技能类型
    ORDER BY avg_damage DESC
""")
```

### 2. Equipment Optimization
```python
# Find equipment with best stats for each type
best_equipment = excel_query_sql("equipment.xlsx", """
    SELECT 
        装备类型,
        装备名,
        攻击力 + 防御力 as total_stats
    FROM equipment
    WHERE 品质 = '稀有'
    ORDER BY total_stats DESC
""")

# Calculate equipment scores
excel_add_formula("equipment.xlsx", "equipment", "I2", 
                  "SUM(C2:G2) + H2 * 10")

# Compare equipment sets
comparison = excel_compare_sheets("equipment_v1.xlsx", "装备表",
                                "equipment_v2.xlsx", "装备表")
```

### 3. Monster Design
```python
# Analyze monster difficulty
monster_analysis = excel_query_sql("monsters.xlsx", """
    SELECT 
        类型,
        AVG(生命值) as avg_hp,
        AVG(攻击力) as avg_attack,
        AVG(防御力) as avg_defense
    FROM monsters
    GROUP BY 类型
""")

# Find monsters with high drop rates
high_drop_monsters = excel_query_sql("monsters.xlsx", 
    "SELECT * FROM monsters WHERE 掉落率 > 70%")

# Balance monster progression
progression_query = excel_query_sql("monsters.xlsx", """
    SELECT 怪物ID, 怪物名, 生命值, 攻击力,
           ROUND(生命值 / 10 + 攻击力 / 2) as difficulty_score
    FROM monsters
    ORDER BY difficulty_score DESC
""")
```

### 4. Game Economy Balance
```python
# Analyze item drop rates
economy_analysis = excel_query_sql("items.xlsx", """
    SELECT 
        物品类型,
        AVG(掉落率) as avg_drop_rate,
        COUNT(*) as item_count
    FROM items
    GROUP BY 物品类型
    HAVING AVG(掉落率) < 50%
""")

# Calculate item values
excel_query_sql("items.xlsx", """
    UPDATE items 
    SET 价值 = 攻击力 * 5 + 防御力 * 3 + 生命值 * 2
    WHERE 物品类型 = '消耗品'
""")

# Find most valuable items
valuable_items = excel_query_sql("items.xlsx", """
    SELECT 物品名, 价值
    FROM items
    WHERE 价值 > 1000
    ORDER BY 价值 DESC
""")
```

## Advanced Game Development Features

### 1. Game Data Validation
```python
# Validate skill data
skill_validation = excel_query_sql("skills.xlsx", """
    SELECT * FROM skills 
    WHERE 伤害 < 0 OR 消耗蓝 < 0 OR 冷却时间 < 0
""")

# Check for duplicate skill IDs
duplicate_skills = excel_query_sql("skills.xlsx", """
    SELECT 技能ID, COUNT(*) as count
    FROM skills
    GROUP BY 技能ID
    HAVING COUNT(*) > 1
""")

# Validate equipment stats
equipment_validation = excel_query_sql("equipment.xlsx", """
    SELECT * FROM equipment 
    WHERE 攻击力 < 0 OR 防御力 < 0
""")
```

### 2. Game Progression Analysis
```python
# Analyze player progression
progression_analysis = excel_query_sql("player_data.xlsx", """
    SELECT 
        等级,
        AVG(攻击力) as avg_attack,
        AVG(防御力) as avg_defense,
        AVG(生命值) as avg_hp
    FROM player_data
    GROUP BY 等级
    ORDER BY 等级
""")

# Find power spikes
power_spikes = excel_query_sql("player_data.xlsx", """
    SELECT 等级,
           AVG(攻击力) - LAG(AVG(攻击力), 1) OVER (ORDER BY 等级) as attack_growth,
           AVG(防御力) - LAG(AVG(防御力), 1) OVER (ORDER BY 等级) as defense_growth
    FROM player_data
    GROUP BY 等级
    HAVING (AVG(攻击力) - LAG(AVG(攻击力), 1) OVER (ORDER BY 等级)) > 20
""")
```

### 3. Multi-Server Data Management
```python
# Combine player data from multiple servers
combined_data = excel_query_sql("""
    SELECT * FROM 
        (SELECT * FROM server1_players UNION ALL 
         SELECT * FROM server2_players) as combined
    WHERE 等级 >= 50
    ORDER BY 等级 DESC
""")

# Compare server balance
server_comparison = excel_query_sql("""
    SELECT 
        'server1' as server_name,
        AVG(等级) as avg_level,
        COUNT(*) as player_count
    FROM server1_players
    
    UNION ALL
    
    SELECT 
        'server2' as server_name,
        AVG(等级) as avg_level,
        COUNT(*) as player_count
    FROM server2_players
""")

# Find top players across servers
top_players = excel_query_sql("""
    SELECT * FROM (
        SELECT *, 'server1' as server FROM server1_players
        UNION ALL
        SELECT *, 'server2' as server FROM server2_players
    ) as all_players
    ORDER BY 战斗力 DESC
    LIMIT 100
""")
```

## Game Development Templates

### 1. New Game Setup Template
```python
# Create new game project structure
excel_create_workbook("new_game.xlsx")
excel_create_sheet("new_game.xlsx", "技能表")
excel_create_sheet("new_game.xlsx", "装备表")
excel_create_sheet("new_game.xlsx", "怪物表")
excel_create_sheet("new_game.xlsx", "物品表")

# Add initial data
initial_skills = [
    ["SK001", "火球术", "火系攻击", 120, 50, 3, "发射火球攻击敌人"],
    ["SK002", "冰箭", "冰系攻击", 80, 30, 2, "发射冰箭造成减速"]
]
excel_write_data("new_game.xlsx", "技能表", initial_skills, start_cell="A1")
```

### 2. Balance Check Template
```python
# Balance check automation
balance_check = excel_query_sql("game_data.xlsx", """
    -- Check skill balance
    SELECT 
        'Skill Balance' as check_type,
        '技能' as category,
        AVG(伤害) as avg_value,
        STDDEV(伤害) as std_dev,
        STDDEV(伤害) / AVG(伤害) as cv_ratio
    FROM skills
    
    UNION ALL
    
    -- Check equipment balance
    SELECT 
        'Equipment Balance' as check_type,
        '装备' as category,
        AVG(攻击力 + 防御力) as avg_value,
        STDDEV(攻击力 + 防御力) as std_dev,
        STDDEV(攻击力 + 防御力) / AVG(攻击力 + 防御力) as cv_ratio
    FROM equipment
""")
```

## Performance Tips for Game Data

### 1. Large Configuration Files
```python
# Use streaming for large files
excel_write_only_workbook("large_config.xlsx")
excel_write_only_override("large_config.xlsx", "Skills", skill_data_stream)

# Query in chunks
for chunk in range(0, total_rows, 10000):
    chunk_data = excel_query_range("large_file.xlsx", "Sheet1", 
                                   f"A{chunk+1}:Z{chunk+10000}")
    # Process chunk
```

### 2. Frequent Updates
```python
# Batch updates for frequent changes
batch_updates = [
    {"sheet": "Skills", "conditions": "技能类型 = '火系'", "updates": {"伤害": "伤害 * 1.1"}},
    {"sheet": "Skills", "conditions": "技能类型 = '冰系'", "updates": {"伤害": "伤害 * 1.05"}},
    {"sheet": "Skills", "conditions": "技能类型 = '雷系'", "updates": {"伤害": "伤害 * 1.15"}}
]

excel_batch_update("game_config.xlsx", batch_updates)
```

### 3. Version Control
```python
# Create backup before major changes
excel_copy_workbook("current_config.xlsx", "config_backup.xlsx")

# Track changes
changes = excel_compare_sheets("config_v1.xlsx", "Skills",
                             "config_v2.xlsx", "Skills")

# Generate changelog
changelog = excel_query_sql("config_changes.xlsx", """
    SELECT 
        changed_id,
        old_value,
        new_value,
        change_type,
        change_time
    FROM changes
    ORDER BY change_time DESC
""")
```

## Best Practices for Game Development

### 1. Data Consistency
- Use consistent naming conventions
- Implement data validation rules
- Create backup before major changes
- Document all configuration changes

### 2. Performance Optimization
- Use batch operations for large datasets
- Implement proper indexing for frequent queries
- Monitor memory usage with large files
- Use streaming writes for data exports

### 3. Team Collaboration
- Maintain separate configuration files for different environments
- Use version control for all game data
- Document configuration standards and conventions
- Regular balance reviews and adjustments