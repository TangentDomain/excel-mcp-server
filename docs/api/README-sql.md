# 🔧 SQL Query Guide

ExcelMCP provides powerful SQL query capabilities for Excel configuration tables. This guide covers all supported SQL features with game development examples.

## Supported SQL Features

### Basic SQL Operations
```sql
-- SELECT basic query
SELECT * FROM skills WHERE 技能类型 = '火系'

-- SELECT specific columns
SELECT 技能名, 伤害, 冷却时间 FROM skills WHERE 伤害 > 100

-- SELECT with calculations
SELECT 技能名, 伤害 * 1.2 as 增强伤害 FROM skills
```

### WHERE Clause Conditions
```sql
-- Basic conditions
SELECT * FROM skills WHERE 伤害 > 100
SELECT * FROM equipment WHERE 防御力 >= 50
SELECT * FROM monsters WHERE 生命值 < 200

-- Multiple conditions (AND)
SELECT * FROM skills WHERE 技能类型 = '火系' AND 伤害 > 100

-- Multiple conditions (OR)
SELECT * FROM skills WHERE 技能类型 = '火系' OR 技能类型 = '冰系'

-- Combined conditions
SELECT * FROM skills WHERE (技能类型 = '火系' OR 技能类型 = '冰系') AND 冷却时间 <= 3
```

### ORDER BY Sorting
```sql
-- Basic sorting
SELECT * FROM skills ORDER BY 伤害 DESC
SELECT * FROM equipment ORDER BY 攻击力 + 防御力 ASC

-- Multi-column sorting
SELECT * FROM skills ORDER BY 技能类型 ASC, 伤害 DESC

-- Random sorting
SELECT * FROM skills ORDER BY RANDOM()
```

### GROUP BY Aggregation
```sql
-- Group by category
SELECT 技能类型, AVG(伤害) as 平均伤害, COUNT(*) as 技能数量
FROM skills
GROUP BY 技能类型

-- Group with multiple columns
SELECT 品质, 装备类型, AVG(攻击力) as 平均攻击力
FROM equipment
GROUP BY 品质, 装备类型

-- Group with conditions
SELECT 技能类型, AVG(伤害) as 平均伤害
FROM skills
WHERE 冷却时间 <= 5
GROUP BY 技能类型
HAVING AVG(伤害) > 80
```

### JOIN Operations
```sql
-- INNER JOIN
SELECT s.技能名, s.伤害, e.装备名, e.攻击力
FROM skills s
INNER JOIN equipment e ON s.装备ID = e.装备ID

-- LEFT JOIN
SELECT s.技能名, s.伤害, e.装备名
FROM skills s
LEFT JOIN equipment e ON s.装备ID = e.装备ID

-- RIGHT JOIN
SELECT s.技能名, s.伤害, e.装备名
FROM skills s
RIGHT JOIN equipment e ON s.装备ID = e.装备ID

-- FULL JOIN
SELECT s.技能名, s.伤害, e.装备名
FROM skills s
FULL JOIN equipment e ON s.装备ID = e.装备ID
```

### Subqueries
```sql
-- Subquery in WHERE clause
SELECT * FROM skills WHERE 伤害 > (SELECT AVG(伤害) FROM skills)

-- Subquery in SELECT clause
SELECT 技能名, 伤害, (SELECT AVG(伤害) FROM skills) as 全局平均伤害
FROM skills

-- Subquery in FROM clause
SELECT * FROM (
    SELECT 技能类型, AVG(伤害) as 平均伤害
    FROM skills
    GROUP BY 技能类型
) as skill_avg
WHERE 平均伤害 > 100
```

### Advanced SQL Features

#### CASE Statements
```sql
-- Conditional logic
SELECT 
    技能名,
    技能类型,
    CASE 
        WHEN 伤害 > 200 THEN '高伤害'
        WHEN 伤害 > 100 THEN '中等伤害'
        ELSE '低伤害'
    END as 伤害等级,
    CASE 
        WHEN 冷却时间 <= 2 THEN '快速'
        WHEN 冷却时间 <= 5 THEN '中等'
        ELSE '慢速'
    END as 冷却速度
FROM skills
```

#### Window Functions
```sql
-- Row numbering
SELECT 
    技能名,
    伤害,
    ROW_NUMBER() OVER (ORDER BY 伤害 DESC) as 伤害排名
FROM skills

-- Running totals
SELECT 
    技能名,
    伤害,
    SUM(伤害) OVER (ORDER BY 技能名) as 累计伤害
FROM skills

-- Moving averages
SELECT 
    技能名,
    伤害,
    AVG(伤害) OVER (ORDER BY 技能名 ROWS 2 PRECEDING) as 移动平均
FROM skills
```

#### Common Table Expressions (CTE)
```sql
-- Simple CTE
WITH 技能统计 AS (
    SELECT 
        技能类型,
        AVG(伤害) as 平均伤害,
        COUNT(*) as 技能数量
    FROM skills
    GROUP BY 技能类型
)
SELECT * FROM 技能统计 WHERE 平均伤害 > 100

-- Recursive CTE
WITH RECURSIVE 技能等级 AS (
    SELECT 技能名, 伤害, 1 as 等级
    FROM skills
    WHERE 伤害 <= 50
    UNION ALL
    SELECT s.技能名, s.伤害, sl.等级 + 1
    FROM skills s
    JOIN 技能等级 sl ON s.伤害 > sl.伤害 + 20
)
SELECT DISTINCT 技能名, 伤害, 等级
FROM 技能等级
ORDER BY 等级, 伤害
```

#### Pivot/Unpivot Operations
```sql
-- Pivot example (turn rows into columns)
SELECT 
    技能名,
    MAX(CASE WHEN 属性 = '火' THEN 数值 ELSE 0 END) as 火属性,
    MAX(CASE WHEN 属性 = '冰' THEN 数值 ELSE 0 END) as 冰属性,
    MAX(CASE WHEN 属性 = '雷' THEN 数值 ELSE 0 END) as 雷属性
FROM skill_attributes
PIVOT (
    SUM(数值)
    FOR 属性 IN (火, 冰, 雷)
)
```

#### Set Operations
```sql
-- UNION (remove duplicates)
SELECT 技能名 FROM 火系技能
UNION
SELECT 技能名 FROM 冰系技能

-- UNION ALL (keep duplicates)
SELECT 技能名 FROM 火系技能
UNION ALL
SELECT 技能名 FROM 冰系技能

-- INTERSECT
SELECT 技能名 FROM 火系技能
INTERSECT
SELECT 技能名 from 冰系技能

-- EXCEPT
SELECT 技能名 FROM 火系技能
EXCEPT
SELECT 技能名 FROM 冰系技能
```

## Game Development SQL Examples

### 1. Skill Analysis
```sql
-- Find most powerful skills
SELECT 技能名, 伤害, 消耗蓝, 
       (伤害 / 消耗蓝) as 伤害效率
FROM skills
ORDER BY 伤害效率 DESC

-- Analyze skill types
SELECT 
    技能类型,
    COUNT(*) as 技能数量,
    AVG(伤害) as 平均伤害,
    MAX(伤害) as 最大伤害,
    MIN(伤害) as 最小伤害,
    STDDEV(伤害) as 伤害标准差
FROM skills
GROUP BY 技能类型
ORDER BY 平均伤害 DESC

-- Find skills with high damage/cooling ratio
SELECT 
    技能名,
    技能类型,
    伤害,
    冷却时间,
    (伤害 / 冷却时间) as 伤害效率
FROM skills
WHERE 冷却时间 > 0
ORDER BY 伤害效率 DESC
```

### 2. Equipment Optimization
```sql
-- Find best equipment combinations
SELECT 
    e1.装备名 as 主装备,
    e2.装备名 as 副装备,
    e1.攻击力 + e2.攻击力 as 总攻击力,
    e1.防御力 + e2.防御力 as 总防御力
FROM equipment e1
CROSS JOIN equipment e2
WHERE e1.装备类型 = '武器' AND e2.装备类型 = '防具'
ORDER BY 总攻击力 DESC

-- Calculate equipment scores
SELECT 
    装备名,
    攻击力 * 0.6 + 防御力 * 0.4 as 综合评分
FROM equipment
ORDER BY 综合评分 DESC

-- Find equipment with best value
SELECT 
    装备名,
    品质,
    攻击力 + 防御力 + 生命值 as 总属性,
    价格,
    (攻击力 + 防御力 + 生命值) / 价格 as 性价比
FROM equipment
ORDER BY 性价比 DESC
```

### 3. Monster Balance Analysis
```sql
-- Analyze monster difficulty
SELECT 
    类型,
    AVG(生命值) as 平均生命值,
    AVG(攻击力) as 平均攻击力,
    AVG(防御力) as 平均防御力,
    AVG(生命值 + 攻击力 + 防御力) as 综合难度
FROM monsters
GROUP BY 类型
ORDER BY 综合难度 DESC

-- Find monsters with specific drop rates
SELECT 
    怪物名,
    类型,
    生命值,
    攻击力,
    掉落率,
    CASE 
        WHEN 掉落率 > 80 THEN '超高掉落'
        WHEN 掉落率 > 60 THEN '高掉落'
        WHEN 掉落率 > 40 THEN '中等掉落'
        ELSE '低掉落'
    END as 掉落等级
FROM monsters
ORDER BY 掉落率 DESC

-- Calculate monster experience value
SELECT 
    怪物ID,
    怪物名,
    类型,
    生命值,
    攻击力,
    防御力,
    ROUND(生命值 * 0.3 + 攻击力 * 0.5 + 防御力 * 0.2) as 经验值
FROM monsters
ORDER BY 经验值 DESC
```

### 4. Game Economy Analysis
```sql
-- Analyze item prices
SELECT 
    物品类型,
    AVG(价格) as 平均价格,
    MIN(价格) as 最低价格,
    MAX(价格) as 最高价格,
    STDDEV(价格) as 价格波动
FROM items
GROUP BY 物品类型
ORDER BY 平均价格 DESC

-- Find most valuable items
SELECT 
    物品名,
    物品类型,
    价格,
    攻击力,
    防御力,
    生命值,
    价格 + 攻击力 * 5 + 防御力 * 3 + 生命值 * 2 as 总价值
FROM items
ORDER BY 总价值 DESC

-- Analyze crafting requirements
SELECT 
    成品名,
    成品类型,
    SUM(材料价格) as 材料成本,
    成品价格,
    (成品价格 - SUM(材料价格)) as 利润,
    (成品价格 - SUM(材料价格)) / 成品价格 as 利润率
FROM recipes
GROUP BY 成品名, 成品类型
ORDER BY 利润率 DESC
```

### 5. Player Progression Analysis
```sql
-- Analyze player level progression
SELECT 
    等级,
    AVG(攻击力) as 平均攻击力,
    AVG(防御力) as 平均防御力,
    AVG(生命值) as 平均生命值,
    AVG(经验值) as 平均经验值
FROM player_data
GROUP BY 等级
ORDER BY 等级

-- Find power spikes
SELECT 
    等级,
    AVG(攻击力) - LAG(AVG(攻击力), 1) OVER (ORDER BY 等级) as 攻击力增长,
    AVG(防御力) - LAG(AVG(防御力), 1) OVER (ORDER BY 等级) as 防御力增长,
    AVG(生命值) - LAG(AVG(生命值), 1) OVER (ORDER BY 等级) as 生命值增长
FROM player_data
GROUP BY 等级
HAVING (AVG(攻击力) - LAG(AVG(攻击力), 1) OVER (ORDER BY 等级)) > 20

-- Calculate player balance metrics
SELECT 
    等级,
    AVG(攻击力) / AVG(生命值) as 攻击生命比,
    AVG(防御力) / AVG(攻击力) as 防御攻击比,
    STDDEV(攻击力) / AVG(攻击力) as 攻击力变异系数
FROM player_data
GROUP BY 等级
ORDER BY 等级
```

## Performance Optimization Tips

### 1. Query Optimization
```sql
-- Use specific column names instead of *
SELECT 技能名, 伤害, 冷却时间 FROM skills

-- Add WHERE clauses to limit results
SELECT * FROM skills WHERE 技能类型 = '火系' AND 伤害 > 100

-- Use indexes effectively
-- ExcelMCP automatically creates indexes on frequently queried columns
```

### 2. Batch Operations
```sql
-- Instead of multiple single queries, use batch operations
-- Update multiple skills at once
UPDATE skills 
SET 伤害 = 伤害 * 1.2 
WHERE 技能类型 = '火系' AND 冷却时间 <= 3

-- Insert multiple skills at once
INSERT INTO skills (技能ID, 技能名, 技能类型, 伤害, 消耗蓝, 冷却时间)
VALUES 
    ('SK100', '烈焰风暴', '火系', 200, 80, 5),
    ('SK101', '冰霜新星', '冰系', 150, 60, 4),
    ('SK102', '雷电链', '雷系', 180, 70, 3)
```

### 3. Memory Management
```sql
-- Use LIMIT for large datasets
SELECT * FROM skills WHERE 技能类型 = '火系' LIMIT 100

-- Use pagination for large results
SELECT * FROM skills ORDER BY 伤害 DESC LIMIT 50 OFFSET 100

-- Clear cache when working with large datasets
-- ExcelMCP automatically manages cache, but manual clearing is possible
```

## Error Handling and Debugging

### Common SQL Errors
```sql
-- Invalid column names
-- ERROR: Column 'skill_name' not found
-- Solution: Use correct column names (技能名)

-- Missing table names
-- ERROR: Table 'skill_table' not found
-- Solution: Use correct table names (skills)

-- Invalid data types
-- ERROR: Can't compare string with number
-- Solution: Ensure data types match in comparisons
```

### Debugging SQL Queries
```sql
-- Step 1: Check table structure
DESCRIBE skills

-- Step 2: Test simple conditions first
SELECT * FROM skills WHERE 技能类型 = '火系'

-- Step 3: Gradually build complex queries
SELECT * FROM skills 
WHERE 技能类型 = '火系' 
AND 伤害 > (SELECT AVG(伤害) FROM skills)
```

### Performance Monitoring
```sql
-- Check query execution time
-- ExcelMCP provides performance metrics for each query

-- Monitor memory usage
-- ExcelMCP automatically tracks memory usage and provides warnings

-- Analyze slow queries
-- Use ExcelMCP's performance analysis tools to identify bottlenecks
```

## Best Practices

### 1. Query Design
- Use specific column names instead of `*`
- Add appropriate WHERE clauses to limit result sets
- Use JOIN operations instead of multiple separate queries
- Use GROUP BY with appropriate HAVING clauses

### 2. Performance Optimization
- Use LIMIT for large datasets
- Use indexes effectively (ExcelMCP creates them automatically)
- Use batch operations for multiple updates
- Monitor memory usage for large files

### 3. Data Integrity
- Validate data before complex operations
- Use transactions for critical operations
- Create backups before major modifications
- Use data validation rules to maintain consistency

### 4. Game Development Specific
- Design queries for game balance analysis
- Use SQL for progression planning
- Create queries for economy analysis
- Use statistical analysis for balance tuning

## Advanced Examples

### 1. Complex Game Analysis
```sql
-- Advanced skill balance analysis
WITH 技能分析 AS (
    SELECT 
        技能类型,
        AVG(伤害) as 平均伤害,
        AVG(消耗蓝) as 平均消耗,
        AVG(冷却时间) as 平均冷却,
        AVG(伤害 / 消耗蓝) as 效率,
        AVG(伤害 / 冷却时间) as 伤害密度
    FROM skills
    GROUP BY 技能类型
),
平衡性评分 AS (
    SELECT 
        技能类型,
        平均伤害,
        平均消耗,
        平均冷却,
        效率,
        伤害密度,
        ABS(效率 - (SELECT AVG(效率) FROM 技能分析)) as 效率偏差,
        ABS(伤害密度 - (SELECT AVG(伤害密度) FROM 技能分析)) as 密度偏差
    FROM 技能分析
)
SELECT 
    技能类型,
    平均伤害,
    效率,
    密度,
    ROUND(效率偏差 / 效率 * 100, 2) as 效率偏差率,
    ROUND(密度偏差 / 密度 * 100, 2) as 密度偏差率,
    CASE 
        WHEN 效率偏差率 < 10 AND 密度偏差率 < 10 THEN '平衡'
        WHEN 效率偏差率 > 20 OR 密度偏差率 > 20 THEN '不平衡'
        ELSE '基本平衡'
    END as 平衡状态
FROM 平衡性评分
ORDER BY 效率偏差率 + 密度偏差率 DESC
```

### 2. Advanced Game Economy
```sql
-- Complex economy analysis
WITH 物品价值分析 AS (
    SELECT 
        物品类型,
        AVG(价格) as 平均价格,
        COUNT(*) as 物品数量,
        AVG(攻击力 + 防御力 + 生命值) as 总属性
    FROM items
    GROUP BY 物品类型
),
经济平衡 AS (
    SELECT 
        i.物品类型,
        i.平均价格,
        i.物品数量,
        i.总属性,
        CASE 
            WHEN i.总属性 > 0 THEN i.平均价格 / i.总属性
            ELSE 0
        END as 属性价值比,
        AVG(w.平均价格) over () as 全局平均价格,
        AVG(w.总属性) over () as 全局平均属性
    FROM 物品价值分析 i
    CROSS JOIN 物品价值分析 w
)
SELECT 
    物品类型,
    平均价格,
    物品数量,
    总属性,
    ROUND(属性价值比, 2) as 属性价值比,
    ROUND(平均价格 - 全局平均价格, 2) as 价格偏差,
    ROUND(总属性 - 全局平均属性, 2) as 属性偏差,
    CASE 
        WHEN 属性价值比 > 全局平均价格 / 全局平均属性 * 1.5 THEN '高价低质'
        WHEN 属性价值比 < 全局平均价格 / 全局平均属性 * 0.5 THEN '低价优质'
        ELSE '价值合理'
    END as 价值评估
FROM 经济平衡
ORDER BY 属性价值比 DESC
```