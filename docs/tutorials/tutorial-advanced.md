# 高级查询分析 - 游戏配置Excel MCP服务器

## 🔍 第4步：高级查询 - 数值平衡分析（8分钟）

### 场景介绍
作为数值策划，你需要分析角色属性平衡性，找出异常数据。

### 4.1 查询攻击力最高的角色
```json
{
  "method": "tools/call",
  "params": {
    "name": "query",
    "arguments": {
      "sheet": "characters",
      "query": "SELECT name, attack FROM characters ORDER BY attack DESC LIMIT 5"
    }
  }
}
```

### 4.2 找出防御力异常的角色
```json
{
  "method": "tools/call", 
  "params": {
    "name": "query",
    "arguments": {
      "sheet": "characters", 
      "query": "SELECT name, level, defense FROM characters WHERE defense > 100 OR defense < 10"
    }
  }
}
```

### 4.3 统计各等级段的角色数量
```json
{
  "method": "tools/call",
  "params": {
    "name": "query",
    "arguments": {
      "sheet": "characters",
      "query": "SELECT level, COUNT(*) as count FROM characters GROUP BY level ORDER BY level"
    }
  }
}
```

### 4.4 跨表查询 - 技能使用统计
```json
{
  "method": "tools/call",
  "params": {
    "name": "query",
    "arguments": {
      "query": "SELECT c.name, s.skill_name, s.type, s.cooldown FROM characters c JOIN character_skills cs ON c.id = cs.character_id JOIN skills s ON cs.skill_id = s.id WHERE s.type = 'attack'"
    }
  }
}
```

### 4.5 复杂计算 - 伤害输出分析
```json
{
  "method": "tools/call",
  "params": {
    "name": "query",
    "arguments": {
      "query": "SELECT name, attack, defense, (attack * 1.5 - defense * 0.8) as estimated_dps FROM characters ORDER BY estimated_dps DESC"
    }
  }
}
```

### 📝 练习4.1
**任务**：
1. 查找HP值超过1000的角色
2. 计算平均攻击力
3. 找出技能数量最多的3个角色
**提示**：使用`query`和SQL语句完成复杂分析

### 🎯 高级查询要点
- **性能优化**：大数据集时使用索引字段
- **结果验证**：手动验证查询结果的合理性
- **数据分析**：通过统计发现数值平衡问题

### 💡 实用查询模板

### 属性分布分析
```sql
SELECT 
  level,
  COUNT(*) as character_count,
  AVG(attack) as avg_attack,
  AVG(defense) as avg_defense,
  AVG(hp) as avg_hp
FROM characters 
GROUP BY level
ORDER BY level
```

### 技能使用频率统计
```sql
SELECT 
  s.type,
  s.cooldown,
  COUNT(*) as skill_count,
  AVG(cs.frequency) as avg_usage
FROM skills s
LEFT JOIN character_skills cs ON s.id = cs.skill_id
GROUP BY s.type, s.cooldown
ORDER BY skill_count DESC
```

### 异常数据检测
```sql
-- 攻击力异常检测
SELECT name, attack FROM characters 
WHERE attack > (SELECT AVG(attack) + 2*STDDEV(attack) FROM characters)
   OR attack < (SELECT AVG(attack) - 2*STDDEV(attack) FROM characters)
```

*下一步：[数据维护优化](tutorial-maintenance.md)*