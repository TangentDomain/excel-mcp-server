# 数据维护优化 - 游戏配置Excel MCP服务器

## 🛠️ 第5步：数据维护和优化（7分钟）

### 场景介绍
游戏更新时，需要批量修改数据和维护数据完整性。

### 5.1 删除无效数据
```json
{
  "method": "tools/call", 
  "params": {
    "name": "delete_rows",
    "arguments": {
      "sheet": "characters",
      "criteria": {
        "level": 0
      }
    }
  }
}
```

### 5.2 批量更新角色等级
```json
{
  "method": "tools/call",
  "params": {
    "name": "update_range",
    "arguments": {
      "sheet": "characters",
      "range": "C2:C100",
      "values": [[5], [5], [5], [5], [5]]  # 将所有角色等级设置为5
    }
  }
}
```

### 5.3 条件更新 - 装备强化
```json
{
  "method": "tools/call",
  "params": {
    "name": "update_range",
    "arguments": {
      "sheet": "equipment",
      "range": "F2:F50",
      "criteria": {
        "rarity": "legendary"
      },
      "values": [[10], [10], [10], [10], [10]]  # 传说装备+10强化
    }
  }
}
```

### 5.4 数据备份和验证
```json
{
  "method": "tools/call",
  "params": {
    "name": "find_last_row",
    "arguments": {
      "sheet": "characters"
    }
  }
}
```

### 5.5 数据完整性检查
```json
{
  "method": "tools/call",
  "params": {
    "name": "query",
    "arguments": {
      "query": "SELECT c.name, COUNT(cs.skill_id) as skill_count FROM characters c LEFT JOIN character_skills cs ON c.id = cs.character_id GROUP BY c.name HAVING skill_count = 0"
    }
  }
}
```

### 📝 练习5.1
**任务**：
1. 将所有攻击力低于50的角色等级提升1级
2. 删除MP值为0的技能
3. 验证数据完整性

### 🎯 数据维护要点
- **备份策略**：重要操作前先备份数据
- **批量操作**：使用批量命令提高效率
- **数据验证**：操作后检查数据完整性
- **错误处理**：有异常时及时回滚

### 💡 维护最佳实践

### 规范化更新
```sql
-- 批量更新属性平衡
UPDATE characters 
SET attack = attack * 1.1, defense = defense * 1.1
WHERE level BETWEEN 10 AND 20
```

### 数据清洗脚本
```sql
-- 删除重复数据
DELETE FROM skills 
WHERE id NOT IN (
  SELECT MIN(id) FROM skills 
  GROUP BY skill_name, type
)
```

### 数据统计报告
```sql
-- 生成维护报告
SELECT 
  'characters' as table_name,
  COUNT(*) as total_records,
  COUNT(CASE WHEN level = 0 THEN 1 END) as invalid_records,
  ROUND(AVG(level), 2) as avg_level
FROM characters

UNION ALL

SELECT 
  'skills' as table_name,
  COUNT(*) as total_records,
  COUNT(CASE WHEN cooldown < 0 THEN 1 END) as invalid_records,
  ROUND(AVG(cooldown), 2) as avg_cooldown
FROM skills
```

*下一步：[综合实战挑战](tutorial-challenge.md)*