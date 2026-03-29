# 技能配置管理 - 游戏配置Excel MCP服务器

## ⚔️ 第3步：技能配置管理（10分钟）

### 场景介绍
管理角色技能，包括技能效果、冷却时间和消耗。

### 3.1 查看技能表结构
```json
{
  "method": "tools/call",
  "params": {
    "name": "describe_table",
    "arguments": {
      "sheet": "skills"
    }
  }
}
```

### 3.2 查找攻击系技能
```json
{
  "method": "tools/call",
  "params": {
    "name": "search",
    "arguments": {
      "sheet": "skills",
      "criteria": {
        "type": "attack"
      }
    }
  }
}
```

### 3.3 批量添加新技能
```json
{
  "method": "tools/call",
  "params": {
    "name": "batch_insert_rows",
    "arguments": {
      "sheet": "skills",
      "rows": [
        ["skill_007", "火焰冲击", "attack", "造成150%攻击力伤害", 3.0, 50],
        ["skill_008", "冰霜护盾", "defense", "减少20%受到的伤害", 0.0, 30]
      ]
    }
  }
}
```

### 3.4 更新技能效果
```json
{
  "method": "tools/call",
  "params": {
    "name": "update_range",
    "arguments": {
      "sheet": "skills",
      "range": "D2:D100",
      "values": [["造成180%攻击力伤害"], ["减少30%受到的伤害"]]
    }
  }
}
```

### 📝 练习3.1
**任务**：
1. 查看所有"治疗"类型的技能
2. 为ID=1的角色添加一个新技能"火球术"
**提示**：使用`search`查找治疗技能，使用`update_range`更新角色技能列表

### 🎯 技能管理要点
- **技能类型**：攻击、防御、治疗、辅助等明确分类
- **冷却时间**：平衡技能使用频率
- **消耗资源**：MP、消耗品等合理配置
- **效果描述**：清晰说明技能效果和适用场景

### 💡 高级技能操作

### 删除过期技能
```json
{
  "method": "tools/call",
  "params": {
    "name": "delete_rows",
    "arguments": {
      "sheet": "skills",
      "criteria": {
        "status": "deprecated"
      }
    }
  }
}
```

### 查询技能统计
```json
{
  "method": "tools/call",
  "params": {
    "name": "query",
    "arguments": {
      "sheet": "skills",
      "query": "SELECT type, COUNT(*) as count, AVG(cooldown) as avg_cooldown FROM skills GROUP BY type"
    }
  }
}
```

*下一步：[高级查询分析](tutorial-advanced.md)*