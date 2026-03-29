# 角色属性管理 - 游戏配置Excel MCP服务器

## 🎮 第2步：实际场景 - 角色属性管理（10分钟）

### 场景介绍
假设你正在开发一款RPG游戏，需要管理角色属性表。

### 2.1 查看现有角色表
```json
{
  "method": "tools/call",
  "params": {
    "name": "list_sheets",
    "arguments": {}
  }
}
```

**预期结果**：
- `characters` - 角色属性表
- `skills` - 技能配置表  
- `equipment` - 装备数据表

### 2.2 查看角色表结构
```json
{
  "method": "tools/call", 
  "params": {
    "name": "get_headers",
    "arguments": {
      "sheet": "characters"
    }
  }
}
```

**预期结果**：
```json
{
  "headers": ["id", "name", "level", "hp", "mp", "attack", "defense", "agility", "skills"]
}
```

### 2.3 查看具体角色数据
```json
{
  "method": "tools/call",
  "params": {
    "name": "get_range", 
    "arguments": {
      "sheet": "characters",
      "range": "A2:D10"
    }
  }
}
```

### 📝 练习2.1
**任务**：查看ID为3的角色的详细属性
**提示**：使用`get_range`获取完整的角色数据，找到ID=3的行

### 2.4 查询角色统计信息
```json
{
  "method": "tools/call",
  "params": {
    "name": "describe_table",
    "arguments": {
      "sheet": "characters"
    }
  }
}
```

### 2.5 搜索特定角色
```json
{
  "method": "tools/call",
  "params": {
    "name": "search",
    "arguments": {
      "sheet": "characters",
      "criteria": {
        "level": 5
      }
    }
  }
}
```

### 🎯 角色管理要点
- **ID唯一性**：确保每个角色有唯一标识符
- **属性平衡**：注意血量、攻击力、防御力的数值平衡
- **技能关联**：角色技能列表应与技能表保持一致

---

## 💡 实用技巧

### 批量查看角色
```json
{
  "method": "tools/call",
  "params": {
    "name": "get_range",
    "arguments": {
      "sheet": "characters",
      "range": "A2:Z100"
    }
  }
}
```

### 查找最高等级角色
```json
{
  "method": "tools/call",
  "params": {
    "name": "query",
    "arguments": {
      "sheet": "characters",
      "query": "SELECT name, level FROM characters ORDER BY level DESC LIMIT 5"
    }
  }
}
```

*下一步：[技能配置管理](tutorial-skills.md)*