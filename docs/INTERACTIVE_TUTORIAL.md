# 互动式教程 - 游戏配置Excel MCP服务器快速上手

> 🎮 **专为游戏策划和分析师设计** - 通过实际游戏场景学习Excel配置操作

## 📚 教程概述

本教程将带你从零开始，掌握使用MCP服务器操作游戏配置表的完整流程。每个步骤都包含实际的游戏场景示例和练习。

### 🎯 学习目标
- 理解MCP服务器的基本概念
- 掌握Excel配置表的核心操作
- 能够独立完成游戏配置管理任务
- 理解高级查询和数据处理功能

### ⏱️ 预计时间
- 初学者：30-45分钟
- 有经验者：15-20分钟

---

## 🚀 第1步：基础概念（5分钟）

### 什么是MCP服务器？
MCP（Message Passing Control）服务器是一个让AI能够直接操作Excel配置表的服务。它就像一个"超级Excel"，可以通过自然语言命令来读写、查询和分析游戏数据。

### 游戏配置表的重要性
游戏配置表包含：
- **角色属性**：血量、攻击力、防御力、技能
- **装备数据**：武器、防具、饰品属性
- **技能配置**：技能效果、冷却时间、消耗
- **数值平衡**：怪物强度、掉落概率、经验值
- **关卡设计**：地图配置、敌人分布、任务条件

---

## 🎮 第2步：实际场景 - 角色属性管理（10分钟）

### 场景介绍
假设你正在开发一款RPG游戏，需要管理角色属性表。

### 2.1 查看现有角色表
```bash
# 启动MCP服务器
excel-mcp-server-fastmcp
```

**MCP命令**：
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

---

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

### 📝 练习3.1
**任务**：
1. 查看所有"治疗"类型的技能
2. 为ID=1的角色添加一个新技能"火球术"
**提示**：使用`search`查找治疗技能，使用`update_range`更新角色技能列表

---

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

### 📝 练习4.1
**任务**：
1. 查找HP值超过1000的角色
2. 计算平均攻击力
3. 找出技能数量最多的3个角色
**提示**：使用`query`和SQL语句完成复杂分析

---

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

### 5.3 数据备份和验证
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

### 📝 练习5.1
**任务**：
1. 将所有攻击力低于50的角色等级提升1级
2. 删除MP值为0的技能
3. 验证数据完整性

---

## 🏆 第6步：综合实战挑战（10分钟）

### 挑战场景
你正在为一个新的副本Boss设计配置，需要：

1. **创建新Boss角色**：
   - 等级：20级
   - HP：5000，MP：200
   - 攻击：150，防御：80
   - 技能：["火焰吐息", "雷电打击", "护甲碎击"]

2. **设计Boss技能**：
   - 火焰吐息：造成200%攻击力伤害，冷却5秒
   - 雷电打击：造成150%攻击力伤害，有麻痹效果，冷却8秒
   - 护甲碎击：减少目标50%防御力，持续3秒，冷却12秒

3. **平衡性测试**：
   - 计算Boss对5个普通玩家的威胁等级
   - 确保战斗时间在3-5分钟内

### 📝 挑战任务
请使用学到的MCP命令，完成以下操作：

1. 在characters表中添加Boss数据
2. 在skills表中添加Boss的3个技能
3. 查询验证数据正确性
4. 计算并记录战斗时间预估

---

## 🎓 教程总结

### ✅ 已掌握的技能
- [x] 基础表操作（查看、读取、描述）
- [x] 数据查询和筛选
- [x] 批量数据操作
- [x] SQL高级查询
- [x] 数据维护和优化
- [x] 综合实战应用

### 📚 下一步学习建议
1. **深入高级功能**：学习子查询、JOIN操作、复杂计算
2. **性能优化**：了解大数据表的查询优化技巧
3. **团队协作**：学习多人协作配置管理
4. **自动化**：掌握批量操作脚本编写

### 🆘 遇到问题？
1. 检查[故障排除指南](./TROUBLESHOOTING.md)
2. 查看[FAQ文档](./FAQ.md) 
3. 访问[示例代码](../examples/)
4. 提交[Issue反馈](https://github.com/TangentDomain/excel-mcp-server/issues)

---

## 🎮 快速参考

### 常用MCP命令速查
| 功能 | 命令 | 用途 |
|------|------|------|
| 查看所有表 | `list_sheets` | 获取所有工作表名称 |
| 查看表结构 | `get_headers` | 获取列标题信息 |
| 读取数据 | `get_range` | 按范围读取数据 |
| 查询表 | `describe_table` | 获取表的统计信息 |
| 搜索数据 | `search` | 按条件搜索数据 |
| 批量插入 | `batch_insert_rows` | 批量添加新数据 |
| SQL查询 | `query` | 执行SQL查询语句 |
| 更新数据 | `update_range` | 按范围更新数据 |
| 删除数据 | `delete_rows` | 按条件删除数据 |
| 查找最后一行 | `find_last_row` | 获取数据行数 |

### 游戏场景示例
```bash
# 示例1：查询所有战士职业角色
query "SELECT * FROM characters WHERE class = 'warrior'"

# 示例2：计算装备总价值
query "SELECT equipment_id, SUM(value) as total_value FROM equipment GROUP BY equipment_id"

# 示例3：查找过期数据
search "sheet='events' AND status='expired'"
```

---

*教程版本：v1.6.49*  
*最后更新：2026-03-29*  
*维护者：excel-mcp-server开发团队*