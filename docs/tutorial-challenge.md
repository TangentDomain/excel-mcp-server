# 综合实战挑战 - 游戏配置Excel MCP服务器

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

#### 任务1：创建Boss角色
```json
{
  "method": "tools/call",
  "params": {
    "name": "batch_insert_rows",
    "arguments": {
      "sheet": "characters",
      "rows": [
        ["boss_001", "炎魔领主", 20, 5000, 200, 150, 80, 50, "火焰吐息,雷电打击,护甲碎击"]
      ]
    }
  }
}
```

#### 任务2：添加Boss技能
```json
{
  "method": "tools/call",
  "params": {
    "name": "batch_insert_rows",
    "arguments": {
      "sheet": "skills",
      "rows": [
        ["skill_boss_001", "火焰吐息", "attack", "造成200%攻击力伤害", 5.0, 100],
        ["skill_boss_002", "雷电打击", "attack", "造成150%攻击力伤害，有麻痹效果", 8.0, 80],
        ["skill_boss_003", "护甲碎击", "debuff", "减少目标50%防御力，持续3秒", 12.0, 60]
      ]
    }
  }
}
```

#### 任务3：关联技能到角色
```json
{
  "method": "tools/call",
  "params": {
    "name": "batch_insert_rows",
    "arguments": {
      "sheet": "character_skills",
      "rows": [
        ["boss_001", "skill_boss_001"],
        ["boss_001", "skill_boss_002"],
        ["boss_001", "skill_boss_003"]
      ]
    }
  }
}
```

#### 任务4：平衡性分析
```json
{
  "method": "tools/call",
  "params": {
    "name": "query",
    "arguments": {
      "query": "SELECT 
        '炎魔领主' as boss_name,
        5000 as boss_hp,
        150 as boss_attack,
        80 as boss_defense
      UNION ALL
      SELECT 
        '普通玩家' as player_name,
        1000 as player_hp,
        80 as player_attack,
        60 as player_defense
      UNION ALL
      SELECT 
        '坦克玩家' as player_name,
        1500 as player_hp,
        70 as player_attack,
        100 as player_defense
      UNION ALL
      SELECT 
        '输出玩家' as player_name,
        800 as player_hp,
        120 as player_attack,
        40 as player_defense
      UNION ALL
      SELECT 
        '治疗玩家' as player_name,
        900 as player_hp,
        60 as player_attack,
        50 as player_defense"
    }
  }
}
```

#### 任务5：战斗时间预估
```json
{
  "method": "tools/call",
  "params": {
    "name": "query",
    "arguments": {
      "query": "SELECT 
        boss_name,
        player_name,
        boss_hp,
        player_dps,
        boss_hp / player_dps as estimated_time_seconds,
        boss_hp / player_dps / 60 as estimated_time_minutes
      FROM (
        SELECT 
          '炎魔领主' as boss_name,
          5000 as boss_hp
        UNION ALL
        SELECT 
          '普通玩家' as player_name,
          80 * 1.2 as player_dps
        UNION ALL
        SELECT 
          '坦克玩家' as player_name,
          70 * 1.1 as player_dps
        UNION ALL
        SELECT 
          '输出玩家' as player_name,
          120 * 1.3 as player_dps
        UNION ALL
        SELECT 
          '治疗玩家' as player_name,
          60 * 0.8 as player_dps
      ) combined_data"
    }
  }
}
```

### 🎯 挑战评估标准

#### ✅ 完成标准
- [ ] Boss角色创建成功
- [ ] Boss技能配置正确
- [ ] 角色技能关联正确
- [ ] 平衡性分析完整
- [ ] 战斗时间预估合理（3-5分钟）

#### 💡 高级要求（可选）
- [ ] 添加技能特殊效果描述
- [ ] 创建Boss战斗策略
- [ ] 设计难度递进机制

### 🚀 完成检查清单

1. **数据验证**：确认所有数据已正确添加
2. **关联检查**：确保角色与技能正确关联
3. **平衡测试**：验证Boss难度适中
4. **性能测试**：查询响应时间正常

---

恭喜完成综合实战挑战！你已经掌握了MCP服务器的高级应用技能。

*返回：[教程总结](tutorial-summary.md)*