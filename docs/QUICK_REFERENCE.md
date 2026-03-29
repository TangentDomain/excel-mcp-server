# MCP快速参考指南 - 游戏配置Excel操作速查

> 🎮 **30秒找到你需要的命令** - 按场景分类的MCP操作速查表

## 📋 快速导航

### 🚀 按场景分类
| 场景类型 | 常用操作 | 见章节 |
|---------|---------|--------|
| 📊 基础读写 | 读取、写入、更新 | 基础操作 |
| 🔍 数据查询 | 精确查找、条件筛选 | 查询操作 |  
| 🔗 关联查询 | 跨表JOIN、子查询 | 高级查询 |
| 📈 批量处理 | 批量增删改、流式处理 | 批量操作 |
| 🎯 游戏专用 | 技能、装备、角色等 | 游戏场景 |

### 🎮 按功能分类
| 功能类型 | MCP命令 | 描述 |
|---------|---------|------|
| **表操作** | `list_sheets` | 列出所有工作表 |
| **表信息** | `describe_table` | 表结构描述 |
| **表头** | `get_headers` | 获取列名 |
| **范围读取** | `get_range` | 读取数据范围 |
| **查询** | `query` | SQL查询支持 |
| **批量插入** | `batch_insert_rows` | 批量添加行 |
| **删除** | `delete_rows` | 删除数据 |
| **更新** | `update_range` | 更新数据 |

---

## 🎮 游戏场景速查

### ⚔️ 技能系统操作
| 需求 | MCP命令示例 |
|------|-------------|
| 查看所有技能 | `query "SELECT * FROM skills"` |
| 找出法师技能 | `query "SELECT * FROM skills WHERE class = '法师'"` |
| 技能攻击力+10% | `update_range "skills!C2:C100" "attack * 1.1"` |
| CT查询技能树 | `query "WITH RECURSIVE skill_tree AS..."` |

### 🛡️ 装备管理
| 需求 | MCP命令示例 |
|------|-------------|
| 装备按稀有度排序 | `query "SELECT * FROM equipment ORDER BY rarity DESC"` |
| 查询套装效果 | `query "SELECT * FROM equipment WHERE set_id IS NOT NULL"` |
| 装备评分计算 | `update_range "equipment!F2:F100" "attack * 0.4 + defense * 0.3 + magic * 0.3"` |

### 👥 角色管理
| 需求 | MCP命令示例 |
|------|-------------|
| 角色列表 | `query "SELECT name, class, level FROM characters"` |
| 高等级角色筛选 | `query "SELECT * FROM characters WHERE level > 50"` |
| 批量创建角色 | `batch_insert_rows "characters" [[name, class, level, hp], ["新角色1", "战士", 1, 100]]` |

---

## 📊 核心功能对比表

| 功能 | 支持情况 | 性能 | 备注 |
|------|----------|------|------|
| **基础CRUD** | ✅ 完全支持 | ⚡ 极快 | 所有Excel操作 |
| **SQL查询** | ✅ 完全支持 | ⚡ 快 | JOIN/CTE/子查询 |
| **批量操作** | ✅ 完全支持 | ⚡ 快 | 支持万级数据 |
| **数据验证** | ✅ 完全支持 | ⚡ 快 | 自动类型检查 |
| **错误处理** | ✅ 智能修复 | ⚡ 快 | 自动修复建议 |
| **版本管理** | ✅ 自动同步 | ⚡ 快 | 防止版本冲突 |

---

## 🚀 一键操作

### 安装与启动
```bash
# 一行安装（推荐）
curl -LsSf https://astral.sh/uv/install.sh | sh && uvx excel-mcp-server-fastmcp

# 传统安装
pip install excel-mcp-server-fastmcp
```

### 配置AI客户端
```json
// Claude Desktop/Cursor配置
{
  "mcpServers": {
    "excelmcp": {
      "command": "uvx",
      "args": ["excel-mcp-server-fastmcp"]
    }
  }
}
```

### 常用操作模板
```bash
# 启动后直接使用MCP命令
{"method": "tools/call", "params": {
  "name": "get_range", 
  "arguments": {
    "sheet": "技能表",
    "range": "A1:Z100"
  }
}}
```

---

## 💡 最佳实践

### 🎯 查询优化
- 使用具体列名而非 `SELECT *` 提升性能
- 复杂查询先在小数据集测试
- 定期使用 `find_last_row` 避免读取空白行

### 🔧 错误处理
- 遇到格式错误时，先检查数据类型
- 批量操作前先备份重要数据
- 使用 `describe_table` 了解表结构

### 📈 性能提示
- 大文件处理时分批操作
- 使用条件查询减少数据量
- 定期清理临时工作表

---

## 🆘 急速故障排除

### 常见问题快速解决

| 现象 | 解决方案 | 验证命令 |
|------|----------|----------|
| 连接失败 | 检查uvx安装: `uvx --version` | `uvx excel-mcp-server-fastmcp --help` |
| 查询无结果 | 检查sheet名称和范围 | `list_sheets` → `describe_table` |
| 格式错误 | 检查数据类型一致性 | `get_range` 查看实际数据 |
| 权限问题 | 文件读写权限检查 | Excel文件属性 → 编辑权限 |

### 性能问题诊断
```bash
# 检查大文件性能
# 1. 先测试小范围: get_range "sheet!A1:A10"
# 2. 逐步扩大范围确认性能瓶颈
# 3. 复杂查询先分步执行
```

---

## 📚 扩展资源

### 🎓 学习路径
1. **基础**：阅读本速查指南 → [互动式教程](INTERACTIVE_TUTORIAL.md)
2. **进阶**：查看 [完整文档](README.md) → [技术规格](README.md#📈-技术规格)
3. **实战**：参考 [游戏场景演示](README.md#🎮-游戏开发场景演示)

### 🛠️ 开发工具
- **API文档**：查看所有53个MCP工具详细说明
- **示例代码**：游戏中常见配置操作模板
- **测试用例**：验证功能的测试文件

> 💡 **提示**：遇到问题先查看故障排除，大部分问题30秒内解决！