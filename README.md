<div align="center">

[中文](README.md) ｜ [English](README.en.md)

</div>

# 🎮 ExcelMCP: 游戏开发专用 Excel 配置表管理器

[![许可证: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 版本](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![技术支持: FastMCP](https://img.shields.io/badge/Powered%20by-FastMCP-orange)](https://github.com/jlowin/fastmcp)
![状态](https://img.shields.io/badge/status-stable-green.svg)
![测试覆盖](https://img.shields.io/badge/tests-1164%20tests-brightgreen.svg)
![工具数量](https://img.shields.io/badge/tools-52%20verified%20tools-green.svg)
[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/)
[![版本](https://img.shields.io/badge/version-v1.6.33-blue.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/#history)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)

**ExcelMCP** 是专为游戏开发设计的Excel配置表管理MCP服务器。通过AI自然语言指令，实现技能配置表、装备数据、怪物属性等游戏配置的智能化操作。基于**FastMCP**构建，读取使用**python-calamine**（Rust引擎，2300x提速），写入使用**openpyxl**，拥有**44个专业工具**和**1164个测试用例**，确保企业级可靠性。

🎯 **核心功能**: 技能系统、装备管理、怪物配置、数值平衡、版本对比、策划工具链

## 🏆 竞品对比

| 功能特性 | ExcelMCP | haris-musa/excelpython |
|---------|----------|----------------------|
| **架构设计** | MCP服务器模式 | Python库模式 |
| **AI集成** | ✅ 原生MCP支持 | ❌ 需额外封装 |
| **性能** | ⚡ Rust引擎(calamine)+2300x提速 | 🐢 纯Python处理 |
| **SQL引擎** | ✅ 45项SQL功能+安全解析 | ❌ 基础查询支持 |
| **游戏垂直优化** | ✅ 专用游戏数据结构 | ❌ 通用表格处理 |
| **跨文件JOIN** | ✅ `@'filepath'`语法 | ❌ 不支持 |
| **错误处理** | ✅ 结构化错误+AI修复提示 | ❌ 基础异常 |
| **测试覆盖** | ✅ 1164个测试用例 | ❌ 有限测试 |
| **当前版本** | ✅ v1.6.31 | ✅ 最新版本 |
| **安装方式** | ✅ `uvx`一键运行 | ⚠️ 需要pip安装 |

## 🚀 为什么选择 ExcelMCP
- **专为游戏开发设计**：技能、装备、怪物等游戏数据的智能化管理
- **AI原生集成**：支持自然语言指令，无需学习复杂API
- **企业级性能**：Rust引擎读取 + 流式写入，支持超大配置表
- **零配置运行**：PyPI一键安装，自动更新，开箱即用

## ⚡ 性能优化

### 流式写入（Streaming Write）
针对大数据量修改操作，ExcelMCP采用**copy-modify-write**方案：
- **读取**：Rust引擎(calamine)快速读取，内存占用与文件大小无关
- **修改**：内存中进行数据修改
- **写入**：openpyxl write_only模式写入，避免全量文件加载

**适用场景**：
- 批量插入1000+行数据
- 大文件配置表修改
- 内存敏感环境

**性能提升**：
- 内存降低90%+
- 大文件处理速度提升5-10倍
- 支持GB级配置表操作

**自动降级**：流式模式失败时自动回退到传统模式，确保功能可靠性

📦 **一键安装**: `uvx excel-mcp-server-fastmcp` — 从PyPI直接运行，自动更新，零配置

---

## ⚡ 30秒上手

**1. 安装**（你的 MCP 客户端配置文件加这段）：
```json
{
  "mcpServers": {
    "excelmcp": {
      "command": "uvx",
      "args": ["excel-mcp-server-fastmcp"]
    }
  }
}
```

**2. 对AI说一句话就行：**
```
→ "查看 skills.xlsx 有哪些表"
→ "搜索所有火系技能"
→ "查询DPM最高的10个技能"
→ "把所有火系技能伤害提升20%"
→ "对比v1和v2版本差异"
```

**搞定。** 不需要记命令、不需要写SQL（当然你也可以写）。支持中文自然语言 + 标准SQL双模式。

## 🔧 SQL实战场景

### 高级查询
```sql
-- 跨文件JOIN查询技能和装备数据
SELECT s.技能名, s.伤害, e.装备名, e.攻击力 
FROM skills.xlsx技能表 s
JOIN equipment.xlsx装备表 e ON s.职业 = e.适用职业
WHERE s.元素 = '火' AND e.品质 = '史诗'
```

### 复杂分析
```sql
-- 使用窗口函数计算技能排名
SELECT 技能名, 伤害, 
       RANK() OVER (ORDER BY 伤害 DESC) as 伤害排名,
       PERCENT_RANK() OVER (ORDER BY 伤害 DESC) as 伤害百分位
FROM skills.xlsx技能表
WHERE 元素 IN ('火', '冰', '雷')
```

### 数据修改
```sql
-- 批量更新技能数值
UPDATE skills.xlsx技能表 
SET 伤害 = 伤害 * 1.2, 
    消耗法力 = 消耗法力 * 1.1
WHERE 元素 = '火' AND 技能类型 = '主动'
```

### 子查询和CTE
```sql
-- 使用WITH子查询构建复杂分析
with 高伤害技能 as (
    SELECT 技能名, 伤害 FROM skills.xlsx技能表 
    WHERE 伤害 > 1000
)
SELECT s.技能名, s.伤害, e.装备名
FROM 高伤害技能 s
JOIN equipment.xlsx装备表 e ON s.技能名 = e.推荐技能
```

<details>
<summary>📖 想了解更详细的安装方式？</summary>

## 🚀 快速入门

### 方式一：uvx 一键运行（推荐）

无需克隆项目，从PyPI直接运行：

```bash
uvx excel-mcp-server-fastmcp
```

MCP客户端配置：
```json
{
  "mcpServers": {
    "excelmcp": {
      "command": "uvx",
      "args": ["excel-mcp-server-fastmcp"]
    }
  }
}
```


### 方式二：从源码安装

1. **克隆项目**
   ```bash
   git clone https://github.com/TangentDomain/excel-mcp-server.git
   cd excel-mcp-server
   ```

2. **安装依赖**
   ```bash
   # 推荐：使用 uv (更快)
   pip install uv && uv sync

   # 备选：使用 pip
   pip install -e .
   ```

3. **配置MCP客户端**
   ```json
   {
     "mcpServers": {
       "excelmcp": {
         "command": "python",
         "args": ["-m", "excel_mcp_server_fastmcp"]
       }
     }
   }
   ```

4. **开始使用**
   准备就绪！让AI助手通过自然语言控制Excel文件。

### 验证安装
```bash
# 检查版本
excel-mcp-server-fastmcp --version

# 运行测试
python -m pytest tests/ --tb=short -q

# 运行性能基准测试
python scripts/benchmark.py --quick        # 快速模式（约30秒）
python scripts/benchmark.py                # 完整模式（含大表测试）
python scripts/benchmark.py --compare      # 与上次结果对比
```

</details>

---

## 📊 为什么选 ExcelMCP？

| 特性 | ExcelMCP | [haris-musa/excel-mcp](https://github.com/haris-musa/excel-mcp) | [excelpython](https://github.com/nicepkg/excelpython) |
|------|----------|-------|------------|
| **SQL查询引擎** | ✅ 完整SQL（JOIN/子查询/窗口函数/CTE） | ❌ | ❌ |
| **读取引擎** | 🦀 python-calamine（Rust，2300x提速） | openpyxl | openpyxl |
| **工具数量** | 44个专业工具 | ~15个基础工具 | ~10个 |
| **游戏开发** | ✅ 垂直优化（DPM/数值平衡/配置表） | ❌ 通用 | ❌ 通用 |
| **双行表头** | ✅ 自动识别中文描述+英文字段 | ❌ | ❌ |
| **SQL UPDATE** | ✅ 条件批量修改 | ❌ | ❌ |
| **跨文件JOIN** | ✅ @filepath 语法 | ❌ | ❌ |
| **测试覆盖** | 1164 tests | ~50 tests | ~30 tests |
| **错误恢复** | ✅ 结构化错误码+AI可修复提示 | ❌ 纯文本 | ❌ 纯文本 |
| **中文列名查询** | ✅ | ❌ | ❌ |
| **备份/恢复** | ✅ | ❌ | ❌ |
| **版本对比** | ✅ Sheet级diff | ❌ | ❌ |

> 💡 **核心差异**：ExcelMCP 是唯一支持 **完整SQL引擎** 的 Excel MCP 服务器。其他方案只能"读/写单元格"，ExcelMCP 能让你用 SQL 做复杂的数据分析和批量修改——这对游戏配置表管理是质变。

---

## ⚡ 快速参考

### 🎯 常用命令速查表

#### ⭐ 基础操作 (新手级)
```text
读取数据:      "读取 sales.xlsx 的 A1:C10 范围数据"
文件信息:      "获取 report.xlsx 的基本信息"
简单搜索:      "在 skills.xlsx 中查找'火球术'"
```

#### ⭐⭐ 数据操作 (进阶级)
```text
更新数据:      "将 skills.xlsx 第2列所有数值乘以1.2"
格式设置:      "把 report.xlsx 第一行设为粗体，背景浅蓝"
插入行:        "在 inventory.xlsx 第5行插入3个空行"
```

#### ⭐⭐⭐ 游戏开发专用 (专家级)
```text
配置对比:      "比较v1.0和v1.1版本技能表，生成变更报告"
批量分析:      "分析所有20-30级怪物的血量攻击比"
属性调整:      "将装备表中传说品质物品属性提升25%"
```

### 🎮 游戏开发场景速查

| 场景 | 推荐工具 | 示例命令 |
|------|----------|----------|
| 技能平衡调整 | `excel_search` + `excel_update_range` | "将所有火系技能伤害提升20%" |
| 装备配置管理 | `excel_format_cells` + `excel_get_range` | "用金色标记所有传说装备" |
| 怪物数据验证 | `excel_check_duplicate_ids` + `excel_search` | "确保怪物ID唯一，血量合理" |
| 版本对比分析 | `excel_compare_sheets` + `excel_compare_files` | "对比新旧版本配置表差异" |
| 数据统计查询 | `excel_query` | "查询技能表中各职业平均攻击力" |
| 条件批量修改 | `excel_update_query` | "把火系技能伤害提升20%" |
| 批量修改前预览 | `excel_assess_data_impact`（detailed=False快速预览，detailed=True全面评估） | "预览删除第5-10行的影响" |
| 修改前备份 | `excel_create_backup` | "备份当前技能表再修改" |
| 公式试算 | `excel_evaluate_formula` | "临时计算SUM(A2:A100)看结果" |

### 🔧 范围表达式参考

| 格式 | 说明 | 示例 |
|------|------|------|
| `Sheet1!A1:C10` | 标准范围 | "技能表!A1:D50" |
| `Sheet1!1:5` | 行范围 | "配置表!2:100" |
| `Sheet1!B:D` | 列范围 | "数据表!B:G" |
| `Sheet1!A1` | 单单元格 | "设置表!A1" |

---

## 🎮 游戏策划完整工作流教程

> 从零开始，手把手教你用自然语言操作Excel配置表。不需要记住任何命令格式，用中文描述你想做什么就行。

### 📦 第一步：了解你的表（DESCRIBE）

拿到一张配置表，先看看有什么数据：

```
"查看 skills.xlsx 的技能表结构"
→ excel_describe_table("skills.xlsx", "SkillConfig")
```

返回结果：
```
列名          | 类型    | 描述     | 非空 | 样本值
skill_id     | int     | 技能ID   | 10/10 | 1001, 1002
skill_name   | str     | 技能名称 | 10/10 | 火球术, 治愈之光
damage       | float   | 伤害值   | 9/10  | 150.0, 200.5
cooldown     | int     | 冷却时间 | 10/10 | 5, 10
```

💡 **小贴士**：双行表头的配置表（第1行中文描述+第2行英文字段名）会被自动识别，列描述自动关联。

### 🔍 第二步：搜索定位（SEARCH）

找特定数据：

```
"在 skills.xlsx 中搜索所有火系技能"
→ excel_search("skills.xlsx", "火")

"搜索所有包含'传说'的装备名称"
→ excel_search("equipment.xlsx", "传说", "EquipmentConfig")
```

### 📊 第三步：SQL查询分析（QUERY）

这是最强大的功能。用标准SQL语法查询配置表：

**基础查询 — 找数据：**
```sql
-- 查看所有10级以上技能
SELECT * FROM SkillConfig WHERE level >= 10

-- 只看技能名和伤害，按伤害排序
SELECT skill_name, damage FROM SkillConfig ORDER BY damage DESC LIMIT 10

-- 分页查看：每页5条，看第3页
SELECT * FROM MonsterConfig ORDER BY level LIMIT 5 OFFSET 10
```

**中文列名查询 — 策划友好：**
```sql
-- 双行表头时直接用中文名查询
SELECT 技能名称, 伤害值 FROM SkillConfig WHERE 等级 >= 10

-- 中文列名 + 英文列名混用也可以
SELECT skill_name, 伤害值 FROM SkillConfig WHERE 技能类型 = '攻击'
```

**聚合统计 — 数值分析：**
```sql
-- 各职业平均伤害
SELECT skill_type, AVG(damage) as avg_dmg, COUNT(*) as cnt
FROM SkillConfig GROUP BY skill_type

-- 哪些技能类型总伤害超过1000
SELECT skill_type, SUM(damage) as total
FROM SkillConfig GROUP BY skill_type HAVING total > 1000

-- 装备品质分布
SELECT DISTINCT quality FROM EquipmentConfig
```

**DPM数值平衡分析：**
```sql
-- 每秒伤害排名（DPM = damage / cooldown）
SELECT skill_name, damage * 1.0 / cooldown as dpm
FROM SkillConfig ORDER BY dpm DESC LIMIT 10
```

**数据质量检查：**
```sql
-- 找出有缺失值的配置
SELECT skill_name, description FROM SkillConfig WHERE description IS NULL

-- 找出特定等级范围的怪物
SELECT name, level, hp FROM MonsterConfig WHERE level BETWEEN 10 AND 20

-- 排除测试数据
SELECT * FROM SkillConfig WHERE skill_name NOT LIKE '%测试%'
```

### ✏️ 第四步：批量修改（UPDATE）

两种方式：

**方式一：SQL UPDATE（推荐，精确条件修改）：**
```
"将所有火系技能伤害提升20%"
→ excel_update_query("skills.xlsx", "UPDATE SkillConfig SET damage = damage * 1.2 WHERE skill_type = '火系'")
```

⚠️ **修改前预览一下**：
```
"预览一下火系技能伤害提升20%会改哪些"
→ excel_update_query("skills.xlsx", "UPDATE SkillConfig SET damage = damage * 1.2 WHERE skill_type = '火系'", dry_run=True)
```

**方式二：范围写入（已知区域批量写入）：**
```
"将 skills.xlsx 第2行到第50行的伤害列数值全部乘以1.15"
→ excel_update_range("skills.xlsx", "SkillConfig!E2:E50", [[...]])
```

⚠️ **修改前一定要备份：**
```
"备份 skills.xlsx"
→ excel_create_backup("skills.xlsx")
```

### 🔄 第五步：版本对比（COMPARE）

改完之后对比一下：

```
"对比 v1.0 和 v1.1 的技能表差异"
→ excel_compare_sheets("skills_v1.0.xlsx", "SkillConfig",
                        "skills_v1.1.xlsx", "SkillConfig")
```

### 📋 常用策划场景速查

| 我想做 | 怎么说 |
|--------|--------|
| 看看表里有什么 | "查看xxx表结构" |
| 找某个技能/装备 | "搜索xxx" |
| 按条件筛选 | "查询等级>10的技能" |
| 统计各类型数量 | "各职业技能有多少个" |
| 找最强的技能 | "DPM最高的10个技能" |
| 找有问题的数据 | "哪些技能描述是空的" |
| 批量改数值 | "把所有火系技能伤害提升20%" |
| 条件批量改 | "UPDATE技能表 SET 伤害=伤害*1.1 WHERE 元素='火'" |
| 对比版本差异 | "对比v1和v2的配置表" |

### ❓ 常见错误和解决

**Q: 列名拼错了怎么办？**
A: 系统会自动推荐相似列名。比如你写 `skil_name`，会提示"你是否想用: skill_name?"

**Q: 表太大查询慢？**
A: 同一张表重复查询会自动缓存，第二次查询速度提升30-100倍。2000行的大表首次~230ms，缓存后仅需2-8ms。

**Q: JOIN怎么用？**
A: 支持同文件内工作表关联查询：
```sql
SELECT a.skill_name, b.equip_name FROM SkillConfig a INNER JOIN EquipConfig b ON a.equip_id = b.equip_id
```

---

## 🛠️ 完整工具列表（44个专业工具）

### 📁 文件与工作表管理
- `excel_create_file` - 创建新Excel文件，支持自定义工作表
- `excel_create_sheet` - 添加新工作表
- `excel_delete_sheet` - 删除工作表
- `excel_list_sheets` - 📋 智能工作表扫描，查询前必用工具
- `excel_rename_sheet` - 重命名工作表
- `excel_copy_sheet` - 复制工作表（含数据和格式），创建配置表变体
- `excel_get_file_info` - 获取文件元数据
- `excel_merge_files` - 合并多个Excel文件

### 📊 数据操作
- `excel_get_range` - 🔍 精确范围读取，支持格式保持
- `excel_update_range` - ✏️ 批量数据写入，支持流式模式
- `excel_get_headers` - 📋 智能表头提取，支持双行表头
- `excel_insert_rows` - 插入空行
- `excel_delete_rows` - 删除行范围
- `excel_insert_columns` - 插入空列
- `excel_delete_columns` - 删除列范围
- `excel_find_last_row` - 查找最后一行有数据位置
- `excel_rename_column` - 重命名列（修改表头单元格值，支持双行表头）
- `excel_upsert_row` - Upsert行（按键列查找，存在则更新，不存在则插入，策划合并配置高频操作）
- `excel_batch_insert_rows` - 批量插入多行数据到工作表末尾（策划批量导入配置）
- `excel_set_formula` - 设置单元格公式（自动计算）
- `excel_evaluate_formula` - 临时执行公式返回结果，不修改文件

### 🔍 搜索与分析
- `excel_search` - 正则表达式搜索
- `excel_search_directory` - 📂 目录批量搜索，跨文件查找目标内容
- `excel_query` - 🚀 SQL查询引擎，游戏配置表超强分析工具
- `excel_update_query` - ⚙️ SQL批量修改器，精确控制数据变更
- `excel_get_headers` - 📋 智能表头提取，支持双行表头

**功能说明**：
专门提取Excel表头信息，支持游戏开发的双行表头模式（第1行中文描述 + 第2行英文字段名），返回结构化的表头信息。

**使用示例**：
```python
# 快速查看特定表的表头
result = excel_get_headers("skills.xlsx", "SkillConfig")

# 返回格式（改进后）：
{
    "success": true,
    "data": {
        "field_names": ["skill_id", "skill_name", "damage", "cooldown"],
        "descriptions": ["技能ID描述", "技能名称描述", "伤害值描述", "冷却时间描述"],
        "dual_rows": true
    },
    "meta": {
        "sheet_name": "SkillConfig",
        "header_row": 1,
        "header_count": 4,
        "dual_row_mode": true
    },
    "message": "成功获取4个表头字段（描述+字段名）"
}

# 批量查看所有表的表头
result = excel_get_headers("skills.xlsx")  # sheet_name省略
```

**使用场景**：
- **快速了解表结构**：在操作前先了解字段名称和描述
- **确认列名拼写**：避免SQL查询时列名错误
- **批量检查多个表**：获取Excel文件中所有工作表的表头信息
- **双行表头支持**：自动识别并分离字段描述和字段名

**与excel_describe_table的区别**：
- `excel_get_headers`：专注表头信息，快速轻量
- `excel_describe_table`：完整表结构，包含类型、样本值、统计信息

`excel_describe_table` - 📊 完整表结构分析，类型+样本+统计信息
- `excel_compare_sheets` - 🆔 工作表对比，发现配置变更和差异
- `excel_compare_files` - 📋 多工作表文件对比，版本追踪
- `excel_check_duplicate_ids` - 🔍 ID重复检测，数据质量保障
- `excel_server_stats` - 📈 服务器运行统计，性能监控与错误分析

### 🔄 批量操作

- `excel_batch_update_ranges` - 🔄 批量更新多个单元格区域，减少IO提升性能
- `excel_merge_multiple_files` - 📁 合并多个Excel文件，支持追加/覆盖/合并模式

### 📊 图表生成

- `excel_create_chart` - 📊 创建柱状图/折线图/饼图等，数据可视化
- `excel_list_charts` - 📋 列出工作表中的所有图表信息

### 📝 数据验证

- `excel_set_data_validation` - 📝 设置数据验证规则，限制输入类型和范围
- `excel_clear_validation` - 🧹 清除数据验证规则

### 🎨 条件格式

- `excel_add_conditional_format` - 🎨 添加条件格式，数据可视化增强
- `excel_clear_conditional_format` - 🧹 清除条件格式规则

### 🛡️ 安全与备份
- `excel_create_backup` - 创建文件自动备份
- `excel_restore_backup` - 从备份恢复文件
- `excel_list_backups` - 列出所有备份记录
- `excel_assess_data_impact` - 评估操作对数据的潜在影响（detailed=False快速预览/detailed=True全面评估）

### 📜 操作历史
- `excel_get_operation_history` - 获取操作历史记录和统计

### 🎨 格式化与样式
- `excel_format_cells` - 应用字体、颜色、对齐格式
- `excel_set_borders` - 设置单元格边框
- `excel_merge_cells` - 合并单元格范围
- `excel_unmerge_cells` - 取消合并单元格范围
- `excel_set_column_width` - 调整列宽
- `excel_set_row_height` - 调整行高

### 🔄 数据转换
- `excel_export_to_csv` - 导出CSV格式
- `excel_import_from_csv` - 从CSV创建Excel文件
- `excel_convert_format` - 格式转换（.xlsx/.xlsm/.csv/.json）

---

## 📖 使用指南

### 🎮 游戏配置表标准格式

**双行表头系统** (游戏开发专用，自动识别):
```
第1行(描述): ['技能ID描述', '技能名称描述', '技能类型描述']
第2行(字段): ['skill_id', 'skill_name', 'skill_type']
```

`excel_query` 会自动检测双行表头格式（第1行中文描述 + 第2行英文字段名），无需手动指定。查询结果中会附带 `column_descriptions` 映射，方便理解字段含义。

**常见配置表结构**:
- **技能配置表**: ID|名称|类型|等级|消耗|冷却|伤害|描述
- **装备配置表**: ID|名称|类型|品质|属性|套装|获取方式
- **怪物配置表**: ID|名称|等级|血量|攻击|防御|技能|掉落

### 📋 标准工作流程

1. **搜索定位**: 使用 `excel_search` 了解数据分布
2. **确定边界**: 使用 `excel_find_last_row` 确认数据范围
3. **读取现状**: 使用 `excel_get_range` 获取当前配置
4. **更新数据**: 使用 `excel_update_range` 进行安全更新
5. **美化显示**: 使用 `excel_format_cells` 标记重要数据
6. **验证结果**: 重新读取确认更新成功

### 🔍 SQL查询参考

`excel_query` 基于sqlglot + pandas实现，支持以下SQL语法：

**支持的语法：**
```sql
-- 基础查询
SELECT * FROM 技能表 WHERE level >= 10 LIMIT 20
SELECT skill_name, damage FROM 技能表 ORDER BY damage DESC

-- 聚合统计
SELECT skill_type, AVG(damage) as avg_dmg, COUNT(*) as cnt FROM 技能表 GROUP BY skill_type

-- HAVING过滤
SELECT skill_type, SUM(damage) as total FROM 技能表 GROUP BY skill_type HAVING total > 1000

-- 数学表达式
SELECT skill_name, damage * 1.2 as boosted_dmg FROM 技能表 WHERE level >= 5

-- LIKE模糊搜索
SELECT * FROM 技能表 WHERE skill_name LIKE '%火%'

-- DISTINCT去重
SELECT DISTINCT skill_type FROM 技能表

-- IN条件
SELECT * FROM 技能表 WHERE skill_type IN ('攻击', '辅助')

-- BETWEEN范围
SELECT * FROM 怪物表 WHERE level BETWEEN 10 AND 20

-- IS NULL / IS NOT NULL 空值检测
SELECT * FROM 技能表 WHERE description IS NULL
SELECT * FROM 技能表 WHERE description IS NOT NULL

-- OFFSET分页（大表分批查看）
SELECT * FROM 怪物表 ORDER BY level LIMIT 20 OFFSET 0
SELECT * FROM 怪物表 ORDER BY level LIMIT 20 OFFSET 20

-- NOT LIKE / NOT IN 排除匹配
SELECT * FROM 技能表 WHERE skill_name NOT LIKE '%测试%'
SELECT * FROM 装备表 WHERE quality NOT IN ('废弃', '内部测试')

-- JOIN 跨表关联查询（同文件内工作表）
SELECT a.skill_name, b.equip_name FROM 技能表 a INNER JOIN 装备表 b ON a.equip_id = b.equip_id
SELECT a.name, b.hp FROM 怪物表 a LEFT JOIN 怪物掉落表 b ON a.id = b.monster_id WHERE a.level > 10

-- 跨文件JOIN（@'filepath'语法，支持绝对路径和相对路径）
SELECT s.技能名, d.掉落物 FROM 技能表@'/path/to/skills.xlsx' s JOIN 掉落表@'/path/to/drops.xlsx' d ON s.技能ID = d.技能ID

-- 子查询（WHERE col IN / NOT IN / 标量子查询）
SELECT * FROM 技能表 WHERE skill_type IN (SELECT DISTINCT skill_type FROM 技能表 WHERE damage > 200)
SELECT * FROM 技能表 WHERE damage > (SELECT AVG(damage) FROM 技能表)

-- EXISTS子查询（关联子查询）
SELECT * FROM 怪物表 WHERE EXISTS (SELECT 1 FROM 掉落表 WHERE 掉落表.monster_id = 怪物表.id)

-- CASE WHEN条件表达式
SELECT skill_name, CASE WHEN damage > 200 THEN '高伤' WHEN damage > 100 THEN '中伤' ELSE '低伤' END as tier FROM 技能表

-- CTE (WITH ... AS ...)
WITH high_dmg AS (SELECT * FROM 技能表 WHERE damage > 150) SELECT * FROM high_dmg ORDER BY damage DESC
WITH mages AS (SELECT * FROM 技能表 WHERE skill_type='法师'), strong AS (SELECT * FROM mages WHERE damage >= 150) SELECT * FROM strong

-- UNION / UNION ALL 合并查询结果
SELECT name, damage FROM 技能表 WHERE skill_type='法师' UNION ALL SELECT name, damage FROM 技能表 WHERE skill_type='战士' ORDER BY damage DESC LIMIT 10
SELECT DISTINCT skill_type FROM 技能表1 UNION SELECT DISTINCT skill_type FROM 技能表2

-- COALESCE / IFNULL 空值替换
SELECT skill_name, COALESCE(description, '无描述') as desc FROM 技能表

-- 字符串函数
SELECT UPPER(skill_name) FROM 技能表 WHERE LOWER(skill_type) = 'mage'
SELECT CONCAT(skill_type, ':', skill_name) as label FROM 技能表
SELECT REPLACE(description, '攻击', '打击') FROM 技能表
SELECT skill_name, LENGTH(skill_name) as name_len FROM 技能表
SELECT SUBSTRING(skill_name, 1, 3) as short_name FROM 技能表
```

**SQL UPDATE 批量修改：**
```sql
-- 常量修改
UPDATE 技能表 SET 伤害 = 500 WHERE skill_type = '终极技能'

-- 算术表达式（列引用）
UPDATE 技能表 SET 伤害 = 伤害 * 1.1 WHERE 元素 = '火'

-- 多列修改
UPDATE 技能表 SET 伤害 = 伤害 * 1.1, 冷却 = 冷却 - 1 WHERE 等级 >= 20

-- dry_run 预览模式（不实际修改）
UPDATE 技能表 SET 伤害 = 伤害 * 1.1 WHERE 元素 = '火'  -- dry_run=True
```

**不支持的语法（有清晰替代方案提示）：**
- 嵌套FROM子查询 `FROM (SELECT ... FROM (SELECT ...))`（提示：改用单层FROM子查询或CTE）
- INSERT/DELETE语句（提示：写入请用excel_upsert_row或excel_update_query）

**FROM子查询（支持单层）：**
```sql
-- 单层FROM子查询，结果作为临时表
SELECT * FROM (SELECT skill_name, damage FROM 技能配置 WHERE damage > 100) AS 高伤技能
```

**窗口函数（ROW_NUMBER/RANK/DENSE_RANK）：**
```sql
-- 按伤害降序排名
SELECT skill_name, damage, ROW_NUMBER() OVER (ORDER BY damage DESC) as rn FROM 技能配置

-- 每个职业内按伤害排名
SELECT skill_name, skill_type, ROW_NUMBER() OVER (PARTITION BY skill_type ORDER BY damage DESC) as rn FROM 技能配置

-- RANK vs DENSE_RANK: 相同伤害并列排名
SELECT skill_name, damage, RANK() OVER (ORDER BY damage DESC) as r, DENSE_RANK() OVER (ORDER BY damage DESC) as dr FROM 技能配置
```

**双行表头自动识别：**
```
# 自动检测到双行表头后，查询结果包含column_descriptions映射
# 例：查询 "技能表" 的 skill_name 列，返回时附带 "技能名称" 描述
```

**智能列名建议：**
```
# 拼写错误时自动推荐相似列名（基于编辑距离匹配）
# 例：SELECT skil_name FROM 技能表
# → 错误：列 'skil_name' 不存在。你是否想用: skill_name?
```

**查询性能：**
- 同一文件重复查询自动缓存，提速30-100倍
- 小表(10行)：首次30-47ms，缓存后2-5ms
- 大表(2000行)：首次~230ms，缓存后2-8ms
- 文件修改后缓存自动失效

### 🎮 游戏策划SQL实战

以下是真实的游戏开发场景，直接复制粘贴就能用：

**场景1：技能数值平衡**
```sql
-- 找出DPM异常的技能（超过平均值的3倍）
SELECT skill_name, damage * 1.0 / cooldown as dpm
FROM SkillConfig
WHERE damage * 1.0 / cooldown > (SELECT AVG(damage * 1.0 / cooldown) * 3 FROM SkillConfig)

-- 各职业DPM统计，找数值失衡的职业
SELECT skill_type, AVG(damage * 1.0 / cooldown) as avg_dpm,
       MAX(damage * 1.0 / cooldown) as max_dpm,
       MIN(damage * 1.0 / cooldown) as min_dpm
FROM SkillConfig GROUP BY skill_type
```

**场景2：装备掉落配置审查**
```sql
-- 查找怪物掉落表中重复配置（同一怪物掉同一物品多次）
SELECT monster_id, item_id, COUNT(*) as dup_count
FROM DropConfig GROUP BY monster_id, item_id HAVING dup_count > 1

-- 检查掉落概率总和是否合理（不等于100%的怪物）
SELECT monster_id, SUM(drop_rate) as total_rate
FROM DropConfig GROUP BY monster_id
HAVING ABS(total_rate - 100) > 0.01
```

**场景3：配置表数据质量检查**
```sql
-- 找出所有引用了不存在ID的外键（技能引用了不存在的装备ID）
SELECT s.skill_name, s.equip_id
FROM SkillConfig s
LEFT JOIN EquipmentConfig e ON s.equip_id = e.equip_id
WHERE s.equip_id IS NOT NULL AND e.equip_id IS NULL

-- 找出数值为0或负数的异常配置
SELECT * FROM SkillConfig WHERE damage <= 0 OR cooldown < 0
```

**场景4：版本迁移辅助**
```sql
-- 找出新版本新增的配置项（用跨文件JOIN）
SELECT b.skill_name FROM 技能表@'skills_v2.xlsx' b
LEFT JOIN 技能表@'skills_v1.xlsx' a ON b.skill_id = a.skill_id
WHERE a.skill_id IS NULL

-- 找出被删除的配置项
SELECT a.skill_name FROM 技能表@'skills_v1.xlsx' a
LEFT JOIN 技能表@'skills_v2.xlsx' b ON a.skill_id = b.skill_id
WHERE b.skill_id IS NULL
```

**场景5：批量数值调整**
```sql
-- 全服平衡性调整：战士技能伤害+15%，冷却-1
UPDATE SkillConfig SET damage = damage * 1.15, cooldown = cooldown - 1 WHERE skill_type = '战士'

-- 等级压缩：把50级以上怪物的HP统一缩放到新曲线
UPDATE MonsterConfig SET hp = hp * 0.8 WHERE level > 50
```

**常见问题解决**:
- **文件被锁定**: 关闭Excel程序后重试
- **中文乱码**: 确保UTF-8编码，检查Python环境编码
- **大文件缓慢**: 使用精确范围，分批处理数据
- **内存不足**: 减少单次处理数据量，及时关闭工作簿
- **权限问题**: 使用管理员权限或检查文件属性

---

## 🔒 安全机制

ExcelMCP 内置多层安全防护，保护用户数据和系统安全：

### 路径安全（SecurityValidator）
- **路径穿越防护**: 拒绝 `../` 等目录遍历攻击
- **符号链接拒绝**: 不跟随符号链接，防止指向敏感文件
- **隐藏文件拒绝**: 不处理以 `.` 开头的隐藏文件
- **扩展名白名单**: 仅允许 `.xlsx`/`.xlsm`/`.xls`/`.csv`/`.json`/`.bak`
- **文件大小限制**: 单文件最大 50MB

### 公式注入防护
- **DDE检测**: 拒绝 `=DDE()` 等动态数据交换公式
- **CMD检测**: 拒绝 `=CMD()` 等系统命令执行
- **SHELL检测**: 拒绝 `=SHELL()` 等Shell命令公式
- **REGISTER检测**: 拒绝 `=REGISTER()` 等DLL注册公式
- **管道检测**: 拒绝包含管道符的危险公式

### 数据安全
- **文件锁**: `excel_update_query` 写入时使用文件锁（fcntl LOCK_EX），防止并发写入冲突
- **事务保护**: UPDATE操作前自动创建备份，失败自动回滚，确保文件不损坏
- **临时文件清理**: 启动时自动清理超过1小时的孤儿 `.bak` 临时文件

### 错误信息
- 安全错误以 🔒 前缀标识，包含具体被拒绝的原因
- 示例: `🔒 安全验证失败: 路径包含非法字符 '..'`

---

## 🏗️ 技术架构

### 包结构
```
src/excel_mcp_server_fastmcp/    # 主包（pip install 后可直接 import）
├── __init__.py                   # 包入口，暴露 main()
├── server.py                     # MCP接口层（52个工具定义）
├── api/                          # API业务逻辑层
│   ├── excel_operations.py       # Excel操作统一入口
│   └── advanced_sql_query.py     # SQL查询引擎
├── core/                         # 核心操作层
│   ├── excel_reader.py           # 读取操作
│   ├── excel_writer.py           # 写入操作
│   ├── excel_search.py           # 搜索操作
│   ├── excel_manager.py          # 工作簿管理
│   ├── excel_compare.py          # 版本对比
│   └── excel_converter.py        # 格式转换
├── models/                       # 数据模型
│   └── types.py                  # 类型定义
└── utils/                        # 工具层
    ├── validators.py             # 路径/数据验证 + 安全防护
    ├── error_handler.py          # 统一错误处理
    ├── formatter.py              # 结果格式化
    ├── parsers.py                # 参数解析
    ├── temp_file_manager.py      # 临时文件管理
    ├── formula_cache.py          # 公式缓存
    └── exceptions.py             # 自定义异常
```

### 分层设计模式
```
MCP接口层 (纯委托)
    ↓
API业务逻辑层 (集中式处理)
    ↓
核心操作层 (单一职责)
    ↓
工具层 (通用功能)
```

### 核心特性
- **纯委托模式**: 接口层零业务逻辑，全部委托
- **集中式处理**: 统一验证、错误处理、结果格式化
- **1-Based索引**: 匹配Excel用户习惯 (第1行=第一行)
- **工作簿缓存**: 缓存命中时性能提升75%
- **现实并发处理**: 正确处理Excel文件并发限制

### 性能优化
- **python-calamine读取引擎**: Rust原生解析，get_range从1.6s降至0.7ms（2300x提速）
- **精确范围读取**: 比整表读取快60-80%
- **批量操作**: 比逐个操作快15-20倍
- **分批处理**: 大文件内存占用降低70%

---

## 📊 项目信息

### 质量验证指标
- **测试用例**: 1159个（行为验证，无覆盖率填充）
- **测试文件**: 49个测试文件
- **测试代码**: 16,496行
- **工具数量**: 44个 (@mcp.tool装饰器验证)
- **架构层次**: 4层分层设计 (MCP→API→Core→Utils)

### 验证命令
```bash
# 运行完整测试套件（并行加速）
python -m pytest tests/ -q --tb=short -n auto --timeout=30

# 验证工具完整性
grep -c "def excel_" src/excel_mcp_server_fastmcp/server.py  # 应输出: 44

# 生成覆盖率报告
python -m pytest tests/ --cov=src --cov-report=html
```

### 开发规范
- **纯委托模式**: server.py严格委托给ExcelOperations
- **集中式业务逻辑**: 统一验证、错误处理、结果格式化
- **分支命名**: 所有功能分支必须以`feature/`开头
- **测试覆盖**: 保持80%以上的测试覆盖率

---

## ❓ 常见问题

### 基础问题
**Q: 支持哪些Excel格式？**
A: 支持`.xlsx`、`.xlsm`格式，通过导入导出支持`.csv`格式

**Q: 如何处理中文工作表名？**
A: 完全支持中文工作表名和内容

**Q: 大文件处理性能如何？**
A: SQL查询自动缓存DataFrame，同一文件重复查询提速30-100倍。大表(2000行)首次~230ms，缓存后2-8ms。

**Q: 如何确保数据安全？**
A: 完整错误处理，默认保留公式，支持操作预览

### 游戏开发专用
**Q: 什么是双行表头系统？**
A: 游戏配置表标准格式：第1行字段描述，第2行字段名

**Q: 如何进行版本对比？**
A: 使用专门的配置表对比工具，支持ID对象跟踪

---

## 🤝 贡献指南

**贡献方式**:
- 🐛 **报告Bug**: 通过GitHub Issues报告问题
- 💡 **功能建议**: 提出新功能需求
- 📝 **文档改进**: 完善使用指南和技术文档
- 🔧 **代码贡献**: 遵循开发规范，提交PR

**许可证**: MIT License - 详见 [LICENSE](LICENSE) 文件

---

<div align="center">

## 📋 更新日志

查看 [CHANGELOG.md](CHANGELOG.md) 了解各版本详细变更记录。

## 🔝 快速导航

| 🎯 **快速开始** | 🛠️ **工具参考** | 📚 **学习指南** |
|----------------|----------------|----------------|
| [🚀 安装配置](#-快速入门-3分钟设置) | [📋 完整工具列表](#️-完整工具列表44个专业工具) | [📖 使用指南](#-使用指南) |
| [⚡ 命令速查](#-快速参考) | [🏗️ 技术架构](#️-技术架构) | [🚨 故障排除](#-故障排除) |
| [🎮 游戏配置管理](#-使用指南) | [📊 项目信息](#-项目信息) | [❓ 常见问题](#-常见问题) |

**[⬆️ 返回顶部](#-excelmcp-游戏开发专用-excel-配置表管理器)**

*✨ 让游戏配置表管理变得简单高效 ✨*

</div>

## Installation

```bash
pip install excel-mcp-server-fastmcp
```

## Usage

```python
from excel_mcp_server_fastmcp import main
main()
```

## 🎮 使用示例

查看 `examples/` 目录获取完整的使用场景，包括：

- **基础操作**：技能表创建、装备配置管理、怪物属性设置
- **进阶操作**：跨文件JOIN查询、批量数据更新、版本对比与回滚
- **实战案例**：技能系统完整设计、数值平衡调整

每个示例都包含完整的Python代码和详细的注释，帮助您快速上手ExcelMCP的核心功能。

## 快速示例
```python
# 创建技能配置表
await excel_create_worksheet(
    filepath="skills.xlsx",
    sheet_name="skills",
    headers=["skill_id", "skill_name", "damage", "cooldown"]
)

# 批量插入技能数据
skills_data = [
    [1, "火球术", 150, 3.0],
    [2, "冰冻术", 120, 4.0]
]
await excel_write_rows(
    filepath="skills.xlsx",
    sheet_name="skills", 
    data=skills_data,
    start_cell="A2"
)
```

**项目状态**
- **版本**: v1.6.29（第169轮文档优化完成）
- **测试**: 1164个测试用例通过
- **工具**: 44个专业工具，全部docstring优化完成
- **更新**: 第154轮工具描述优化成果已记录

## 🔧 常见问题与故障排除

### 安装问题
**问题**: `uvx excel-mcp-server-fastmcp` 命令不存在
- **解决**: 确保已安装uv: `curl -LsSf https://astral.sh/uv/install.sh | sh`
- **验证**: `uvx --version` 应显示版本号

**问题**: MCP客户端连接失败
- **检查**: 确保配置文件JSON格式正确，没有多余逗号
- **验证**: 手动运行 `uvx excel-mcp-server-fastmcp` 查看是否正常启动

### 使用问题
**问题**: 中文查询没有结果
- **原因**: 可能是列名不匹配或文件编码问题
- **解决**: 
  1. 使用 `excel_list_sheets` 确认表名
  2. 使用 `excel_get_headers` 确认列名
  3. 检查文件是否为UTF-8编码

**问题**: 跨文件JOIN失败
- **原因**: 文件路径或语法错误
- **解决**: 
  1. 确认文件路径正确，使用 `@'filename.xlsx'` 语法
  2. 检查JOIN条件是否匹配
  3. 使用简单JOIN测试，逐步复杂化

### 性能问题
**问题**: 大文件操作很慢
- **解决**: 启用流式写入 `streaming=True` 参数
- **示例**: 
  ```python
  await excel_write_rows(
      filepath="large_file.xlsx",
      data=large_data,
      streaming=True  # 内存优化模式
  )
  ```

### 获取帮助
- **文档**: 查看 [详细使用指南](#-使用指南)
- **示例**: 参考 [examples/](examples/) 目录中的完整示例
- **问题**: 提交 [GitHub Issues](https://github.com/TangentDomain/excel-mcp-server/issues)