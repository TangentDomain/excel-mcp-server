<div align="center">

[中文](README.md) ｜ [English](README.en.md)

</div>

# 🎮 ExcelMCP: 游戏开发专用 Excel 配置表管理器

[![许可证: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 版本](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![技术支持: FastMCP](https://img.shields.io/badge/Powered%20by-FastMCP-orange)](https://github.com/jlowin/fastmcp)
![状态](https://img.shields.io/badge/status-stable-green.svg)
![测试覆盖](https://img.shields.io/badge/tests-698%20passed-brightgreen.svg)
![覆盖率](https://img.shields.io/badge/coverage-78.58%25-blue.svg)
![工具数量](https://img.shields.io/badge/tools-38%20verified%20tools-green.svg)

**ExcelMCP** 是专为游戏开发设计的Excel配置表管理MCP服务器。通过AI自然语言指令，实现技能配置表、装备数据、怪物属性等游戏配置的智能化操作。基于**FastMCP**和**openpyxl**构建，拥有**38个专业工具**和**698个测试用例**，确保企业级可靠性。

🎯 **核心功能**: 技能系统、装备管理、怪物配置、数值平衡、版本对比、策划工具链

---

## 🚀 快速入门 (3分钟设置)

### 安装步骤

1. **克隆项目**
   ```bash
   git clone https://github.com/tangjian/excel-mcp-server.git
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
         "args": ["-m", "src.server"],
         "env": {"PYTHONPATH": "${workspaceRoot}"}
       }
     }
   }
   ```

4. **开始使用**
   准备就绪！让AI助手通过自然语言控制Excel文件。

### 验证安装
```bash
python -m pytest tests/ --tb=short -q
```

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
| 版本对比分析 | `excel_compare_sheets` | "对比新旧版本配置表差异" |

### 🔧 范围表达式参考

| 格式 | 说明 | 示例 |
|------|------|------|
| `Sheet1!A1:C10` | 标准范围 | "技能表!A1:D50" |
| `Sheet1!1:5` | 行范围 | "配置表!2:100" |
| `Sheet1!B:D` | 列范围 | "数据表!B:G" |
| `Sheet1!A1` | 单单元格 | "设置表!A1" |

---

## 🛠️ 完整工具列表（38个专业工具）

### 📁 文件与工作表管理
- `excel_create_file` - 创建新Excel文件，支持自定义工作表
- `excel_create_sheet` - 添加新工作表
- `excel_delete_sheet` - 删除工作表
- `excel_list_sheets` - 列出工作表名称
- `excel_rename_sheet` - 重命名工作表
- `excel_get_file_info` - 获取文件元数据
- `excel_get_sheet_headers` - 获取所有工作表表头
- `excel_merge_files` - 合并多个Excel文件

### 📊 数据操作
- `excel_get_range` - 读取单元格/行/列范围
- `excel_update_range` - 写入/更新数据范围，支持公式保留
- `excel_get_headers` - 从任意行提取表头
- `excel_insert_rows` - 插入空行
- `excel_delete_rows` - 删除行范围
- `excel_insert_columns` - 插入空列
- `excel_delete_columns` - 删除列范围
- `excel_find_last_row` - 查找最后一行有数据位置

### 🔍 搜索与分析
- `excel_search` - 正则表达式搜索
- `excel_search_directory` - 目录批量搜索
- `excel_compare_sheets` - 工作表对比（游戏配置优化）
- `excel_check_duplicate_ids` - ID重复检测

### 🎨 格式化与样式
- `excel_format_cells` - 应用字体、颜色、对齐格式
- `excel_set_borders` - 设置单元格边框
- `excel_merge_cells` - 合并单元格范围
- `excel_unmerge_cells` - 取消合并单元格
- `excel_set_column_width` - 调整列宽
- `excel_set_row_height` - 调整行高

### 🔄 数据转换
- `excel_export_to_csv` - 导出CSV格式
- `excel_import_from_csv` - 从CSV创建Excel文件
- `excel_convert_format` - 格式转换（.xlsx/.xlsm/.csv/.json）

---

## 📖 使用指南

### 🎮 游戏配置表标准格式

**双行表头系统** (游戏开发专用):
```
第1行(描述): ['技能ID描述', '技能名称描述', '技能类型描述']
第2行(字段): ['skill_id', 'skill_name', 'skill_type']
```

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

### 🚨 故障排除

**常见问题解决**:
- **文件被锁定**: 关闭Excel程序后重试
- **中文乱码**: 确保UTF-8编码，检查Python环境编码
- **大文件缓慢**: 使用精确范围，分批处理数据
- **内存不足**: 减少单次处理数据量，及时关闭工作簿
- **权限问题**: 使用管理员权限或检查文件属性

---

## 🏗️ 技术架构

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
- **精确范围读取**: 比整表读取快60-80%
- **批量操作**: 比逐个操作快15-20倍
- **分批处理**: 大文件内存占用降低70%

---

## 📊 项目信息

### 质量验证指标
- **测试用例**: 699个 (698通过, 1跳过)
- **测试代码**: 13,515行 (全面验证)
- **工具数量**: 38个 (@mcp.tool装饰器验证)
- **测试覆盖**: 78.58%
- **架构层次**: 4层分层设计 (MCP→API→Core→Utils)

### 验证命令
```bash
# 运行完整测试套件
python -m pytest tests/ -v

# 验证工具完整性
grep -r "@mcp.tool" src/ | wc -l  # 应输出: 38

# 生成覆盖率报告
python -m pytest tests/ --cov=src --cov-report=html
```

### 开发规范
- **纯委托模式**: server.py严格委托给ExcelOperations
- **集中式业务逻辑**: 统一验证、错误处理、结果格式化
- **分支命名**: 所有功能分支必须以`feature/`开头
- **测试覆盖**: 保持78%以上的测试覆盖率

---

## ❓ 常见问题

### 基础问题
**Q: 支持哪些Excel格式？**
A: 支持`.xlsx`、`.xlsm`格式，通过导入导出支持`.csv`格式

**Q: 如何处理中文工作表名？**
A: 完全支持中文工作表名和内容

**Q: 大文件处理性能如何？**
A: 基于openpyxl性能，建议对大文件进行分批处理

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

## 🔝 快速导航

| 🎯 **快速开始** | 🛠️ **工具参考** | 📚 **学习指南** |
|----------------|----------------|----------------|
| [🚀 安装配置](#-快速入门-3分钟设置) | [📋 完整工具列表](#️-完整工具列表38个专业工具) | [📖 使用指南](#-使用指南) |
| [⚡ 命令速查](#-快速参考) | [🏗️ 技术架构](#️-技术架构) | [🚨 故障排除](#-故障排除) |
| [🎮 游戏配置管理](#-使用指南) | [📊 项目信息](#-项目信息) | [❓ 常见问题](#-常见问题) |

**[⬆️ 返回顶部](#-excelmcp-游戏开发专用-excel-配置表管理器)**

*✨ 让游戏配置表管理变得简单高效 ✨*

</div>