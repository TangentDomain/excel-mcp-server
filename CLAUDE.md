# CLAUDE.md

本文件为 Claude Code (claude.ai/code) 在此代码库中工作时提供指导。

## 项目概览

ExcelMCP 是专为游戏开发设计的 Excel 配置表管理 MCP (Model Context Protocol) 服务器。提供 32 个专业工具管理 Excel 文件，配备 295 个测试用例确保 100% 覆盖率和企业级可靠性。

### 核心用途
- **游戏开发专业化**: 专精于技能配置表、装备数据、怪物属性和游戏配置管理
- **AI-自然语言接口**: 让 AI 助手通过自然语言命令控制 Excel 文件
- **Excel 配置管理**: 处理复杂的游戏数据结构，支持双行表头（描述 + 字段名）

## 架构

### 分层架构模式
```
MCP 接口层 (src/server.py)
    ↓ 委托给
API 业务逻辑层 (src/api/excel_operations.py)
    ↓ 使用
核心操作层 (src/core/*)
    ↓ 使用
工具层 (src/utils/*)
```

### 核心设计原则
1. **纯委托模式**: `server.py` 中的 MCP 工具将所有业务逻辑委托给 `ExcelOperations`
2. **集中式业务逻辑**: `ExcelOperations` 类处理参数验证、业务逻辑、错误处理和结果格式化
3. **标准化结果**: 所有操作返回 `{success, data, message, metadata}` 结构
4. **1-Based 索引**: 匹配 Excel 约定（第1行 = 第一行，A列 = 1）

### 范围表达式系统
- **标准格式**: `"Sheet1!A1:C10"`（必须包含工作表名）
- **行范围**: `"Sheet1!1:5"`（第1-5行）
- **列范围**: `"Sheet1!B:D"`（B-D列）
- **单元素**: `"Sheet1!5"` 或 `"Sheet1!C"`（单行/单列）

## 开发工作流

### 运行测试
```bash
# 运行所有测试并生成覆盖率报告
python scripts/run_tests.py

# 运行特定测试模块
pytest tests/test_api_excel_operations.py -v
pytest tests/test_core.py -v
pytest tests/test_server.py -v

# 运行详细输出
pytest tests/ -v --tb=short

# 运行特定功能的测试
pytest tests/ -k "test_get_range" -v
```

### 测试结构
- **API 测试**: `test_api_excel_operations.py` - 使用 Mock 隔离测试业务逻辑
- **核心测试**: `test_core.py` - 测试 Excel 操作模块
- **MCP 测试**: `test_server.py` - 测试 MCP 接口委托
- **功能测试**: 各种功能特定的测试文件
- **配置**: `conftest.py` 提供临时 Excel 文件的 fixtures

### MCP 客户端配置
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

## 开发标准

### 代码组织规则
1. **MCP 接口**: `src/server.py` 仅包含 MCP 工具定义，零业务逻辑
2. **API 层**: `src/api/excel_operations.py` 处理集中式业务逻辑，包含全面验证
3. **核心模块**: `src/core/` 模块处理单一职责的 Excel 操作
4. **工具函数**: `src/utils/` 用于格式化助手和通用工具

### 方法复杂度指南
- **主干方法**: ≤20 行，确保可读性
- **分支方法**: ≤50 行，专注特定功能
- **命名规范**: 清晰的动词+名词组合，避免 handle/process

### Excel 约定
- **1-based 索引**: 匹配 Excel 行/列编号
- **默认行为**: 保留公式（`preserve_formulas=True`）
- **游戏特定**: 双行表头（描述 + 字段名）
- **文件格式**: 支持 `.xlsx` 和 `.xlsm` 格式

### 性能模式
- **工作簿缓存**: 避免重复加载大型 Excel 文件
- **精确范围**: 指定具体单元格范围，避免全表读取
- **批量操作**: 优先批量更新而非单单元格操作

## 工具分类

### 32 个专业工具
1. **文件和工作表管理** (8个工具): create_file, list_sheets, create_sheet, delete_sheet, rename_sheet 等
2. **数据操作** (8个工具): get_range, update_range, insert_rows, delete_rows, insert_columns, delete_columns, find_last_row, get_headers
3. **搜索和分析** (4个工具): search, search_directory, get_headers, get_sheet_headers
4. **格式化和样式** (6个工具): format_cells, merge_cells, unmerge_cells, set_borders, set_row_height, set_column_width
5. **导入导出和转换** (3个工具): export_to_csv, import_from_csv, convert_format, merge_files

### 游戏开发专业化
- **技能表**: `TrSkill` 结构，包含 ID|名称|类型|等级|消耗|冷却|伤害|描述
- **装备表**: `TrItem` 结构，包含 ID|名称|类型|品质|属性|套装|获取方式
- **怪物表**: `TrMonster` 结构，包含 ID|名称|等级|血量|攻击|防御|技能|掉落
- **配置对比**: 专业化版本对比，支持 ID 对象跟踪

## 主要依赖
- **FastMCP**: MCP 服务器框架
- **openpyxl**: 核心 Excel 文件操作
- **xlcalculator/formulas**: 公式计算引擎
- **xlwings**: 可选的 Excel 应用集成
- **pytest/pytest-asyncio**: 测试框架

## 常用操作

### Excel 配置表工作流
1. **搜索定位**: 使用 `excel_search` 了解数据分布
2. **确定边界**: 使用 `excel_find_last_row` 确认数据范围
3. **读取现状**: 使用 `excel_get_range` 获取当前配置
4. **更新数据**: 使用 `excel_update_range` 进行安全更新
5. **美化显示**: 使用 `excel_format_cells` 标记重要数据
6. **验证结果**: 重新读取确认更新成功

### 版本对比
```python
excel_compare_sheets("旧配置.xlsx", "TrSkill", "新配置.xlsx", "TrSkill")
```

### ID 验证
```python
excel_check_duplicate_ids("技能表.xlsx", "技能配置表", id_column=1)
```

## 错误处理

### 常见问题
- **文件被锁定**: 检查 Excel 是否打开，关闭后重试
- **权限不足**: 使用管理员权限或检查文件属性
- **范围越界**: 先用 `excel_find_last_row` 确认实际数据范围
- **中文乱码**: 确认编码格式，使用 utf-8
- **公式错误**: 设置 `preserve_formulas=False` 强制覆盖
- **内存不足**: 分批处理大文件

### 日志级别
- DEBUG: 详细操作日志
- INFO: 操作摘要
- ERROR: 异常详情和上下文

## 文件组织

### 源代码结构
```
src/
├── server.py              # 仅 MCP 接口（委托模式）
├── api/
│   └── excel_operations.py # 集中式业务逻辑
├── core/
│   ├── excel_reader.py    # 读取操作
│   ├── excel_writer.py    # 写入操作
│   ├── excel_manager.py   # 文件/工作表管理
│   ├── excel_search.py    # 搜索功能
│   ├── excel_compare.py   # 对比操作
│   └── excel_converter.py # 格式转换
├── utils/
│   └── formatter.py       # 结果格式化
└── models/
    └── types.py          # 类型定义
```

### 测试结构
```
tests/
├── conftest.py           # 测试 fixtures 和配置
├── test_api_excel_operations.py  # API 层测试
├── test_core.py          # 核心操作测试
├── test_server.py        # MCP 接口测试
└── [各种功能测试] # 特定功能测试
```

## 中文/Unicode 支持
- 完整支持中文字符的工作表名
- 双行表头系统（描述 + 字段名）
- Unicode 文本处理和标准化
- 缺失表头数据的回退机制
- 本地化 Excel 功能处理