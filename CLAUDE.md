<!-- OPENSPEC:START -->
# OpenSpec Instructions

These instructions are for AI assistants working in this project.

Always open `@/openspec/AGENTS.md` when the request:
- Mentions planning or proposals (words like proposal, spec, change, plan)
- Introduces new capabilities, breaking changes, architecture shifts, or big performance/security work
- Sounds ambiguous and you need the authoritative spec before coding

Use `@/openspec/AGENTS.md` to learn:
- How to create and apply change proposals
- Spec format and conventions
- Project structure and guidelines

Keep this managed block so 'openspec update' can refresh the instructions.

<!-- OPENSPEC:END -->

# CLAUDE.md

本文件为 Claude Code (claude.ai/code) 在此代码库中工作时提供指导。

## 项目概览

ExcelMCP 是专为游戏开发设计的 Excel 配置表管理 MCP (Model Context Protocol) 服务器。提供 30 个专业工具管理 Excel 文件，配备 698 个测试用例确保高质量覆盖和企业级可靠性。

### 核心用途
- **游戏开发专业化**: 专精于技能配置表、装备数据、怪物属性和游戏配置管理
- **AI-自然语言接口**: 让 AI 助手通过自然语言命令控制 Excel 文件
- **Excel 配置管理**: 处理复杂的游戏数据结构，支持双行表头（描述 + 字段名）

## 架构设计

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

### MCP 接口架构设计理念

#### 1. 纯委托模式 (Pure Delegation Pattern)
`src/server.py` 作为 MCP 接口层，严格遵循**纯委托模式**：

```python
# ❌ 错误：在MCP接口层实现业务逻辑
@mcp.tool()
def excel_get_range(file_path: str, range: str):
    workbook = openpyxl.load_workbook(file_path)  # 不应该在此实现

# ✅ 正确：纯委托给业务逻辑层
@mcp.tool()
def excel_get_range(file_path: str, range: str):
    return self.excel_ops.get_range(file_path, range)
```

**设计意图**：
- **接口纯净性**: MCP 层仅负责接口定义和参数传递，零业务逻辑
- **职责分离**: 业务逻辑、验证、错误处理全部集中在 `ExcelOperations` 类
- **可测试性**: 接口层和业务逻辑层可以独立测试

#### 2. 集中式业务逻辑处理
`ExcelOperations` 类作为**单一业务逻辑入口**，统一处理：
- 参数验证
- 业务逻辑执行
- 结果格式化
- 错误处理

**核心优势**：
- **统一验证**: 所有输入参数都经过相同的验证逻辑
- **统一错误处理**: 标准化的异常处理和错误响应
- **统一格式**: 所有操作返回相同结构的 `OperationResult`

#### 3. 现实并发操作处理
**Excel 文件不支持真正的并发写入**，我们采用现实的方法：
- **错误检测**: 正确识别并发操作导致的文件锁定
- **优雅降级**: 提供序列化操作队列作为备选方案
- **性能平衡**: 允许并发读取，序列化写入操作

#### 4. 系统化测试修复方法论
从 24 个失败测试到 100% 通过的**系统性方法**：
- **根因分析**: 按照错误根本原因分类（编码、公式、并发、数据格式）
- **渐进修复**: 优先修复影响最大的基础性问题
- **验证闭环**: 每个修复都包含完整的测试验证

### 核心设计原则
1. **纯委托模式**: `server.py` 中的 MCP 工具将所有业务逻辑委托给 `ExcelOperations`
2. **集中式业务逻辑**: `ExcelOperations` 类处理参数验证、业务逻辑、错误处理和结果格式化
3. **标准化结果**: 所有操作返回 `{success, data, message, metadata}` 结构
4. **1-Based 索引**: 匹配 Excel 约定（第1行 = 第一行，A列 = 1）
5. **现实并发处理**: 正确处理 Excel 文件的并发限制，提供序列化解决方案
6. **系统性测试**: 基于根因分析的测试修复方法论

### 范围表达式系统
- **标准格式**: `"Sheet1!A1:C10"`（必须包含工作表名）
- **行范围**: `"Sheet1!1:5"`（第1-5行）
- **列范围**: `"Sheet1!B:D"`（B-D列）
- **单元素**: `"Sheet1!5"` 或 `"Sheet1!C"`（单行/单列）

## 开发工作流

### 测试运行
```bash
# 推荐方式：使用Python模块方式运行测试
python -m pytest tests/ -v

# 生成覆盖率报告
python -m pytest tests/ --cov=src --cov-report=html --cov-report=term

# 运行特定测试
python -m pytest tests/test_api_excel_operations.py -v
python -m pytest tests/test_core.py -v
python -m pytest tests/test_server.py -v

# 运行单个测试方法
python -m pytest tests/test_api_excel_operations.py::TestExcelOperations::test_get_range_success_flow -v -s
```

**注意**: 推荐使用 `python -m pytest` 而不是直接 `pytest`，这样可以避免Python路径问题。

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

### 架构规范执行

#### 严格遵循 CLAUDE.md 规范
当发现可能违反现有架构规范的操作时：

1. **识别违规**: 检测到可能违反分层架构或设计原则的操作
2. **明确提示**: 立即提示用户当前操作可能违反既定规范
3. **提供选择**:
   - 方案A：调整操作以符合现有规范
   - 方案B：如果确实需要打破规范，则更新 CLAUDE.md 文档

#### 规范更新流程
当架构演进需要打破现有规范时：

**更新原则**：
- **向后兼容**: 尽量保持现有规范的兼容性
- **充分讨论**: 重大架构变更需要充分论证
- **文档同步**: 代码实现与文档规范同步更新
- **测试验证**: 新规范需要相应的测试用例验证

### 开发约定
- **方法复杂度**: 主干方法 ≤20 行，分支方法 ≤50 行
- **命名规范**: 清晰的动词+名词组合，避免 handle/process
- **Excel 约定**: 1-based 索引、默认保留公式、支持双行表头、.xlsx/.xlsm 格式
- **性能模式**: 工作簿缓存、精确范围读取、批量操作优先

## 工具分类

### 30 个专业工具 (已启用)
1. **文件和工作表管理** (8个工具):
   - `excel_list_sheets` - 列出工作表
   - `excel_get_file_info` - 获取文件信息
   - `excel_create_file` - 创建新文件
   - `excel_create_sheet` - 创建工作表
   - `excel_delete_sheet` - 删除工作表
   - `excel_rename_sheet` - 重命名工作表
   - `excel_get_sheet_headers` - 获取所有工作表表头

2. **数据操作** (8个工具):
   - `excel_get_range` - 读取数据范围
   - `excel_update_range` - 更新数据范围
   - `excel_get_headers` - 获取表头信息
   - `excel_insert_rows` - 插入行
   - `excel_delete_rows` - 删除行
   - `excel_insert_columns` - 插入列
   - `excel_delete_columns` - 删除列
   - `excel_find_last_row` - 查找最后一行

3. **搜索和分析** (4个工具):
   - `excel_search` - 单文件搜索
   - `excel_search_directory` - 目录批量搜索
   - `excel_check_duplicate_ids` - ID重复检测
   - `excel_compare_sheets` - 工作表对比

4. **格式化和样式** (6个工具):
   - `excel_format_cells` - 单元格格式化
   - `excel_merge_cells` - 合并单元格
   - `excel_unmerge_cells` - 取消合并
   - `excel_set_borders` - 设置边框
   - `excel_set_row_height` - 调整行高
   - `excel_set_column_width` - 调整列宽

5. **导入导出和转换** (4个工具):
   - `excel_export_to_csv` - 导出CSV
   - `excel_import_from_csv` - 导入CSV
   - `excel_convert_format` - 格式转换
   - `excel_merge_files` - 文件合并

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

### 游戏配置表示例

#### 技能配置表批量操作
```python
# 搜索特定技能类型
search_result = excel_search(
    file_path="skills.xlsx",
    pattern=r"火系|冰系|雷系",
    sheet_name="技能配置表",
    use_regex=True
)

# 批量调整伤害值
damage_data = excel_get_range("skills.xlsx", "技能配置表!G2:G100")
if damage_data['success']:
    updated_damage = [[row[0] * 1.2] if row and isinstance(row[0], (int, float)) else row
                     for row in damage_data['data']]

    excel_update_range("skills.xlsx", "技能配置表!G2:G100", updated_damage)
    excel_format_cells("skills.xlsx", "技能配置表", "G2:G100", preset="highlight")
```

#### 安全文件操作函数
```python
def safe_update_config(file_path, sheet_name, range_expr, new_data):
    """安全的配置表更新函数"""
    # 检查文件存在性、工作表有效性、数据格式
    # 执行更新并验证结果
    # 完整的错误处理和资源清理
```

## 文件组织

### 核心目录结构
```
excel-mcp-server/
├── src/                        # 源代码
│   ├── server.py               # MCP服务器入口（仅接口定义）
│   ├── api/excel_operations.py # 集中式业务逻辑
│   ├── core/                   # 核心Excel操作
│   ├── utils/                  # 工具函数
│   └── models/                 # 数据模型
├── tests/                      # 测试用例（698个测试）
├── scripts/                    # 构建脚本
├── docs/                       # 项目文档
└── pyproject.toml             # 项目配置
```

### 文件命名约定
- **核心模块**: `excel_reader.py`, `excel_writer.py`, `excel_search.py` 等
- **工具模块**: `formatter.py`, `validators.py`, `exceptions.py` 等
- **测试文件**: `test_[功能名].py` 格式

## 环境配置

### 快速设置
```bash
# 使用 uv (推荐)
pip install uv
git clone https://github.com/tangjian/excel-mcp-server.git
cd excel-mcp-server
uv sync
source .venv/bin/activate  # Linux/macOS 或 .venv\Scripts\activate (Windows)
python -m pytest tests/  # 验证环境

# 使用 pip (备选)
python -m venv .venv
source .venv/bin/activate
pip install -e .
python scripts/run_tests.py
```

### IDE 配置要点
- **VS Code**: 设置 Python 解释器为 `./.venv/bin/python`
- **PyCharm**: 标记 `src` 为源代码根目录，配置 pytest 运行器
- **环境变量**: `PYTHONPATH="${PWD}/src:${PYTHONPATH}"`

## 部署指南

### Docker 部署
```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY pyproject.toml uv.lock ./
COPY src/ ./src/
RUN pip install uv && uv sync --frozen
CMD ["uv", "run", "python", "-m", "src.server"]
```

### 系统服务
```bash
# Linux systemd 服务配置
sudo nano /etc/systemd/system/excelmcp.service
# 配置工作目录、环境变量、启动命令
sudo systemctl enable excelmcp && sudo systemctl start excelmcp
```

## 错误处理

### 常见问题及解决方案
- **文件被锁定**: 关闭 Excel 进程后重试
- **权限不足**: 使用管理员权限或检查文件属性
- **范围越界**: 先用 `excel_find_last_row` 确认数据范围
- **中文乱码**: 确认使用 utf-8 编码
- **公式错误**: 设置 `preserve_formulas=False` 强制覆盖
- **内存不足**: 分批处理大文件

### 性能优化
- **工作簿缓存**: 避免重复加载大型文件
- **精确范围**: 指定具体单元格范围
- **批量操作**: 优先批量更新而非单单元格操作
- **分批处理**: 大文件分批读取，避免内存溢出

## 中文支持特性
- 完整支持中文字符的工作表名
- 双行表头系统（描述 + 字段名）
- Unicode 文本处理和标准化
- 缺失表头数据的回退机制
- 本地化 Excel 功能处理

---

**规范遵循**: 本文档严格定义项目的架构设计和开发规范。当需要打破现有规范时，必须更新本文档并确保新规范的一致性和测试覆盖。

**动态更新**: 随着项目演进，CLAUDE.md 文档需要同步更新以反映最新的架构决策和最佳实践。