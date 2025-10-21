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

ExcelMCP 是专为游戏开发设计的 Excel 配置表管理 MCP (Model Context Protocol) 服务器。提供 30 个专业工具管理 Excel 文件，配备 289 个测试用例确保高质量覆盖和企业级可靠性。

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

### MCP 接口架构设计理念

#### 1. 纯委托模式 (Pure Delegation Pattern)
`src/server.py` 作为 MCP 接口层，严格遵循**纯委托模式**：

```python
# ❌ 错误：在MCP接口层实现业务逻辑
@mcp.tool()
def excel_get_range(file_path: str, range: str):
    # 不应该在这里实现读取逻辑
    workbook = openpyxl.load_workbook(file_path)
    # ... 复杂的业务逻辑
    return result

# ✅ 正确：纯委托给业务逻辑层
@mcp.tool()
def excel_get_range(file_path: str, range: str):
    return self.excel_ops.get_range(file_path, range)
```

**设计意图**：
- **接口纯净性**: MCP 层仅负责接口定义和参数传递，零业务逻辑
- **职责分离**: 业务逻辑、验证、错误处理全部集中在 `ExcelOperations` 类
- **可测试性**: 接口层和业务逻辑层可以独立测试
- **维护性**: 修改业务逻辑无需触及 MCP 接口定义

#### 2. 集中式业务逻辑处理
`ExcelOperations` 类作为**单一业务逻辑入口**：

```python
class ExcelOperations:
    def get_range(self, file_path: str, range: str):
        # 步骤1：参数验证
        validated_params = self._validate_get_range_params(file_path, range)

        # 步骤2：业务逻辑执行
        result = self.reader.get_range(validated_params)

        # 步骤3：结果格式化
        formatted_result = format_operation_result(result)

        # 步骤4：统一响应结构
        return formatted_result
```

**核心优势**：
- **统一验证**: 所有输入参数都经过相同的验证逻辑
- **统一错误处理**: 标准化的异常处理和错误响应
- **统一格式**: 所有操作返回相同结构的 `OperationResult`
- **审计能力**: 集中的日志记录和性能监控

#### 3. 现实并发操作处理
**Excel 文件不支持真正的并发写入**，我们的架构采用现实的方法：

```python
# ❌ 不现实：期望并发Excel写入成功
def test_concurrent_excel_writes():
    # 多个线程同时写入同一Excel文件 - 必然失败

# ✅ 现实：检测错误并提供序列化方案
def test_concurrent_write_error_handling():
    # 验证系统能正确检测文件锁定错误

def test_sequential_operations_with_thread_safety():
    # 使用队列序列化写入操作，确保安全
```

**架构原则**：
- **错误检测**: 正确识别并发操作导致的文件锁定
- **优雅降级**: 提供序列化操作队列作为备选方案
- **用户友好**: 明确告知用户 Excel 文件的并发限制
- **性能平衡**: 允许并发读取，序列化写入操作

#### 4. 系统化测试修复方法论
从 24 个失败测试到 100% 通过的**系统性方法**：

```python
# 测试修复分类处理
def analyze_test_failures(failures):
    categories = {
        'encoding_issues': [],      # GBK编码问题
        'formula_calculation': [],  # 公式计算类型检测
        'concurrent_operations': [], # 并发操作处理
        'data_format_inconsistency': [] # 数据格式不一致
    }
    # 分类 → 分析 → 修复 → 验证
```

**修复原则**：
- **根因分析**: 按照错误根本原因分类，而非表面现象
- **渐进修复**: 优先修复影响最大的基础性问题
- **向后兼容**: 修复时保持 API 兼容性
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

### 运行测试
```bash
# 推荐方式：使用Python模块方式运行测试（解决路径问题）
python -m pytest tests/ -v

# 运行所有测试并生成覆盖率报告
python -m pytest tests/ --cov=src --cov-report=html --cov-report=term

# 运行特定测试模块
python -m pytest tests/test_api_excel_operations.py -v
python -m pytest tests/test_core.py -v
python -m pytest tests/test_server.py -v

# 运行详细输出
python -m pytest tests/ -v --tb=short

# 运行特定功能的测试
python -m pytest tests/ -k "test_get_range" -v

# 运行单个测试方法
python -m pytest tests/test_api_excel_operations.py::TestExcelOperations::test_get_range_success_flow -v -s

# 传统方式（如果有路径问题）
python scripts/run_tests.py
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

```python
# 示例：需要添加新的架构层次
def propose_architecture_change():
    """
    规范更新提案流程：
    1. 识别当前规范限制
    2. 提出新架构方案
    3. 分析影响范围
    4. 更新 CLAUDE.md 规范
    5. 验证新规范一致性
    """
```

**更新原则**：
- **向后兼容**: 尽量保持现有规范的兼容性
- **充分讨论**: 重大架构变更需要充分论证
- **文档同步**: 代码实现与文档规范同步更新
- **测试验证**: 新规范需要相应的测试用例验证

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

### 备用工具 (开发中)
- `excel_set_formula` - 设置公式 (已注释)
- `excel_evaluate_formula` - 计算公式 (已注释)

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

## 代码示例和最佳实践

### 游戏配置表管理示例

#### 1. 技能配置表批量操作
```python
# 步骤1: 搜索特定技能类型
search_result = excel_search(
    file_path="skills.xlsx",
    pattern=r"火系|冰系|雷系",
    sheet_name="技能配置表",
    use_regex=True
)

# 步骤2: 获取技能伤害数据
damage_data = excel_get_range(
    file_path="skills.xlsx",
    range="技能配置表!G2:G100"  # 伤害列
)

# 步骤3: 批量提升伤害值 (示例: 提升20%)
if damage_data['success']:
    updated_damage = []
    for row in damage_data['data']:
        if row and row[0] and isinstance(row[0], (int, float)):
            updated_damage.append([row[0] * 1.2])  # 提升20%
        else:
            updated_damage.append(row)

    # 步骤4: 更新数据
    excel_update_range(
        file_path="skills.xlsx",
        range="技能配置表!G2:G100",
        data=updated_damage,
        preserve_formulas=False
    )

# 步骤5: 高亮显示修改的技能
excel_format_cells(
    file_path="skills.xlsx",
    sheet_name="技能配置表",
    range="G2:G100",
    preset="highlight"
)
```

#### 2. 装备配置表分析
```python
# 获取装备表结构
headers = excel_get_headers("items.xlsx", "装备配置表")
print(f"装备表字段: {headers['field_names']}")

# 查找传奇装备
legendary_items = excel_search(
    file_path="items.xlsx",
    pattern="传奇",
    sheet_name="装备配置表",
    whole_word=True
)

# 分析装备品质分布
quality_column = "D"  # 假设品质在第4列
all_items = excel_get_range("items.xlsx", "装备配置表!D2:D200")

# 统计各品质数量
quality_count = {}
for item in all_items['data']:
    if item and item[0]:
        quality = item[0]
        quality_count[quality] = quality_count.get(quality, 0) + 1

print(f"装备品质分布: {quality_count}")
```

#### 3. 怪物配置表数值平衡
```python
# 检查怪物数值平衡
monsters = excel_get_range("monsters.xlsx", "怪物配置表!A2:F100")

unbalanced_monsters = []
for i, monster in enumerate(monsters['data'], start=2):  # 从第2行开始
    if len(monster) >= 6:  # 确保有足够的数据
        monster_id, name, level, hp, attack, defense = monster[:6]

        if isinstance(level, (int, float)) and level >= 20 and level <= 30:
            # 检查血量和攻击力的比例是否合理
            hp_attack_ratio = hp / attack if attack > 0 else 0

            # 如果比例超过某个阈值，标记为不平衡
            if hp_attack_ratio > 50 or hp_attack_ratio < 10:
                unbalanced_monsters.append({
                    'row': i,
                    'id': monster_id,
                    'name': name,
                    'level': level,
                    'hp_attack_ratio': hp_attack_ratio
                })

# 输出不平衡的怪物
if unbalanced_monsters:
    print(f"发现 {len(unbalanced_monsters)} 个数值不平衡的怪物:")
    for monster in unbalanced_monsters:
        print(f"  行{monster['row']}: {monster['name']} (等级{monster['level']}, 比例{monster['hp_attack_ratio']:.2f})")
```

### 错误处理最佳实践

#### 1. 安全的文件操作
```python
def safe_update_config(file_path, sheet_name, range_expr, new_data):
    """安全的配置表更新函数"""
    try:
        # 步骤1: 检查文件是否存在
        if not os.path.exists(file_path):
            return {'success': False, 'error': f'文件不存在: {file_path}'}

        # 步骤2: 检查工作表是否存在
        sheets = excel_list_sheets(file_path)
        if not sheets['success'] or sheet_name not in sheets['sheets']:
            return {'success': False, 'error': f'工作表不存在: {sheet_name}'}

        # 步骤3: 验证数据格式
        if not isinstance(new_data, list) or not all(isinstance(row, list) for row in new_data):
            return {'success': False, 'error': '数据格式错误，需要二维数组'}

        # 步骤4: 执行更新
        result = excel_update_range(
            file_path=file_path,
            range=range_expr,
            data=new_data,
            insert_mode=True  # 使用插入模式更安全
        )

        # 步骤5: 验证更新结果
        if result['success']:
            verification = excel_get_range(file_path, range_expr)
            if verification['success']:
                print(f"✅ 成功更新 {len(new_data)} 行数据")

        return result

    except Exception as e:
        return {'success': False, 'error': f'更新失败: {str(e)}'}
```

#### 2. 批量操作模式
```python
def batch_process_game_configs(config_dir, operation_type):
    """批量处理游戏配置表"""
    import os
    import glob

    # 获取所有Excel文件
    excel_files = glob.glob(os.path.join(config_dir, "*.xlsx"))

    results = []
    for file_path in excel_files:
        try:
            # 获取文件信息
            file_info = excel_get_file_info(file_path)
            if not file_info['success']:
                continue

            # 获取所有工作表
            sheets = excel_list_sheets(file_path)
            if not sheets['success']:
                continue

            for sheet_name in sheets['sheets']:
                # 根据操作类型执行相应处理
                if operation_type == "validate_ids":
                    # 验证ID重复
                    duplicate_check = excel_check_duplicate_ids(
                        file_path, sheet_name, id_column=1
                    )
                    if duplicate_check['has_duplicates']:
                        results.append({
                            'file': os.path.basename(file_path),
                            'sheet': sheet_name,
                            'duplicates': duplicate_check['duplicate_count']
                        })

                elif operation_type == "find_last_row":
                    # 查找数据边界
                    last_row = excel_find_last_row(file_path, sheet_name)
                    if last_row['success']:
                        results.append({
                            'file': os.path.basename(file_path),
                            'sheet': sheet_name,
                            'last_row': last_row['last_row']
                        })

        except Exception as e:
            results.append({
                'file': os.path.basename(file_path),
                'error': str(e)
            })

    return results
```

### 性能优化建议

#### 1. 大文件处理
```python
# 对于大文件，使用分批处理
def process_large_excel(file_path, sheet_name, batch_size=1000):
    """分批处理大型Excel文件"""
    last_row = excel_find_last_row(file_path, sheet_name)
    if not last_row['success']:
        return

    total_rows = last_row['last_row']
    processed = 0

    while processed < total_rows:
        start_row = processed + 2  # 跳过表头
        end_row = min(processed + batch_size + 1, total_rows)

        # 读取一批数据
        range_expr = f"{sheet_name}!A{start_row}:Z{end_row}"
        batch_data = excel_get_range(file_path, range_expr)

        if batch_data['success']:
            # 处理这批数据
            process_batch(batch_data['data'])
            processed += batch_size
            print(f"已处理 {processed}/{total_rows} 行")
```

#### 2. 缓存策略
```python
# 利用工作簿缓存提高性能
def cached_read_operations(file_path, operations):
    """使用缓存执行多个读取操作"""
    results = {}

    try:
        # ExcelReader会自动缓存工作簿
        from ..core.excel_reader import ExcelReader
        reader = ExcelReader(file_path)

        for op_name, range_expr in operations:
            result = reader.get_range(range_expr)
            results[op_name] = result

        reader.close()
        return results

    except Exception as e:
        return {'error': f'缓存读取失败: {str(e)}'}
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

### 目录结构要求

项目必须严格遵循以下目录结构，确保代码组织的一致性和可维护性：

```
excel-mcp-server/
├── src/                        # 源代码目录 (必需)
│   ├── server.py               # MCP服务器入口，仅包含MCP接口定义 (必需)
│   ├── __init__.py            # 包初始化文件 (必需)
│   ├── api/                    # API业务逻辑层 (必需)
│   │   ├── __init__.py        # 包初始化文件 (必需)
│   │   └── excel_operations.py # 集中式业务逻辑处理 (必需)
│   ├── core/                   # 核心操作层 (必需)
│   │   ├── __init__.py        # 包初始化文件 (必需)
│   │   ├── excel_reader.py    # Excel读取操作 (必需)
│   │   ├── excel_writer.py    # Excel写入操作 (必需)
│   │   ├── excel_manager.py   # 文件和工作表管理 (必需)
│   │   ├── excel_search.py    # 搜索功能 (必需)
│   │   ├── excel_compare.py   # Excel对比操作 (必需)
│   │   └── excel_converter.py # 格式转换功能 (必需)
│   ├── utils/                  # 工具层 (必需)
│   │   ├── __init__.py        # 包初始化文件 (必需)
│   │   ├── formatter.py       # 结果格式化工具 (必需)
│   │   ├── validators.py      # 数据验证工具 (必需)
│   │   ├── parsers.py         # 数据解析工具 (必需)
│   │   ├── exceptions.py      # 异常定义 (必需)
│   │   └── error_handler.py   # 错误处理工具 (必需)
│   └── models/                 # 数据模型层 (必需)
│       ├── __init__.py        # 包初始化文件 (必需)
│       └── types.py           # 类型定义 (必需)
├── tests/                      # 测试目录 (必需)
│   ├── __init__.py           # 测试包初始化 (必需)
│   ├── conftest.py           # 测试配置和fixtures (必需)
│   ├── test_api_excel_operations.py  # API层测试 (必需)
│   ├── test_core.py          # 核心操作测试 (必需)
│   ├── test_server.py        # MCP接口测试 (必需)
│   ├── test_utils.py         # 工具函数测试 (必需)
│   ├── test_search.py        # 搜索功能测试 (必需)
│   ├── test_excel_converter.py # 格式转换测试 (必需)
│   ├── test_excel_compare.py  # 对比功能测试 (必需)
│   ├── test_error_handler.py # 错误处理测试 (必需)
│   ├── test_features.py      # 功能特性测试 (必需)
│   ├── test_new_features.py  # 新功能测试 (必需)
│   ├── test_new_apis.py      # 新API测试 (必需)
│   ├── test_excel_operations_extended.py # 扩展操作测试 (必需)
│   ├── test_range_search.py  # 范围搜索测试 (必需)
│   ├── test_duplicate_ids.py # ID重复检测测试 (必需)
│   └── test_data/            # 测试数据目录 (必需)
│       ├── demo_test.xlsx    # 演示测试文件 (必需)
│       └── comprehensive_test.xlsx # 综合测试文件 (必需)
├── scripts/                    # 脚本工具目录 (必需)
│   └── run_tests.py         # 测试运行脚本 (必需)
├── docs/                       # 文档目录 (必需)
│   ├── 游戏开发Excel配置表比较指南.md # 游戏开发指南 (必需)
│   └── archive/             # 归档文档目录 (必需)
├── htmlcov/                    # 覆盖率报告目录 (自动生成)
├── pyproject.toml             # 项目配置文件 (必需)
├── uv.lock                    # 依赖锁定文件 (必需)
├── README.md                  # 项目说明文档 (必需)
├── README.en.md              # 英文版说明文档 (必需)
├── DEPLOYMENT.md              # 部署指南 (必需)
├── LICENSE                    # 开源许可证 (必需)
├── CONTRIBUTING.md            # 贡献指南 (必需)
├── CONTRIBUTING.zh-CN.md      # 中文贡献指南 (必需)
├── CLAUDE.md                  # Claude代码指导 (必需)
├── deploy.bat                 # Windows部署脚本 (必需)
├── mcp-windows.json           # Windows MCP配置 (必需)
├── mcp-direct.json            # 直接运行MCP配置 (必需)
├── mcp-generated.json         # 生成的MCP配置 (必需)
└── 项目说明.md                 # 项目中文说明 (必需)
```

### 目录结构验证

#### 自动验证脚本
```bash
# 验证目录结构完整性
python -c "
import os
from pathlib import Path

required_dirs = [
    'src', 'src/api', 'src/core', 'src/utils', 'src/models',
    'tests', 'tests/test_data', 'scripts', 'docs', 'docs/archive'
]

required_files = [
    'src/server.py',
    'src/api/excel_operations.py',
    'src/core/excel_reader.py',
    'src/utils/formatter.py',
    'src/models/types.py',
    'tests/conftest.py',
    'tests/test_api_excel_operations.py',
    'tests/test_core.py',
    'tests/test_server.py',
    'scripts/run_tests.py',
    'pyproject.toml',
    'README.md',
    'CLAUDE.md'
]

# 检查目录和文件是否存在
for dir_path in required_dirs:
    if not Path(dir_path).exists():
        print(f'❌ 缺少目录: {dir_path}')
    else:
        print(f'✅ 目录存在: {dir_path}')

for file_path in required_files:
    if not Path(file_path).exists():
        print(f'❌ 缺少文件: {file_path}')
    else:
        print(f'✅ 文件存在: {file_path}')
"
```

### 目录组织原则

1. **分层结构**: 严格按照MCP接口层 → API层 → 核心层 → 工具层的分层组织
2. **职责分离**: 每个目录只负责单一职责，避免混合功能模块
3. **命名规范**:
   - 文件名使用小写字母和下划线
   - 测试文件以`test_`开头
   - 核心模块以`excel_`开头
4. **包结构**: 每个目录都必须包含`__init__.py`文件
5. **测试覆盖**: 每个核心模块都必须有对应的测试文件

### 文件命名约定

#### 核心模块文件
- `excel_reader.py` - Excel读取操作
- `excel_writer.py` - Excel写入操作
- `excel_manager.py` - 文件和工作表管理
- `excel_search.py` - 搜索功能
- `excel_compare.py` - 对比操作
- `excel_converter.py` - 格式转换

#### 工具模块文件
- `formatter.py` - 结果格式化
- `validators.py` - 数据验证
- `parsers.py` - 数据解析
- `exceptions.py` - 异常定义
- `error_handler.py` - 错误处理

#### 测试文件
- `test_api_excel_operations.py` - API层测试
- `test_core.py` - 核心功能测试
- `test_server.py` - MCP接口测试
- `test_[功能名].py` - 特定功能测试

### 自动整理功能

如果检测到目录结构不符合要求，可以运行以下自动整理脚本：

```bash
# 自动创建缺失的目录结构
python scripts/ensure_directory_structure.py

# 自动验证和报告目录结构状态
python scripts/validate_directory_structure.py

# 自动整理文件到正确的目录
python scripts/organize_files.py
```

### 持续集成检查

在CI/CD流水线中集成目录结构检查：

```yaml
# .github/workflows/validate-structure.yml
name: Validate Directory Structure
on: [push, pull_request]
jobs:
  validate-structure:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - name: Validate directory structure
        run: python scripts/validate_directory_structure.py
      - name: Check required files
        run: python scripts/check_required_files.py
```

## 开发环境配置

### Python 环境要求
- **Python 版本**: 3.10 或更高
- **操作系统**: Windows/macOS/Linux
- **内存要求**: 建议 4GB+ (处理大文件时)

### 开发工具推荐
- **IDE**: VS Code / PyCharm / Cursor
- **包管理器**: uv (推荐) / pip
- **版本控制**: Git
- **测试框架**: pytest (已配置)

### 快速开发环境设置

#### 1. 使用 uv (推荐)
```bash
# 安装 uv
pip install uv

# 克隆项目
git clone https://github.com/tangjian/excel-mcp-server.git
cd excel-mcp-server

# 同步依赖
uv sync

# 激活虚拟环境
source .venv/bin/activate  # Linux/macOS
# 或者
.venv\Scripts\activate     # Windows

# 运行测试验证环境
python scripts/run_tests.py
```

#### 2. 使用 pip
```bash
# 创建虚拟环境
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
.venv\Scripts\activate     # Windows

# 安装依赖
pip install -e .

# 运行测试
python scripts/run_tests.py
```

### IDE 配置

#### VS Code 配置
```json
// .vscode/settings.json
{
    "python.defaultInterpreterPath": "./.venv/bin/python",
    "python.linting.enabled": true,
    "python.linting.pylintEnabled": true,
    "python.formatting.provider": "black",
    "python.testing.pytestEnabled": true,
    "python.testing.pytestArgs": ["-m", "pytest", "tests/", "-v"],
    "files.exclude": {
        "**/__pycache__": true,
        "**/*.pyc": true,
        ".pytest_cache": true,
        "htmlcov": true
    }
}
```

#### PyCharm 配置
1. 打开项目后设置 Python 解释器指向 `.venv`
2. 将 `src` 目录标记为源代码根目录
3. 配置 pytest 为默认测试运行器

### 环境变量配置
```bash
# Windows
set PYTHONPATH=%cd%\src;%PYTHONPATH%
set EXCELMCP_DEBUG=1

# Linux/macOS
export PYTHONPATH="${PWD}/src:${PYTHONPATH}"
export EXCELMCP_DEBUG=1
```

### 调试配置

#### 启用调试模式
```python
# 在 src/server.py 中修改日志级别
logging.basicConfig(
    level=logging.DEBUG,  # 改为 DEBUG 级别
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
```

#### 测试单个工具
```bash
# 测试特定功能
python -m pytest tests/test_api_excel_operations.py::TestExcelOperations::test_get_range_success_flow -v -s

# 显示详细输出
python -m pytest tests/test_search.py -v -s --tb=long

# 运行覆盖率测试
python -m pytest tests/ --cov=src --cov-report=html --cov-report=term

# 传统方式（备用）
python scripts/run_tests.py
```

## 部署指南

### 生产环境部署

#### 1. Docker 部署 (推荐)
```dockerfile
# Dockerfile
FROM python:3.11-slim

WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖文件
COPY pyproject.toml uv.lock ./
COPY src/ ./src/

# 安装 uv 并同步依赖
RUN pip install uv && \
    uv sync --frozen

# 暴露端口
EXPOSE 8000

# 启动命令
CMD ["uv", "run", "python", "-m", "src.server"]
```

```yaml
# docker-compose.yml
version: '3.8'
services:
  excelmcp:
    build: .
    ports:
      - "8000:8000"
    volumes:
      - ./data:/app/data
      - ./logs:/app/logs
    environment:
      - PYTHONPATH=/app/src
      - EXCELMCP_LOG_LEVEL=INFO
    restart: unless-stopped
```

#### 2. 系统服务部署
```bash
# 创建系统服务 (Linux)
sudo nano /etc/systemd/system/excelmcp.service
```

```ini
[Unit]
Description=Excel MCP Server
After=network.target

[Service]
Type=simple
User=excelmcp
WorkingDirectory=/opt/excelmcp
Environment=PYTHONPATH=/opt/excelmcp/src
ExecStart=/opt/excelmcp/.venv/bin/python -m src.server
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

```bash
# 启用服务
sudo systemctl enable excelmcp
sudo systemctl start excelmcp
sudo systemctl status excelmcp
```

### 监控和日志

#### 日志配置
```python
# 在 src/server.py 中配置日志
import logging
from logging.handlers import RotatingFileHandler

# 配置文件日志
file_handler = RotatingFileHandler(
    'logs/excelmcp.log',
    maxBytes=10*1024*1024,  # 10MB
    backupCount=5
)
file_handler.setFormatter(
    logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
)

logger.addHandler(file_handler)
```

#### 健康检查
```python
# 添加健康检查端点
def health_check():
    """简单的健康检查"""
    try:
        # 测试基本功能
        test_file = "health_check_test.xlsx"
        result = excel_create_file(test_file, ["test"])

        if result['success']:
            os.remove(test_file)  # 清理测试文件
            return {"status": "healthy", "timestamp": time.time()}
        else:
            return {"status": "unhealthy", "error": result.get('error')}
    except Exception as e:
        return {"status": "error", "message": str(e)}
```

### 性能调优

#### 内存优化
```python
# 在 ExcelReader 中配置缓存大小
class ExcelReader:
    def __init__(self, file_path: str, cache_size: int = 10):
        self.file_path = file_path
        self._workbook_cache = {}
        self._cache_size = cache_size
```

#### 并发处理
```python
# 使用线程池处理多个文件
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor

def process_multiple_files(file_paths, operation):
    """并发处理多个Excel文件"""
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {
            executor.submit(operation, file_path): file_path
            for file_path in file_paths
        }

        results = {}
        for future in concurrent.futures.as_completed(futures):
            file_path = futures[future]
            try:
                results[file_path] = future.result()
            except Exception as e:
                results[file_path] = {"error": str(e)}

        return results
```

## 中文/Unicode 支持
- 完整支持中文字符的工作表名
- 双行表头系统（描述 + 字段名）
- Unicode 文本处理和标准化
- 缺失表头数据的回退机制
- 本地化 Excel 功能处理

## 故障排除

### 常见部署问题

#### 1. Windows 文件锁定
```bash
# 检查是否有Excel进程
tasklist | findstr excel

# 强制结束Excel进程
taskkill /f /im excel.exe
```

#### 2. Linux 权限问题
```bash
# 检查文件权限
ls -la /path/to/excel/files

# 修改权限
chmod 755 /path/to/excel/files
chown -R user:group /path/to/excel/files
```

#### 3. 内存不足
```python
# 监控内存使用
import psutil
import logging

def monitor_memory():
    """监控内存使用情况"""
    memory = psutil.virtual_memory()
    if memory.percent > 80:
        logging.warning(f"内存使用率过高: {memory.percent}%")
        # 触发清理操作
        cleanup_cache()
```

### 性能监控

#### 系统指标
```python
# 添加性能监控装饰器
import time
import functools
from typing import Callable

def performance_monitor(func: Callable) -> Callable:
    """性能监控装饰器"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()

        execution_time = end_time - start_time
        if execution_time > 5.0:  # 超过5秒记录警告
            logging.warning(f"{func.__name__} 执行时间过长: {execution_time:.2f}s")

        return result
    return wrapper
```
- 记得动态的更新CLAUDE.md文件