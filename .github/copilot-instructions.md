# Excel MCP Server - AI编码助手指南

基于FastMCP和openpyxl的Excel文件MCP服务器，支持AI通过自然语言操作Excel文件。

**项目规模**: 28个工具，281个测试（100%通过率），85%+场景覆盖，企业级可靠性

## 核心架构

### 模块组织
- **src/server.py**: 纯MCP工具定义，委托给ExcelOperations
- **src/api/excel_operations.py**: 集中化业务逻辑处理中心
- **src/core/**: Excel操作模块(reader, writer, manager, search, compare)
- **src/utils/**: 格式化器和工具函数
- **src/models/**: 类型定义和数据模型

### 设计模式

#### 纯委托架构(重构后)
所有MCP工具使用简单委托：
```python
@mcp.tool()
def excel_list_sheets(file_path: str) -> Dict[str, Any]:
    return ExcelOperations.list_sheets(file_path)
```

**ExcelOperations作为中心化业务逻辑处理器**:
- 集中处理所有Excel业务逻辑
- 统一参数验证和错误处理
- 标准化结果格式化
- 与核心模块的解耦通信

#### 统一结果格式
```python
{
    'success': bool,
    'data': Any,        # 核心数据
    'message': str,
    'metadata': dict    # 附加上下文
}
```

#### 范围表达式支持
- 带工作表: `"Sheet1!A1:C10"` 或 `"TrSkill!A1:Z100"`
- 行范围: `"Sheet1!1:5"` 或 `"3:8"`
- 列范围: `"Sheet1!A:C"` 或 `"B:E"`
- 单行/列: `"Sheet1!5"` 或 `"C"`

## 开发工作流

### 运行和测试
```bash
# 开发运行
python -m src.server

# 完整测试(221个)
pytest tests/ -v

# 模块测试
pytest tests/test_api_excel_operations.py -v
pytest tests/test_core.py -v
```

### MCP客户端配置
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

### 测试策略
- **281个测试，100%通过率**
- **分层测试**: API/核心/MCP接口测试
- **Mock隔离**: API层使用Mock确保独立性
- 关键文件: `test_api_excel_operations.py`(API), `test_core.py`(核心), `test_server.py`(MCP)

#### 测试层级架构
```
MCP层测试 (test_server.py)
    ↓ 委托调用测试
API层测试 (test_api_excel_operations.py)
    ↓ Mock隔离测试
核心层测试 (test_core.py)
    ↓ 集成测试
工具层测试 (test_utils.py)
```

## 开发实践规范

### 委托模式实现规范
- **server.py**: 仅包含MCP工具定义，零业务逻辑
- **ExcelOperations**: 集中化业务逻辑，统一错误处理
- **Core模块**: 专注单一职责的Excel操作
- **严格分离**: MCP接口 → API层 → 核心层 → 工具层

### 代码质量标准
- **方法复杂度**: 主干方法≤20行，分支方法≤50行
- **Excel约定**: 使用1-based索引匹配Excel惯例
- **默认行为**: 保留公式(`preserve_formulas=True`)
- **命名规范**: 明确的动词+名词组合，避免handle/process

### 性能优化模式
- **工作簿缓存**: 避免重复加载大型Excel文件
- **分层优化**: 主干优先可读性，分支专注性能
- **内存管理**: 及时释放资源，防止内存泄漏

### 错误处理策略
- **防洪水机制**: 错误日志聚合去重，避免日志泛滥
- **统一格式**: 所有操作返回标准结构`{success, data, message, metadata}`
- **graceful degradation**: 处理Excel文件损坏、权限等边界情况

## 工具分类和使用约定

### 28工具分类体系
1. **文件和工作表管理** (8个): create_file, list_sheets, create_sheet等
2. **数据操作** (7个): get_range, update_range, insert_rows等
3. **搜索和发现** (4个): search, search_directory, get_headers等
4. **格式化和样式** (6个): format_cells, merge_cells, set_borders等
5. **导入导出转换** (3个): export_to_csv, import_from_csv, compare_sheets

### 范围表达式约定
- **完整范围**: `"Sheet1!A1:C10"` - 必须包含工作表名
- **行范围**: `"Sheet1!1:5"` - 指定行范围
- **列范围**: `"Sheet1!A:C"` - 指定列范围
- **单元素**: `"Sheet1!5"` 或 `"Sheet1!C"` - 单行或单列

## 项目特色

### 游戏开发特化
- **Excel配置表比较**: 专为游戏配置设计
- **ID对象跟踪**: 检测新增/修改/删除对象
- **紧凑数组格式**: 优化游戏数据传输
- **TrSkill表分析**: 技能配置专项比较

### 中文/Unicode全支持
- 中文工作表名处理和编码
- Unicode标准化文本处理
- 本地化Excel特性回退机制

### Excel操作约定
- **1-based索引**匹配Excel惯例
- **默认保留公式**(`preserve_formulas=True`)
- 支持`.xlsx`和`.xlsm`格式
- 游戏表结构感知能力

## 关键依赖
- **FastMCP**: MCP服务器框架
- **openpyxl**: 核心Excel文件操作
- **xlcalculator/formulas**: 公式评估引擎
- **xlwings**: 可选Excel应用集成
- **pytest/pytest-asyncio**: 测试框架

## 常用操作

### 文件和工作表管理
文件创建、工作表CRUD、中文名支持、活动表自动管理

### 数据操作
基于范围的读写、格式保留、行列插入删除、单元格格式化预设

### 搜索分析
正则搜索、目录批量搜索、游戏配置表比较、公式评估

开发时始终使用统一的委托模式，将实现委托给核心模块，保持一致的结果格式。
