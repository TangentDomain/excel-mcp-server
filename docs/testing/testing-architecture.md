# 🔬 Excel MCP Server 测试架构指南

> **测试架构和分类专题文档**

## 测试分层架构

### 测试金字塔模型

```
    E2E Tests (端到端测试) - 少量，关键用户流程
        ↓  
    Integration Tests (集成测试) - 中等数量，模块间协作
        ↓
    Unit Tests (单元测试) - 大量，快速反馈
```

### 各层测试特点

#### 1. 单元测试 (Unit Tests) 📦
**文件位置**: `tests/test_*.py`

**目标**:
- 测试单个函数、类的独立功能
- 验证业务逻辑的正确性
- 快速反馈问题定位

**覆盖范围**:
- API层测试: `test_api_excel_operations.py`
- 核心模块测试: `test_core.py`
- 工具函数测试: `test_utils.py`
- 特定功能测试: `test_search.py`, `test_excel_compare.py` 等

**示例**:
```python
class TestExcelReader:
    def test_load_workbook_success(self):
        # 测试正常加载工作簿
        workbook = ExcelReader.load_workbook("test.xlsx")
        assert workbook is not None
        
    def test_load_workbook_file_not_found(self):
        # 测试文件不存在的情况
        with pytest.raises(FileNotFoundError):
            ExcelReader.load_workbook("nonexistent.xlsx")
```

#### 2. 集成测试 (Integration Tests) 🔗
**文件位置**: `tests/integration/`

**目标**:
- 测试模块间的协作
- 验证数据流转的正确性
- 测试外部依赖的集成

**覆盖范围**:
- MCP工具集成测试: `test_mcp_tools.py`
- 数据库连接测试: `test_database_integration.py`
- 文件系统操作测试: `test_file_operations.py`

**示例**:
```python
class TestMCPIntegration:
    def test_query_join_operation(self):
        # 测试跨表JOIN操作
        result = mcp_client.query("SELECT * FROM skills JOIN classes ON skills.class_id = classes.id")
        assert len(result) > 0
        assert 'class_name' in result[0]
```

#### 3. 端到端测试 (E2E Tests) 🚀
**文件位置**: `tests/e2e/`

**目标**:
- 测试完整的用户流程
- 验证系统的端到端功能
- 模拟真实使用场景

**覆盖范围**:
- 完整配置管理流程: `test_full_workflow.py`
- 性能压力测试: `test_performance.py`
- 用户场景测试: `test_user_scenarios.py`

**示例**:
```python
class TestEndToEnd:
    def test_skill_configuration_workflow(self):
        # 测试完整的技能配置流程
        # 1. 创建技能配置
        skill_data = create_test_skill()
        # 2. 验证配置正确性
        validate_skill_configuration(skill_data)
        # 3. 应用到游戏系统
        apply_to_game_system(skill_data)
        # 4. 验证游戏效果
        verify_game_effects()
```

## 测试分类详解

### 按功能分类

#### 核心功能测试
- **Excel操作测试**: 读写、格式处理、数据转换
- **SQL查询测试**: WHERE、JOIN、GROUP BY、子查询
- **MCP工具测试**: 53个游戏专用工具的验证

#### 性能测试
- **大文件处理**: 10万+行Excel文件读写性能
- **并发处理**: 多用户同时访问的稳定性
- **内存使用**: 内存占用和垃圾回收优化

#### 安全测试
- **输入验证**: 防止SQL注入和恶意数据
- **权限控制**: 文件访问权限验证
- **异常处理**: 错误信息泄露防护

### 按数据类型分类

#### 测试数据策略
```python
# 测试数据分类
TEST_DATA_CATEGORIES = {
    'normal': 正常数据，测试标准功能
    'edge': 边界数据，测试极限情况
    'invalid': 无效数据，测试错误处理
    'large': 大量数据，测试性能表现
}
```

### 按测试阶段分类

#### 开发阶段测试
- **单元测试**: 开发者编写，快速反馈
- **功能测试**: 功能完成后验证

#### 发布阶段测试  
- **集成测试**: 模块组合测试
- **回归测试**: 确保新功能不破坏现有功能
- **性能测试**: 系统性能基准测试

## 测试执行策略

### 测试优先级
1. **P0 - 核心功能**: 关键业务逻辑，必须100%通过
2. **P1 - 重要功能**: 主要用户流程，必须通过  
3. **P2 - 一般功能**: 次要功能，建议通过
4. **P3 - 边缘功能**: 可选功能，按需执行

### 测试执行顺序
```bash
# 开发阶段
pytest tests/unit/ -v                    # 单元测试（3-5秒）
pytest tests/integration/test_core.py -v  # 核心集成测试（10-15秒）

# 发布阶段  
pytest tests/ -q --tb=short -n auto      # 全量测试并行（2-3分钟）
pytest tests/e2e/ -v                     # 端到端测试（5-10分钟）
```

## 测试环境配置

### 测试数据管理
```python
# 测试数据目录结构
tests/
├── data/                    # 测试数据文件
│   ├── normal/             # 正常测试数据
│   ├── edge/               # 边界测试数据
│   └── invalid/            # 无效测试数据
├── fixtures/               # 测试固件
├── mocks/                 # Mock对象
└── reports/               # 测试报告
```

### 测试工具链
- **pytest**: 测试框架
- **pytest-xdist**: 并行执行
- **pytest-mock**: Mock对象
- **coverage.py**: 代码覆盖率
- **pytest-benchmark**: 性能测试

---

*本文档是测试指南系列专题之一，更多内容请查看相关专题文档*