# Excel MCP Server 测试指南

本文档提供 Excel MCP Server 项目的全面测试指南，帮助开发者编写高质量、可维护的测试代码。

## 目录

1. [测试架构和分类](#测试架构和分类)
2. [测试命名规范](#测试命名规范)
3. [测试数据管理](#测试数据管理)
4. [Mock使用指南](#mock使用指南)
5. [覆盖率要求](#覆盖率要求)
6. [性能测试指南](#性能测试指南)
7. [集成测试策略](#集成测试策略)
8. [故障排除指南](#故障排除指南)

---

## 测试架构和分类

### 测试分层架构

```
测试金字塔:
    E2E Tests (端到端测试) - 少量，关键用户流程
        ↓
    Integration Tests (集成测试) - 中等数量，模块间协作
        ↓
    Unit Tests (单元测试) - 大量，快速反馈
```

### 测试分类

#### 1. 单元测试 (Unit Tests)
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
        """测试成功加载工作簿"""
        reader = ExcelReader(self.temp_file)
        workbook = reader._workbook
        assert workbook is not None
        reader.close()

    def test_load_workbook_file_not_found(self):
        """测试文件不存在的情况"""
        with pytest.raises(FileNotFoundError):
            ExcelReader("nonexistent.xlsx")
```

#### 2. 集成测试 (Integration Tests)
**目标**:
- 测试模块间的协作
- 验证数据流转的正确性
- 确保接口契约的遵守

**示例**:
```python
class TestExcelOperationsIntegration:
    def test_create_and_populate_workflow(self):
        """测试创建和填充Excel的完整工作流"""
        # 1. 创建文件
        create_result = excel_create_file(self.temp_file, ["TestSheet"])
        assert create_result['success']

        # 2. 写入数据
        data = [["ID", "Name"], [1, "Test"]]
        update_result = excel_update_range(
            self.temp_file, "TestSheet!A1:B2", data
        )
        assert update_result['success']

        # 3. 读取验证
        read_result = excel_get_range(self.temp_file, "TestSheet!A1:B2")
        assert read_result['success']
        assert read_result['data'] == data
```

#### 3. 端到端测试 (E2E Tests)
**目标**:
- 测试完整的用户场景
- 验证系统的端到端功能
- 确保用户体验的一致性

**示例**:
```python
class TestMCPWorkflow:
    def test_complete_game_config_management(self):
        """测试完整的游戏配置管理工作流"""
        # 1. 创建技能配置表
        file_result = excel_create_file("skills_config.xlsx", ["TrSkill"])

        # 2. 设置双行表头
        headers = [
            ["技能ID描述", "技能名称描述", "技能类型描述"],
            ["skill_id", "skill_name", "skill_type"]
        ]
        header_result = excel_update_range(
            "skills_config.xlsx", "TrSkill!A1:C2", headers
        )

        # 3. 添加技能数据
        skills = [
            [1001, "火球术", "active"],
            [1002, "冰盾", "passive"],
            [1003, "闪电链", "active"]
        ]
        data_result = excel_update_range(
            "skills_config.xlsx", "TrSkill!A3:C5", skills
        )

        # 4. 验证数据完整性
        validation_result = excel_check_duplicate_ids(
            "skills_config.xlsx", "TrSkill", id_column=1
        )

        assert not validation_result['has_duplicates']
```

### 测试文件组织

```
tests/
├── conftest.py              # 测试配置和fixtures
├── unit/                    # 单元测试
│   ├── test_core.py        # 核心模块测试
│   ├── test_utils.py       # 工具函数测试
│   └── test_models.py      # 数据模型测试
├── integration/             # 集成测试
│   ├── test_workflows.py   # 工作流测试
│   └── test_apis.py        # API集成测试
├── features/               # 功能测试
│   ├── test_search.py      # 搜索功能
│   ├── test_compare.py     # 对比功能
│   └── test_converter.py   # 转换功能
├── performance/            # 性能测试
│   ├── test_memory.py      # 内存使用测试
│   └── test_speed.py       # 执行速度测试
├── e2e/                    # 端到端测试
│   └── test_scenarios.py   # 用户场景测试
└── test_data/              # 测试数据
    ├── demo_test.xlsx
    └── comprehensive_test.xlsx
```

---

## 测试命名规范

### 文件命名规范

#### 测试文件
- **格式**: `test_[module_name].py`
- **示例**:
  - `test_excel_reader.py` - 测试Excel读取模块
  - `test_api_excel_operations.py` - 测试API操作
  - `test_search_functionality.py` - 测试搜索功能

#### 测试类
- **格式**: `Test[ClassName]` 或 `Test[FeatureName]`
- **示例**:
  - `TestExcelReader` - 测试ExcelReader类
  - `TestSearchFunctionality` - 测试搜索功能
  - `TestDuplicateIdDetection` - 测试ID重复检测

#### 测试方法
- **格式**: `test_[scenario]_[expected_result]`
- **示例**:
  - `test_get_range_success_flow` - 测试成功获取范围的流程
  - `test_update_range_with_invalid_data_should_fail` - 测试无效数据更新应失败
  - `test_search_with_regex_pattern_should_find_matches` - 测试正则模式搜索

### 命名约定详解

#### 1. 成功场景命名
```python
def test_create_file_successfully(self):
    """测试成功创建文件"""
    pass

def test_get_range_returns_valid_data(self):
    """测试获取范围返回有效数据"""
    pass

def test_search_finds_expected_matches(self):
    """测试搜索找到预期匹配"""
    pass
```

#### 2. 失败场景命名
```python
def test_update_file_with_invalid_path_should_raise_error(self):
    """测试无效路径更新文件应抛出错误"""
    pass

def test_get_range_with_out_of_bounds_should_return_empty(self):
    """测试越界范围获取应返回空结果"""
    pass

def test_search_with_empty_pattern_should_return_all_results(self):
    """测试空模式搜索应返回所有结果"""
    pass

def test_operation_with_corrupted_file_should_fail_gracefully(self):
    """测试损坏文件操作应优雅失败"""
    pass
```

#### 3. 边界条件命名
```python
def test_operation_with_empty_data_should_handle_correctly(self):
    """测试空数据操作应正确处理"""
    pass

def test_operation_with_maximum_data_should_not_crash(self):
    """测试最大数据量操作不应崩溃"""
    pass

def test_concurrent_operations_should_be_thread_safe(self):
    """测试并发操作应是线程安全的"""
    pass
```

#### 4. 参数化测试命名
```python
@pytest.mark.parametrize("file_format,expected_extension", [
    ("xlsx", ".xlsx"),
    ("xlsm", ".xlsm"),
    ("csv", ".csv")
])
def test_convert_format_supports_multiple_formats(self, file_format, expected_extension):
    """测试格式转换支持多种格式"""
    pass

@pytest.mark.parametrize("invalid_range", [
    "Invalid!Range",
    "Sheet1!A0:A1",  # Excel中行从1开始
    "Sheet1!A1:ZZ999",  # 超出Excel列范围
    ""
])
def test_get_range_with_invalid_ranges_should_fail(self, invalid_range):
    """测试无效范围获取应失败"""
    pass
```

### 测试描述规范

#### Docstring格式
```python
def test_excel_operations_with_chinese_worksheet_names(self):
    """
    测试中文工作表名的Excel操作

    测试场景:
    1. 创建包含中文工作表的Excel文件
    2. 在中文工作表中执行读写操作
    3. 验证中文编码的正确性

    预期结果:
    - 工作表名应正确保存和读取
    - 数据应完整无乱码
    - 操作应返回成功状态
    """
    pass
```

#### 注释规范
```python
def test_search_with_complex_regex_pattern(self):
    # Arrange: 准备测试数据和搜索模式
    test_data = [["email@example.com", "123-456-7890"], ["test@test.org"]]
    pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

    # Act: 执行搜索操作
    result = excel_search(self.test_file, pattern, use_regex=True)

    # Assert: 验证搜索结果
    assert result['success']
    assert result['match_count'] > 0
    assert any('email@example.com' in match['cell_value']
               for match in result['matches'])
```

---

## 测试数据管理

### 测试数据策略

#### 1. 临时文件管理
```python
# conftest.py
import pytest
import tempfile
import os
from pathlib import Path

@pytest.fixture
def temp_excel_file():
    """创建临时Excel文件"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        temp_path = f.name

    yield temp_path

    # 清理临时文件
    try:
        os.unlink(temp_path)
    except FileNotFoundError:
        pass

@pytest.fixture
def temp_excel_directory():
    """创建临时Excel目录"""
    with tempfile.TemporaryDirectory() as temp_dir:
        yield temp_dir
```

#### 2. 预定义测试数据
```python
@pytest.fixture
def sample_game_skills_data():
    """提供示例游戏技能数据"""
    return {
        'headers': [
            ["技能ID描述", "技能名称描述", "技能类型描述", "技能等级描述"],
            ["skill_id", "skill_name", "skill_type", "skill_level"]
        ],
        'data': [
            [1001, "火球术", "active", 1],
            [1002, "火球术", "active", 2],
            [1003, "火球术", "active", 3],
            [2001, "冰盾", "passive", 1],
            [2002, "冰盾", "passive", 2],
            [3001, "闪电链", "active", 1],
        ]
    }

@pytest.fixture
def sample_equipment_data():
    """提供示例装备数据"""
    return {
        'headers': [
            ["装备ID", "装备名称", "装备类型", "装备品质"],
            ["item_id", "item_name", "item_type", "item_quality"]
        ],
        'data': [
            [10001, "铁剑", "weapon", "common"],
            [10002, "钢剑", "weapon", "rare"],
            [10003, "魔法剑", "weapon", "epic"],
            [20001, "皮甲", "armor", "common"],
            [20002, "锁子甲", "armor", "rare"],
        ]
    }
```

#### 3. 动态测试数据生成
```python
@pytest.fixture
def large_excel_file():
    """生成大型Excel文件用于性能测试"""
    import random
    from openpyxl import Workbook

    file_path = tempfile.mktemp(suffix='.xlsx')
    wb = Workbook()
    ws = wb.active
    ws.title = "LargeData"

    # 写入表头
    ws.append(['ID', 'Name', 'Value', 'Category'])

    # 生成大量数据
    categories = ['A', 'B', 'C', 'D', 'E']
    for i in range(1, 10001):  # 10,000行数据
        ws.append([
            i,
            f"Item_{i}",
            random.randint(1, 1000),
            random.choice(categories)
        ])

    wb.save(file_path)
    wb.close()

    yield file_path

    os.unlink(file_path)

@pytest.fixture
def excel_with_formulas():
    """创建包含公式的Excel文件"""
    from openpyxl import Workbook

    file_path = tempfile.mktemp(suffix='.xlsx')
    wb = Workbook()
    ws = wb.active
    ws.title = "Formulas"

    # 写入数据和公式
    ws.append(['A', 'B', 'Formula'])
    ws.append([10, 20, '=A2+B2'])
    ws.append([30, 40, '=A3*B3'])
    ws.append(['=SUM(A2:A3)', '=AVERAGE(B2:B3)', '=A4/B4'])

    wb.save(file_path)
    wb.close()

    yield file_path

    os.unlink(file_path)
```

### 测试数据版本管理

#### 1. 测试数据文件
```
tests/test_data/
├── fixtures/
│   ├── skills_config_v1.xlsx      # 技能配置模板v1
│   ├── skills_config_v2.xlsx      # 技能配置模板v2
│   ├── equipment_data.xlsx        # 装备数据模板
│   └── monster_data.xlsx          # 怪物数据模板
├── corrupted/
│   ├── invalid_format.xlsx        # 格式错误的文件
│   ├── password_protected.xlsx    # 密码保护的文件
│   └── corrupted_file.xlsx        # 损坏的文件
└── large/
    ├── 100k_rows.xlsx            # 10万行数据
    └── complex_formulas.xlsx      # 复杂公式文件
```

#### 2. 数据版本控制
```python
@pytest.fixture(params=[
    "skills_config_v1.xlsx",
    "skills_config_v2.xlsx"
])
def skills_config_file(request):
    """支持多版本的技能配置文件"""
    file_path = Path(__file__).parent / "test_data" / "fixtures" / request.param
    return str(file_path)

@pytest.fixture
def legacy_format_file():
    """提供旧格式文件用于兼容性测试"""
    file_path = Path(__file__).parent / "test_data" / "legacy" / "old_format.xls"
    return str(file_path)
```

### 测试数据隔离

#### 1. 数据库/文件隔离
```python
@pytest.fixture(autouse=True)
def isolate_test_data():
    """自动隔离测试数据"""
    # 测试前设置
    original_cwd = os.getcwd()
    temp_dir = tempfile.mkdtemp()
    os.chdir(temp_dir)

    yield

    # 测试后清理
    os.chdir(original_cwd)
    shutil.rmtree(temp_dir, ignore_errors=True)
```

#### 2. 环境变量隔离
```python
@pytest.fixture
def clean_environment():
    """提供清洁的环境变量"""
    original_env = os.environ.copy()

    # 清理相关环境变量
    env_keys_to_remove = ['EXCELMCP_CONFIG', 'PYTHONPATH', 'DEBUG']
    for key in env_keys_to_remove:
        os.environ.pop(key, None)

    yield

    # 恢复原始环境
    os.environ.clear()
    os.environ.update(original_env)
```

---

## Mock使用指南

### Mock基础概念

#### 1.何时使用Mock
- **外部依赖**: 文件系统、网络请求、数据库
- **时间依赖**: 时间戳、定时任务
- **复杂对象**: Excel应用实例、大型数据结构
- **错误场景**: 模拟异常、网络故障、权限错误

#### 2.何时不使用Mock
- **简单逻辑**: 纯函数计算、数据转换
- **核心业务逻辑**: 应测试实际实现
- **集成测试**: 需要真实组件协作

### Mock最佳实践

#### 1. 使用unittest.mock
```python
from unittest.mock import Mock, patch, MagicMock

class TestExcelOperations:
    @patch('src.api.excel_operations.ExcelReader')
    def test_get_range_with_mocked_reader(self, mock_reader_class):
        """测试使用Mock的Excel读取操作"""
        # Arrange: 设置Mock行为
        mock_reader = Mock()
        mock_reader_class.return_value = mock_reader
        mock_reader.get_range.return_value = {
            'success': True,
            'data': [['A1', 'B1'], ['A2', 'B2']],
            'message': 'Success'
        }

        # Act: 执行测试
        from src.api.excel_operations import excel_operations
        result = excel_operations.get_range("test.xlsx", "Sheet1!A1:B2")

        # Assert: 验证结果和Mock调用
        assert result['success']
        mock_reader_class.assert_called_once_with("test.xlsx")
        mock_reader.get_range.assert_called_once_with("Sheet1!A1:B2")
```

#### 2. Mock文件系统操作
```python
@patch('builtins.open')
@patch('os.path.exists')
def test_file_operations_with_mocked_filesystem(self, mock_exists, mock_open):
    """测试Mock文件系统操作"""
    # Arrange: 设置Mock
    mock_exists.return_value = True
    mock_file = Mock()
    mock_file.read.return_value = b"fake excel content"
    mock_open.return_value.__enter__.return_value = mock_file

    # Act: 执行需要文件系统的操作
    result = some_file_operation("test.xlsx")

    # Assert: 验证文件系统调用
    mock_exists.assert_called_once_with("test.xlsx")
    mock_open.assert_called_once_with("test.xlsx", "rb")
```

#### 3. Mock异常场景
```python
@patch('src.core.excel_reader.openpyxl.load_workbook')
def test_handling_corrupted_file_error(self, mock_load):
    """测试处理损坏文件错误"""
    # Arrange: 模拟文件损坏异常
    mock_load.side_effect = Exception("File is corrupted")

    # Act & Assert: 验证错误处理
    with pytest.raises(Exception) as exc_info:
        ExcelReader("corrupted.xlsx")

    assert "File is corrupted" in str(exc_info.value)
```

### 高级Mock技巧

#### 1. 使用side_effect
```python
def test_dynamic_responses_with_side_effect(self):
    """测试使用side_effect的动态响应"""
    from unittest.mock import Mock

    mock_function = Mock()

    # 使用列表提供连续的返回值
    mock_function.side_effect = [
        {'success': True, 'data': [[1, 2]]},
        {'success': False, 'error': 'File not found'},
        Exception("Network error")
    ]

    # 测试不同返回值
    assert mock_function().success  # 第一次调用成功
    assert not mock_function().success  # 第二次调用失败
    with pytest.raises(Exception):  # 第三次调用抛出异常
        mock_function()

def test_callable_side_effect(self):
    """测试可调用side_effect"""
    def dynamic_response(*args, **kwargs):
        """根据参数动态返回结果"""
        if args[0] == "success.xlsx":
            return {'success': True, 'data': [[1, 2]]}
        else:
            return {'success': False, 'error': 'File not found'}

    mock_function = Mock(side_effect=dynamic_response)

    # 测试动态行为
    result1 = mock_function("success.xlsx")
    assert result1['success']

    result2 = mock_function("fail.xlsx")
    assert not result2['success']
```

#### 2. Mock属性和方法
```python
def test_mocking_complex_objects(self):
    """测试Mock复杂对象"""
    # 创建Mock对象
    mock_workbook = Mock()
    mock_worksheet = Mock()

    # 设置属性
    mock_workbook.active = mock_worksheet
    mock_worksheet.title = "TestSheet"
    mock_worksheet.max_row = 100
    mock_worksheet.max_column = 10

    # 设置方法行为
    mock_worksheet.cell.return_value.value = "test_value"
    mock_worksheet.iter_rows.return_value = [
        [Mock(value="A1"), Mock(value="B1")],
        [Mock(value="A2"), Mock(value="B2")]
    ]

    # 使用Mock对象
    reader = ExcelReader("test.xlsx")
    reader._workbook = mock_workbook

    # 验证Mock调用
    cell_value = reader._workbook.active.cell(row=1, column=1).value
    assert cell_value == "test_value"
```

#### 3. 使用Patch装饰器
```python
# 类级别Patch - 应用于整个测试类
@patch('src.core.excel_reader.openpyxl')
class TestExcelReaderWithPatches:
    def test_read_with_patches(self, mock_openpyxl):
        """使用Patch的读取测试"""
        mock_workbook = Mock()
        mock_openpyxl.load_workbook.return_value = mock_workbook

        reader = ExcelReader("test.xlsx")
        mock_openpyxl.load_workbook.assert_called_once_with("test.xlsx")

    def test_close_with_patches(self, mock_openpyxl):
        """使用Patch的关闭测试"""
        mock_workbook = Mock()
        mock_openpyxl.load_workbook.return_value = mock_workbook

        reader = ExcelReader("test.xlsx")
        reader.close()
        mock_workbook.close.assert_called_once()

# 方法级别Patch - 应用于单个测试
def test_specific_method_with_patch(self):
    """特定方法的Patch测试"""
    with patch('src.utils.validators.validate_file_path') as mock_validate:
        mock_validate.return_value = True

        result = validate_excel_file("test.xlsx")

        mock_validate.assert_called_once_with("test.xlsx")
```

### Mock和实际测试的平衡

#### 1. 分层测试策略
```python
class TestExcelOperations:
    # 单元测试 - 使用Mock
    @patch('src.core.excel_reader.ExcelReader')
    def test_get_range_unit(self, mock_reader):
        """单元测试：使用Mock"""
        mock_reader.return_value.get_range.return_value = {
            'success': True, 'data': [[1, 2]]
        }

        result = excel_operations.get_range("test.xlsx", "Sheet1!A1:B2")
        assert result['success']

    # 集成测试 - 使用真实文件
    def test_get_range_integration(self, temp_excel_file):
        """集成测试：使用真实文件"""
        # 准备测试文件
        prepare_test_excel_file(temp_excel_file)

        # 执行真实操作
        result = excel_operations.get_range(temp_excel_file, "Sheet1!A1:B2")

        # 验证结果
        assert result['success']
        assert isinstance(result['data'], list)
```

#### 2. 混合使用策略
```python
def test_partial_mocking(self, temp_excel_file):
    """部分Mock：只Mock外部依赖"""
    # 准备真实文件
    prepare_test_excel_file(temp_excel_file)

    # 只Mock网络操作
    with patch('requests.get') as mock_get:
        mock_get.return_value.json.return_value = {
            'exchange_rates': {'USD': 1.0, 'EUR': 0.85}
        }

        # 测试包含网络调用的Excel操作
        result = excel_operations.convert_currency_in_file(
            temp_excel_file, "Sheet1!A1:A10", "USD", "EUR"
        )

        assert result['success']
        mock_get.assert_called_once()
```

---

## 覆盖率要求

### 覆盖率目标

#### 1. 整体覆盖率目标
- **行覆盖率**: ≥ 90%
- **分支覆盖率**: ≥ 85%
- **函数覆盖率**: ≥ 95%
- **语句覆盖率**: ≥ 90%

#### 2. 模块覆盖率分解
```
src/
├── server.py               # ≥ 95% (MCP接口关键)
├── api/
│   └── excel_operations.py # ≥ 90% (核心业务逻辑)
├── core/
│   ├── excel_reader.py     # ≥ 90% (读取功能)
│   ├── excel_writer.py     # ≥ 90% (写入功能)
│   ├── excel_manager.py    # ≥ 85% (管理功能)
│   ├── excel_search.py     # ≥ 85% (搜索功能)
│   ├── excel_compare.py    # ≥ 85% (对比功能)
│   └── excel_converter.py  # ≥ 85% (转换功能)
├── utils/
│   ├── formatter.py        # ≥ 90% (格式化工具)
│   ├── validators.py       # ≥ 95% (验证工具)
│   ├── parsers.py          # ≥ 85% (解析工具)
│   ├── exceptions.py       # ≥ 95% (异常定义)
│   └── error_handler.py    # ≥ 90% (错误处理)
└── models/
    └── types.py            # ≥ 95% (类型定义)
```

### 覆盖率测量工具

#### 1. pytest-cov配置
```bash
# 安装覆盖率工具
pip install pytest-cov coverage

# 运行完整覆盖率测试
python -m pytest tests/ --cov=src --cov-report=html --cov-report=term

# 生成HTML报告
python -m pytest tests/ --cov=src --cov-report=html

# 查看详细覆盖率
python -m pytest tests/ --cov=src --cov-report=term-missing
```

#### 2. pyproject.toml配置
```toml
[tool.coverage.run]
source = ["src"]
omit = [
    "*/tests/*",
    "*/test_*.py",
    "*/conftest.py",
    "*/__pycache__/*",
    "*/venv/*",
    "*/virtualenv/*"
]

[tool.coverage.report]
exclude_lines = [
    "pragma: no cover",
    "def __repr__",
    "if self.debug:",
    "if settings.DEBUG",
    "raise AssertionError",
    "raise NotImplementedError",
    "if 0:",
    "if __name__ == .__main__.:",
    "class .*\\bProtocol\\):",
    "@(abc\\.)?abstractmethod",
]
precision = 2
show_missing = true
fail_under = 90

[tool.coverage.html]
directory = "htmlcov"
```

#### 3. 覆盖率CI配置
```yaml
# .github/workflows/test.yml
name: Test Coverage
on: [push, pull_request]

jobs:
  test:
    runs-on: ubuntu-latest

    steps:
      - uses: actions/checkout@v3

      - name: Set up Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.10'

      - name: Install dependencies
        run: |
          pip install -e .
          pip install pytest pytest-cov

      - name: Run tests with coverage
        run: |
          python -m pytest tests/ --cov=src --cov-report=xml --cov-report=html

      - name: Check coverage threshold
        run: |
          coverage report --fail-under=90

      - name: Upload coverage to Codecov
        uses: codecov/codecov-action@v3
        with:
          file: ./coverage.xml
```

### 覆盖率分析和改进

#### 1. 覆盖率报告分析
```bash
# 生成详细报告
python -m pytest tests/ --cov=src --cov-report=term-missing

# 生成HTML报告并在浏览器中查看
python -m pytest tests/ --cov=src --cov-report=html
open htmlcov/index.html  # macOS
start htmlcov/index.html  # Windows
```

#### 2. 未覆盖代码分析
```python
# 查看特定模块的覆盖率
python -m pytest tests/test_core.py --cov=src.core.excel_reader --cov-report=term-missing

# 查看测试覆盖的代码行
coverage annotate --directory=coverage_annotate src/core/excel_reader.py
```

#### 3. 覆盖率改进策略

**识别关键路径**:
```python
# 分析关键业务路径
def get_coverage_hotspots():
    """识别需要重点关注的高价值代码路径"""
    critical_paths = [
        "src/api/excel_operations.py:ExcelOperations.get_range",
        "src/core/excel_reader.py:ExcelReader._load_workbook",
        "src/utils/validators.py:validate_range_expression",
        "src/utils/error_handler.py:handle_excel_errors"
    ]
    return critical_paths
```

**优先级排序**:
```python
# 按重要性排序测试用例
coverage_priorities = {
    "critical": [  # 核心功能，100%覆盖率
        "API层所有公共方法",
        "Excel文件读写操作",
        "错误处理和异常情况"
    ],
    "high": [  # 重要功能，90%覆盖率
        "搜索和过滤功能",
        "数据验证逻辑",
        "格式转换功能"
    ],
    "medium": [  # 辅助功能，80%覆盖率
        "工具函数和辅助方法",
        "日志记录功能",
        "配置管理"
    ],
    "low": [  # 次要功能，70%覆盖率
        "调试和开发工具",
        "实验性功能",
        "向后兼容代码"
    ]
}
```

### 覆盖率质量标准

#### 1. 测试质量 vs 覆盖率
```python
# 高质量测试的特征
class TestQualityMetrics:
    """测试质量指标"""

    def test_assertions_per_test(self):
        """每个测试应有足够的断言"""
        # 好的例子：多个相关断言
        result = excel_get_range("test.xlsx", "Sheet1!A1:C3")
        assert result['success']
        assert len(result['data']) == 3  # 3行数据
        assert len(result['data'][0]) == 3  # 3列数据
        assert isinstance(result['metadata'], dict)

    def test_descriptive_assertions(self):
        """断言应该描述性明确"""
        # 好的例子：明确的失败消息
        expected_data = [["A1", "B1"], ["A2", "B2"]]
        actual_data = result['data']
        assert actual_data == expected_data, f"Expected {expected_data}, got {actual_data}"
```

#### 2. 避免虚假覆盖率
```python
# 避免为了覆盖率而写无意义的测试
class AvoidFalseCoverage:
    """避免虚假覆盖率的例子"""

    def bad_test_trivial_code(self):
        """不好的测试：只为了覆盖率"""
        # 这种测试没有实际价值
        assert True

    def good_test_meaningful_behavior(self):
        """好的测试：测试有意义的行为"""
        # 测试实际的业务逻辑
        result = excel_operations.validate_cell_value("123", "number")
        assert result['success']
        assert result['data'] == 123  # 验证类型转换
```

---

## 性能测试指南

### 性能测试分类

#### 1. 基准测试 (Benchmark Tests)
```python
import time
import pytest
from src.core.excel_reader import ExcelReader

class TestExcelPerformance:

    def test_load_workbook_performance(self):
        """测试工作簿加载性能"""
        file_path = "large_file.xlsx"

        # 记录加载时间
        start_time = time.time()
        reader = ExcelReader(file_path)
        load_time = time.time() - start_time

        # 断言性能要求
        assert load_time < 5.0, f"加载时间过长: {load_time:.2f}秒"

        reader.close()

    @pytest.mark.parametrize("file_size", [1000, 5000, 10000, 50000])
    def test_read_performance_by_file_size(self, file_size):
        """测试不同文件大小的读取性能"""
        # 生成测试文件
        test_file = generate_excel_file(file_size)

        start_time = time.time()
        result = excel_get_range(test_file, "Sheet1!A1:Z1000")
        read_time = time.time() - start_time

        # 性能要求：每秒至少处理1000行
        rows_per_second = file_size / read_time
        assert rows_per_second >= 1000, f"读取速度过慢: {rows_per_second:.0f}行/秒"

        os.unlink(test_file)
```

#### 2. 内存使用测试
```python
import psutil
import gc
from memory_profiler import profile

class TestMemoryUsage:

    def test_memory_leak_detection(self):
        """检测内存泄漏"""
        initial_memory = psutil.Process().memory_info().rss / 1024 / 1024  # MB

        # 执行大量操作
        for _ in range(100):
            reader = ExcelReader("test.xlsx")
            data = reader.get_range("Sheet1!A1:Z1000")
            reader.close()

        gc.collect()  # 强制垃圾回收

        final_memory = psutil.Process().memory_info().rss / 1024 / 1024  # MB
        memory_increase = final_memory - initial_memory

        # 内存增长不应超过50MB
        assert memory_increase < 50, f"可能存在内存泄漏: {memory_increase:.2f}MB"

    @profile
    def test_memory_profiling(self):
        """使用memory_profiler进行详细分析"""
        # 创建大文件
        large_file = create_large_excel_file(50000)

        # 读取数据
        reader = ExcelReader(large_file)
        data = reader.get_range("Sheet1!A1:AZ50000")

        # 处理数据
        processed_data = process_large_dataset(data['data'])

        reader.close()
        os.unlink(large_file)

        return len(processed_data)
```

#### 3. 并发性能测试
```python
import threading
import concurrent.futures
from queue import Queue

class TestConcurrentPerformance:

    def test_thread_safety_performance(self):
        """测试线程安全性和并发性能"""
        file_path = "concurrent_test.xlsx"
        num_threads = 10
        operations_per_thread = 100

        results = Queue()
        errors = Queue()

        def worker():
            try:
                for i in range(operations_per_thread):
                    result = excel_get_range(file_path, f"Sheet1!A{i+1}:Z{i+1}")
                    results.put(result['success'])
            except Exception as e:
                errors.put(e)

        # 启动多个线程
        threads = []
        start_time = time.time()

        for _ in range(num_threads):
            thread = threading.Thread(target=worker)
            threads.append(thread)
            thread.start()

        # 等待所有线程完成
        for thread in threads:
            thread.join()

        execution_time = time.time() - start_time

        # 验证结果
        assert errors.empty(), f"并发操作出现错误: {list(errors.queue)}"
        successful_operations = sum(1 for success in results.queue if success)
        expected_operations = num_threads * operations_per_thread

        assert successful_operations == expected_operations
        assert execution_time < 30.0, f"并发执行时间过长: {execution_time:.2f}秒"

    def test_thread_pool_performance(self):
        """测试线程池性能"""
        file_paths = [f"test_{i}.xlsx" for i in range(20)]

        # 创建测试文件
        for file_path in file_paths:
            create_test_excel_file(file_path)

        # 使用线程池并发处理
        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            futures = [
                executor.submit(excel_get_range, file_path, "Sheet1!A1:Z100")
                for file_path in file_paths
            ]

            start_time = time.time()
            results = [future.result() for future in concurrent.futures.as_completed(futures)]
            execution_time = time.time() - start_time

        # 验证结果
        assert len(results) == len(file_paths)
        assert all(result['success'] for result in results)
        assert execution_time < 10.0, f"线程池执行时间过长: {execution_time:.2f}秒"

        # 清理测试文件
        for file_path in file_paths:
            os.unlink(file_path)
```

### 性能基准和监控

#### 1. 性能基准设定
```python
class PerformanceBenchmarks:
    """性能基准定义"""

    BENCHMARKS = {
        "file_loading": {
            "small_file": {"size_mb": 1, "max_time_seconds": 0.1},
            "medium_file": {"size_mb": 10, "max_time_seconds": 0.5},
            "large_file": {"size_mb": 100, "max_time_seconds": 3.0},
        },
        "data_reading": {
            "rows_per_second": 10000,
            "cells_per_second": 100000,
        },
        "data_writing": {
            "rows_per_second": 5000,
            "cells_per_second": 50000,
        },
        "memory_usage": {
            "base_memory_mb": 50,
            "memory_per_1000_rows_mb": 2,
            "max_memory_mb": 500,
        },
        "concurrent_operations": {
            "max_concurrent_threads": 50,
            "operations_per_second": 1000,
        }
    }

    @classmethod
    def verify_performance(cls, operation_type, measured_value, file_size=None):
        """验证性能是否达标"""
        benchmarks = cls.BENCHMARKS.get(operation_type, {})

        if operation_type == "file_loading" and file_size:
            # 根据文件大小选择合适的基准
            if file_size < 5:
                benchmark = benchmarks["small_file"]["max_time_seconds"]
            elif file_size < 50:
                benchmark = benchmarks["medium_file"]["max_time_seconds"]
            else:
                benchmark = benchmarks["large_file"]["max_time_seconds"]

            return measured_value <= benchmark

        return True
```

#### 2. 性能监控和报告
```python
import functools
import statistics
from typing import List, Dict

class PerformanceMonitor:
    """性能监控器"""

    def __init__(self):
        self.measurements: List[Dict] = []

    def measure_execution_time(self, operation_name: str):
        """装饰器：测量执行时间"""
        def decorator(func):
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                start_time = time.time()
                result = func(*args, **kwargs)
                end_time = time.time()

                execution_time = end_time - start_time
                memory_usage = psutil.Process().memory_info().rss / 1024 / 1024

                self.measurements.append({
                    'operation': operation_name,
                    'execution_time': execution_time,
                    'memory_usage': memory_usage,
                    'timestamp': time.time()
                })

                return result
            return wrapper
        return decorator

    def get_performance_report(self) -> Dict:
        """生成性能报告"""
        if not self.measurements:
            return {"error": "No measurements available"}

        # 按操作类型分组
        operations = {}
        for measurement in self.measurements:
            op_name = measurement['operation']
            if op_name not in operations:
                operations[op_name] = []
            operations[op_name].append(measurement)

        # 计算统计信息
        report = {}
        for op_name, measurements in operations.items():
            execution_times = [m['execution_time'] for m in measurements]
            memory_usages = [m['memory_usage'] for m in measurements]

            report[op_name] = {
                'count': len(measurements),
                'avg_time': statistics.mean(execution_times),
                'max_time': max(execution_times),
                'min_time': min(execution_times),
                'avg_memory': statistics.mean(memory_usages),
                'max_memory': max(memory_usages)
            }

        return report

# 使用示例
monitor = PerformanceMonitor()

@monitor.measure_execution_time("excel_read")
def test_large_file_read_performance():
    """测试大文件读取性能"""
    large_file = create_large_excel_file(100000)
    result = excel_get_range(large_file, "Sheet1!A1:AZ100000")
    os.unlink(large_file)
    return result

def test_performance_report_generation():
    """测试性能报告生成"""
    # 执行多次测试
    for _ in range(10):
        test_large_file_read_performance()

    # 生成报告
    report = monitor.get_performance_report()

    # 验证性能指标
    excel_stats = report['excel_read']
    assert excel_stats['avg_time'] < 5.0  # 平均时间应少于5秒
    assert excel_stats['max_memory'] < 200  # 最大内存应少于200MB
```

---

## 集成测试策略

### 集成测试层次

#### 1. 组件集成测试
```python
class TestComponentIntegration:
    """测试组件间的集成"""

    def test_reader_writer_integration(self):
        """测试读取器和写入器的集成"""
        # 创建测试文件
        test_file = tempfile.mktemp(suffix='.xlsx')

        # 使用写入器创建文件
        from src.core.excel_writer import ExcelWriter
        writer = ExcelWriter(test_file)
        writer.create_sheet("TestSheet")

        test_data = [
            ["ID", "Name", "Value"],
            [1, "Test1", 100],
            [2, "Test2", 200]
        ]
        writer.write_range("TestSheet!A1:C3", test_data)
        writer.close()

        # 使用读取器验证文件
        from src.core.excel_reader import ExcelReader
        reader = ExcelReader(test_file)
        result = reader.get_range("TestSheet!A1:C3")
        reader.close()

        # 验证数据一致性
        assert result['success']
        assert result['data'] == test_data

        os.unlink(test_file)

    def test_api_core_integration(self):
        """测试API层和核心层的集成"""
        from src.api.excel_operations import ExcelOperations

        operations = ExcelOperations()

        # 创建文件
        create_result = operations.create_file("integration_test.xlsx", ["TestSheet"])
        assert create_result['success']

        # 写入数据
        data = [["A", "B"], [1, 2]]
        update_result = operations.update_range(
            "integration_test.xlsx", "TestSheet!A1:B2", data
        )
        assert update_result['success']

        # 读取验证
        read_result = operations.get_range("integration_test.xlsx", "TestSheet!A1:B2")
        assert read_result['success']
        assert read_result['data'] == data

        # 清理
        os.unlink("integration_test.xlsx")
```

#### 2. 系统集成测试
```python
class TestSystemIntegration:
    """测试整个系统的集成"""

    def test_complete_excel_workflow(self):
        """测试完整的Excel工作流"""
        workflow_steps = []

        try:
            # 步骤1: 创建文件
            create_result = excel_create_file(
                "workflow_test.xlsx",
                ["Skills", "Equipment", "Monsters"]
            )
            workflow_steps.append(("create_file", create_result['success']))
            assert create_result['success']

            # 步骤2: 设置技能表结构
            headers = [
                ["技能ID", "技能名称", "技能类型", "技能等级"],
                ["skill_id", "skill_name", "skill_type", "skill_level"]
            ]
            header_result = excel_update_range(
                "workflow_test.xlsx", "Skills!A1:D2", headers
            )
            workflow_steps.append(("setup_headers", header_result['success']))
            assert header_result['success']

            # 步骤3: 添加技能数据
            skills = [
                [1001, "火球术", "active", 1],
                [1002, "火球术", "active", 2],
                [2001, "冰盾", "passive", 1]
            ]
            data_result = excel_update_range(
                "workflow_test.xlsx", "Skills!A3:D5", skills
            )
            workflow_steps.append(("add_data", data_result['success']))
            assert data_result['success']

            # 步骤4: 搜索特定技能
            search_result = excel_search(
                "workflow_test.xlsx", "火球术", "Skills", whole_word=True
            )
            workflow_steps.append(("search", search_result['success']))
            assert search_result['success']
            assert search_result['match_count'] == 2

            # 步骤5: 验证ID唯一性
            validation_result = excel_check_duplicate_ids(
                "workflow_test.xlsx", "Skills", id_column=1
            )
            workflow_steps.append(("validate_ids", not validation_result['has_duplicates']))
            assert not validation_result['has_duplicates']

            # 步骤6: 复制工作表数据对比
            excel_create_sheet("workflow_test.xlsx", "Skills_Copy")
            copy_result = excel_update_range(
                "workflow_test.xlsx", "Skills_Copy!A1:D5",
                headers + skills
            )
            workflow_steps.append(("copy_data", copy_result['success']))
            assert copy_result['success']

            # 步骤7: 比较工作表
            compare_result = excel_compare_sheets(
                "workflow_test.xlsx", "Skills",
                "workflow_test.xlsx", "Skills_Copy"
            )
            workflow_steps.append(("compare_sheets", compare_result['success']))
            assert compare_result['success']
            assert compare_result['data']['total_differences'] == 0

        finally:
            # 清理文件
            if os.path.exists("workflow_test.xlsx"):
                os.unlink("workflow_test.xlsx")

        # 验证所有步骤都成功
        failed_steps = [step for step, success in workflow_steps if not success]
        assert not failed_steps, f"失败的步骤: {failed_steps}"
```

#### 3. 外部系统集成测试
```python
class TestExternalIntegration:
    """测试与外部系统的集成"""

    def test_file_system_integration(self):
        """测试文件系统集成"""
        # 测试不同路径格式
        test_paths = [
            os.path.abspath("./test_file.xlsx"),
            os.path.join(os.getcwd(), "test_file.xlsx"),
            r"D:\temp\test_file.xlsx"
        ]

        for test_path in test_paths:
            with self.subTest(path=test_path):
                # 确保目录存在
                os.makedirs(os.path.dirname(test_path), exist_ok=True)

                # 创建文件
                result = excel_create_file(test_path, ["Test"])
                assert result['success']
                assert os.path.exists(test_path)

                # 清理
                os.unlink(test_path)

    def test_unicode_file_names(self):
        """测试Unicode文件名支持"""
        unicode_names = [
            "测试文件.xlsx",
            "файл.xlsx",
            "ファイル.xlsx",
            "ملف.xlsx"
        ]

        for file_name in unicode_names:
            with self.subTest(file_name=file_name):
                result = excel_create_file(file_name, ["测试"])
                assert result['success']

                # 验证中文工作表名
                update_result = excel_update_range(
                    file_name, "测试!A1", ["中文数据"]
                )
                assert update_result['success']

                read_result = excel_get_range(file_name, "测试!A1")
                assert read_result['success']
                assert read_result['data'][0][0] == "中文数据"

                os.unlink(file_name)
```

### 集成测试环境管理

#### 1. 测试环境配置
```python
# conftest.py中的集成测试配置
@pytest.fixture(scope="session")
def integration_test_environment():
    """集成测试环境配置"""
    # 创建测试目录
    test_dir = tempfile.mkdtemp(prefix="excel_integration_")

    # 设置环境变量
    original_env = os.environ.copy()
    os.environ['EXCELMCP_TEST_DIR'] = test_dir
    os.environ['EXCELMCP_TEST_MODE'] = 'true'

    yield test_dir

    # 清理环境
    os.environ.clear()
    os.environ.update(original_env)
    shutil.rmtree(test_dir, ignore_errors=True)

@pytest.fixture
def integration_test_files(integration_test_environment):
    """提供集成测试用的文件集合"""
    files = {}

    # 创建技能配置文件
    skills_file = os.path.join(integration_test_environment, "skills.xlsx")
    create_skills_config_file(skills_file)
    files['skills'] = skills_file

    # 创建装备配置文件
    equipment_file = os.path.join(integration_test_environment, "equipment.xlsx")
    create_equipment_config_file(equipment_file)
    files['equipment'] = equipment_file

    # 创建怪物配置文件
    monster_file = os.path.join(integration_test_environment, "monsters.xlsx")
    create_monster_config_file(monster_file)
    files['monsters'] = monster_file

    yield files

    # 清理文件
    for file_path in files.values():
        if os.path.exists(file_path):
            os.unlink(file_path)
```

#### 2. 数据库集成测试 (如果适用)
```python
class TestDatabaseIntegration:
    """测试与数据库的集成（如果有）"""

    @pytest.fixture
    def test_database(self):
        """创建测试数据库"""
        # 这里可以配置内存数据库或测试数据库
        from src.database import DatabaseManager

        db_manager = DatabaseManager(database_url="sqlite:///:memory:")
        db_manager.create_tables()

        yield db_manager

        db_manager.close()

    def test_excel_to_database_sync(self, test_database, temp_excel_file):
        """测试Excel到数据库的同步"""
        # 准备Excel数据
        create_test_skills_excel(temp_excel_file)

        # 同步到数据库
        sync_result = sync_excel_to_database(temp_excel_file, test_database)
        assert sync_result['success']

        # 验证数据库数据
        db_skills = test_database.get_all_skills()
        assert len(db_skills) > 0

        # 验证Excel和数据库数据一致性
        excel_data = excel_get_range(temp_excel_file, "Skills!A2:D100")
        assert len(excel_data['data']) == len(db_skills)
```

### 集成测试最佳实践

#### 1. 测试隔离和清理
```python
class TestIntegrationPractices:
    """集成测试最佳实践"""

    @pytest.fixture(autouse=True)
    def setup_test_isolation(self):
        """自动设置测试隔离"""
        # 测试前设置
        self.test_files = []
        self.temp_dirs = []

        yield

        # 测试后清理
        for file_path in self.test_files:
            if os.path.exists(file_path):
                os.unlink(file_path)

        for temp_dir in self.temp_dirs:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def create_temp_file(self, content=None, suffix=".xlsx"):
        """创建临时文件并注册清理"""
        file_path = tempfile.mktemp(suffix=suffix)
        self.test_files.append(file_path)

        if content:
            create_excel_file(file_path, content)

        return file_path

    def test_with_proper_cleanup(self):
        """演示正确的清理实践"""
        # 使用自动清理的临时文件
        test_file = self.create_temp_file()

        # 执行测试操作
        result = excel_create_file(test_file, ["Test"])
        assert result['success']

        # 文件会在fixture的cleanup中自动删除
```

#### 2. 测试数据和环境的复用
```python
@pytest.fixture(scope="module")
def shared_test_data():
    """模块级别的共享测试数据"""
    data_cache = {}

    # 创建一次，多次使用
    skills_file = tempfile.mktemp(suffix='.xlsx')
    create_comprehensive_skills_file(skills_file)
    data_cache['skills_file'] = skills_file

    yield data_cache

    # 模块结束时清理
    os.unlink(skills_file)

def test_multiple_operations_on_shared_data(shared_test_data):
    """在共享数据上执行多个操作"""
    skills_file = shared_test_data['skills_file']

    # 操作1: 搜索
    search_result = excel_search(skills_file, "火系", "Skills")
    assert search_result['success']

    # 操作2: 验证ID
    validate_result = excel_check_duplicate_ids(skills_file, "Skills")
    assert not validate_result['has_duplicates']

    # 操作3: 读取特定范围
    range_result = excel_get_range(skills_file, "Skills!A2:D10")
    assert range_result['success']
    assert len(range_result['data']) == 9  # 9行数据
```

---

## 故障排除指南

### 常见测试问题和解决方案

#### 1. 路径和环境问题

**问题**: Python路径和模块导入错误
```bash
# 错误信息
ModuleNotFoundError: No module named 'src.core.excel_reader'
ImportError: cannot import name 'ExcelReader' from 'src.core.excel_reader'
```

**解决方案**:
```python
# 方法1: 在conftest.py中添加路径
import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# 方法2: 使用PYTHONPATH环境变量
# 在测试前设置
os.environ['PYTHONPATH'] = str(project_root)

# 方法3: 使用推荐的方式运行测试
# python -m pytest tests/ -v  # 而不是 pytest tests/
```

**问题**: 临时文件清理失败
```python
# 问题代码
@pytest.fixture
def temp_file():
    temp_path = tempfile.mktemp(suffix='.xlsx')
    return temp_path  # 没有清理机制

# 解决方案
@pytest.fixture
def temp_file():
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        temp_path = f.name

    yield temp_path

    # 确保文件被清理
    try:
        if os.path.exists(temp_path):
            os.unlink(temp_path)
    except (OSError, PermissionError):
        # 忽略权限错误，Windows上常见
        pass
```

#### 2. Excel文件锁定问题

**问题**: Excel进程锁定文件
```bash
# 错误信息
PermissionError: [Errno 13] Permission denied: 'test_file.xlsx'
Exception: The process cannot access the file because it is being used by another process
```

**解决方案**:
```python
import time
import psutil

class ExcelFileHandler:
    @staticmethod
    def wait_for_file_release(file_path, timeout=10):
        """等待文件释放"""
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                # 尝试打开文件以检查是否被锁定
                with open(file_path, 'r+b') as f:
                    pass
                return True
            except (PermissionError, OSError):
                time.sleep(0.5)
                # 强制结束Excel进程 (仅测试环境)
                ExcelFileHandler.kill_excel_processes()

        return False

    @staticmethod
    def kill_excel_processes():
        """强制结束Excel进程（仅用于测试环境）"""
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and 'excel' in proc.info['name'].lower():
                    proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

# 在测试中使用
def test_with_excel_file_lock_handling(self, temp_file):
    """测试Excel文件锁定处理"""
    # 确保文件未被锁定
    assert ExcelFileHandler.wait_for_file_release(temp_file)

    try:
        # 执行Excel操作
        result = excel_get_range(temp_file, "Sheet1!A1:C1")
        assert result['success']
    finally:
        # 确保文件可被删除
        ExcelFileHandler.wait_for_file_release(temp_file)
```

#### 3. 内存和性能问题

**问题**: 测试过程中内存不足
```python
# 问题症状
MemoryError: Unable to allocate array
pytest crashed during execution
```

**解决方案**:
```python
import gc
import tracemalloc

class MemoryTestHelper:
    @staticmethod
    def setup_memory_tracking():
        """设置内存跟踪"""
        tracemalloc.start()
        gc.collect()  # 清理现有垃圾

    @staticmethod
    def get_memory_usage():
        """获取当前内存使用"""
        process = psutil.Process()
        return process.memory_info().rss / 1024 / 1024  # MB

    @staticmethod
    def check_memory_leak(test_function, max_increase_mb=50):
        """检查内存泄漏"""
        initial_memory = MemoryTestHelper.get_memory_usage()

        # 运行测试函数多次
        for _ in range(10):
            test_function()
            gc.collect()

        final_memory = MemoryTestHelper.get_memory_usage()
        memory_increase = final_memory - initial_memory

        if memory_increase > max_increase_mb:
            current, peak = tracemalloc.get_traced_memory()
            raise AssertionError(
                f"Possible memory leak detected. "
                f"Increased by {memory_increase:.2f}MB, "
                f"Peak usage: {peak / 1024 / 1024:.2f}MB"
            )

# 使用示例
def test_large_file_operations_without_memory_leak(self):
    """测试大文件操作不导致内存泄漏"""
    helper = MemoryTestHelper()
    helper.setup_memory_tracking()

    def large_file_operation():
        large_file = create_large_excel_file(10000)
        result = excel_get_range(large_file, "Sheet1!A1:AZ10000")
        os.unlink(large_file)
        return result

    helper.check_memory_leak(large_file_operation, max_increase_mb=100)
```

#### 4. 并发和竞争条件

**问题**: 并发测试不稳定
```python
# 问题症状
Flaky tests - sometimes pass, sometimes fail
Race conditions in file access
Timeout errors during concurrent execution
```

**解决方案**:
```python
import threading
import time
from concurrent.futures import ThreadPoolExecutor

class ConcurrencyTestHelper:
    @staticmethod
    def run_with_retry(func, max_retries=3, delay=0.1):
        """重试机制处理竞争条件"""
        for attempt in range(max_retries):
            try:
                return func()
            except (PermissionError, TimeoutError) as e:
                if attempt == max_retries - 1:
                    raise e
                time.sleep(delay * (2 ** attempt))  # 指数退避

    @staticmethod
    def synchronized_access_test(test_func, num_threads=5, operations_per_thread=10):
        """同步访问测试"""
        results = []
        errors = []
        lock = threading.Lock()

        def worker():
            try:
                for i in range(operations_per_thread):
                    result = ConcurrencyTestHelper.run_with_retry(
                        lambda: test_func(i)
                    )
                    with lock:
                        results.append(result)
            except Exception as e:
                with lock:
                    errors.append(e)

        # 启动线程
        threads = []
        for _ in range(num_threads):
            thread = threading.Thread(target=worker)
            threads.append(thread)
            thread.start()

        # 等待完成
        for thread in threads:
            thread.join()

        return results, errors

# 使用示例
def test_concurrent_file_access(self):
    """测试并发文件访问"""
    temp_file = self.create_temp_file()

    def file_operation(operation_id):
        # 每个线程使用不同的范围避免冲突
        range_expr = f"Sheet1!A{operation_id}:Z{operation_id}"
        data = [[f"Data_{operation_id}_{i}" for i in range(26)]]

        return excel_update_range(temp_file, range_expr, data)

    results, errors = ConcurrencyTestHelper.synchronized_access_test(
        file_operation, num_threads=3, operations_per_thread=5
    )

    assert not errors, f"并发操作出现错误: {errors}"
    assert len(results) == 15  # 3线程 * 5操作
    assert all(result['success'] for result in results)
```

### 调试和诊断工具

#### 1. 测试调试工具
```python
import logging
import pdb

class TestDebugger:
    @staticmethod
    def enable_debug_logging():
        """启用调试日志"""
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('test_debug.log'),
                logging.StreamHandler()
            ]
        )

    @staticmethod
    def debug_test_function():
        """调试测试函数的装饰器"""
        def decorator(func):
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                print(f"\n=== 调试测试: {func.__name__} ===")
                print(f"参数: args={args}, kwargs={kwargs}")

                try:
                    result = func(*args, **kwargs)
                    print(f"结果: {result}")
                    return result
                except Exception as e:
                    print(f"异常: {type(e).__name__}: {e}")
                    print("进入调试器...")
                    pdb.set_trace()
                    raise
            return wrapper
        return decorator

# 使用示例
@TestDebugger.debug_test_function()
def test_failing_function(self):
    """调试失败的测试函数"""
    # 这个函数会在失败时启动调试器
    result = some_complex_operation()
    assert result['success'], f"操作失败: {result}"
```

#### 2. 性能分析工具
```python
import cProfile
import pstats
from line_profiler import LineProfiler

class TestProfiler:
    @staticmethod
    def profile_test_function(func, *args, **kwargs):
        """分析测试函数性能"""
        profiler = cProfile.Profile()
        profiler.enable()

        try:
            result = func(*args, **kwargs)
        finally:
            profiler.disable()

        # 保存分析结果
        stats = pstats.Stats(profiler)
        stats.sort_stats('cumulative')
        stats.print_stats(20)  # 显示前20个最耗时的函数

        return result

    @staticmethod
    def line_profile_function(func, *args, **kwargs):
        """行级性能分析"""
        profiler = LineProfiler()
        profiler_wrapper = profiler(func)

        result = profiler_wrapper(*args, **kwargs)
        profiler.print_stats()

        return result

# 使用示例
def test_performance_analysis(self):
    """性能分析示例"""
    def slow_operation():
        # 模拟慢操作
        time.sleep(0.1)
        return {"success": True}

    # 使用性能分析
    TestProfiler.profile_test_function(slow_operation)

    # 或者使用行级分析
    TestProfiler.line_profile_function(slow_operation)
```

#### 3. 测试数据验证工具
```python
class TestDataValidator:
    """测试数据验证工具"""

    @staticmethod
    def validate_excel_file(file_path, expected_sheets=None, min_rows=None):
        """验证Excel文件结构"""
        validation_result = {
            'valid': True,
            'errors': [],
            'warnings': []
        }

        try:
            # 检查文件存在
            if not os.path.exists(file_path):
                validation_result['valid'] = False
                validation_result['errors'].append(f"文件不存在: {file_path}")
                return validation_result

            # 检查工作表
            sheets_result = excel_list_sheets(file_path)
            if not sheets_result['success']:
                validation_result['valid'] = False
                validation_result['errors'].append("无法读取工作表列表")
                return validation_result

            actual_sheets = sheets_result['sheets']

            # 验证期望的工作表
            if expected_sheets:
                missing_sheets = set(expected_sheets) - set(actual_sheets)
                if missing_sheets:
                    validation_result['valid'] = False
                    validation_result['errors'].append(f"缺少工作表: {missing_sheets}")

            # 检查数据行数
            if min_rows:
                for sheet_name in actual_sheets:
                    last_row = excel_find_last_row(file_path, sheet_name)
                    if last_row['success'] and last_row['last_row'] < min_rows:
                        validation_result['warnings'].append(
                            f"工作表 {sheet_name} 数据行数少于 {min_rows}"
                        )

        except Exception as e:
            validation_result['valid'] = False
            validation_result['errors'].append(f"验证过程中出现异常: {e}")

        return validation_result

# 使用示例
def test_with_data_validation(self):
    """使用数据验证的测试"""
    temp_file = self.create_temp_file()

    # 创建文件
    result = excel_create_file(temp_file, ["Sheet1", "Sheet2"])
    assert result['success']

    # 添加数据
    excel_update_range(temp_file, "Sheet1!A1:C10",
                      [[f"Data_{i}_{j}" for j in range(3)] for i in range(10)])

    # 验证文件
    validation = TestDataValidator.validate_excel_file(
        temp_file,
        expected_sheets=["Sheet1", "Sheet2"],
        min_rows=10
    )

    assert validation['valid'], f"文件验证失败: {validation['errors']}"
```

### 测试故障排除检查清单

#### 1. 环境检查清单
- [ ] Python版本是否正确 (≥3.10)
- [ ] 依赖包是否正确安装 (`pip install -e .`)
- [ ] PYTHONPATH是否正确设置
- [ ] 测试文件权限是否正确
- [ ] 临时目录是否可写

#### 2. 文件系统检查清单
- [ ] Excel进程是否已关闭
- [ ] 测试文件是否被锁定
- [ ] 临时文件是否正确清理
- [ ] 文件路径格式是否正确
- [ ] Unicode文件名是否支持

#### 3. 内存和性能检查清单
- [ ] 大文件测试是否分批处理
- [ ] 是否有内存泄漏
- [ ] 垃圾回收是否及时
- [ ] 测试超时设置是否合理
- [ ] 并发测试是否同步正确

#### 4. 数据完整性检查清单
- [ ] 测试数据是否正确初始化
- [ ] 测试数据是否正确清理
- [ ] 文件格式是否正确
- [ ] 数据类型是否匹配
- [ ] 边界条件是否覆盖

---

## 总结

本测试指南提供了Excel MCP Server项目的全面测试框架和最佳实践。通过遵循这些指导原则，开发者可以：

1. **构建高质量的测试套件**，确保代码质量和系统稳定性
2. **实施有效的测试策略**，从单元测试到集成测试的全覆盖
3. **维护测试环境的稳定性**，通过正确的隔离和清理机制
4. **监控和优化性能**，通过系统化的性能测试和监控
5. **快速诊断和解决问题**，通过丰富的调试和故障排除工具

记住，好的测试不仅是发现问题的工具，更是确保代码质量和可维护性的重要投资。定期审查和改进测试策略，将有助于项目的长期成功。

**关键建议**:
- 保持测试的简单性和可读性
- 优先测试重要的业务逻辑
- 使用适当的Mock策略平衡测试速度和可靠性
- 建立持续集成的测试流水线
- 定期分析测试覆盖率和质量指标
- 保持测试数据的独立性和可重现性