# 📝 Excel MCP Server 测试命名规范

> **测试命名规范专题文档**

## 文件命名规范

### Python测试文件
```python
# 标准格式：test_模块名_功能.py
test_excel_operations.py      # Excel操作测试
test_sql_queries.py           # SQL查询测试  
test_mcp_tools.py             # MCP工具测试
test_utils.py                 # 工具函数测试

# 特定功能格式：test_功能_场景.py
test_skill_balance.py         # 技能平衡测试
test_equipment_rarity.py      # 装备稀有度测试
test_monster_ai.py           # 怪物AI行为测试
```

### 测试数据文件
```python
# 数据文件：test_data_场景.xlsx
test_data_skills.xlsx        # 技能测试数据
test_data_equipment.xlsx     # 装备测试数据
test_data_monsters.xlsx      # 怪物测试数据
```

### 测试报告文件
```python
# 报告文件：report_日期_测试类型.html
report_20260329_unit.html    # 单元测试报告
report_20260329_integration.html # 集成测试报告
report_20260329_e2e.html     # 端到端测试报告
```

## 函数命名规范

### 测试类命名
```python
# 格式：Test被测模块
class TestExcelOperations:
    pass

class TestSQLQueries:
    pass

class TestMCPTools:
    pass

# 格式：Test功能_场景
class TestSkillBalance:
    pass

class TestEquipmentRarity:
    pass
```

### 测试方法命名
```python
# 标准格式：test_预期行为_条件
def test_load_workbook_success(self):
    pass

def test_load_workbook_file_not_found(self):
    pass

def test_query_join_multiple_tables(self):
    pass

def test_insert_row_valid_data(self):
    pass

def test_insert_row_duplicate_key(self):
    pass

# 边界条件测试
def test_load_empty_workbook(self):
    pass

def test_query_no_results(self):
    pass

def test_update_nonexistent_row(self):
    pass
```

### 测试固件命名
```python
# setup/teardown
def setup_method(self):
    pass

def teardown_method(self):
    pass

# 类级别固件
def setup_class(self):
    pass

def teardown_class(self):
    pass

# 模块级别固件
def setup_module(self):
    pass

def teardown_module(self):
    pass
```

## 变量命名规范

### 测试数据变量
```python
# 测试数据
test_skill_data = {...}
equipment_test_records = [...]
monster_config_template = {...}

# 预期结果
expected_result = {...}
expected_count = 42
expected_skill_name = "Fireball"

# 实际结果
actual_result = {...}
actual_count = 0
actual_skill_name = ""

# 测试固件
test_workbook = None
test_session = None
test_client = None
```

### Mock对象变量
```python
# Mock对象
mock_excel_reader = None
mock_database_client = None
mock_mcp_server = None

# Spy对象
skill_update_spy = None
equipment_query_spy = None
```

### 常量定义
```python
# 测试常量
TEST_SKILL_ID = "skill_001"
TEST_EQUIPMENT_ID = "equip_001"
TEST_MONSTER_ID = "monster_001"

# 测试数据路径
TEST_DATA_DIR = "tests/data/"
TEST_SKILLS_FILE = "tests/data/test_data_skills.xlsx"
TEST_EQUIPMENT_FILE = "tests/data/test_data_equipment.xlsx"

# 性能测试常量
PERFORMANCE_THRESHOLD = 1000  # 毫秒
LARGE_DATA_SIZE = 10000      # 大数据量阈值
```

## 断言命名规范

### 标准断言
```python
# 使用assert关键字，清晰表达预期
assert result is not None
assert len(result) == expected_count
assert result['skill_name'] == expected_skill_name
assert isinstance(result, dict)

# 带消息的断言
assert len(result) > 0, "查询结果不应为空"
assert skill_level >= 1, f"技能等级不能小于1，当前值: {skill_level}"
```

### 异常测试断言
```python
# 使用pytest.raises
with pytest.raises(FileNotFoundError):
    ExcelReader.load_workbook("nonexistent.xlsx")

with pytest.raises(ValueError) as exc_info:
    SkillManager.create_skill invalid_skill_data)
    assert "技能名称不能为空" in str(exc_info.value)
```

### 复杂断言
```python
# 多条件断言
assert result['status'] == 'success'
assert len(result['data']) >= 1
assert all(isinstance(item, dict) for item in result['data'])

# JSON结构断言
assert expected_json_structure == actual_result, f"JSON结构不匹配: {actual_result}"
```

## 注释规范

### 测试方法注释
```python
def test_skill_balance_calculation(self):
    """
    测试技能平衡计算功能
    
    验证技能伤害计算是否正确考虑：
    - 基础伤害
    - 等级加成
    - 属性修正
    - 装备加成
    
    Returns:
        None
    """
    pass
```

### 复杂逻辑注释
```python
def test_advanced_sql_query(self):
    """
    测试复杂SQL查询功能
    
    包括：
    1. 多表JOIN查询
    2. 子查询嵌套
    3. GROUP BY聚合
    4. HAVING过滤
    
    验证查询结果的准确性和性能
    """
    # 构建测试数据
    test_skills = create_test_skills_data()
    test_classes = create_test_classes_data()
    
    # 执行查询
    result = self.mcp_client.execute_query("""
        SELECT s.*, c.class_name, 
               AVG(s.damage) as avg_damage
        FROM skills s
        JOIN classes c ON s.class_id = c.id
        WHERE s.level > 5
        GROUP BY s.id, c.class_name
        HAVING AVG(s.damage) > 50
    """)
    
    # 验证结果
    assert len(result) > 0
    assert 'class_name' in result[0]
    assert 'avg_damage' in result[0]
```

### 测试数据注释
```python
# 测试数据定义
SKILL_TEST_DATA = {
    # 基础技能数据
    'basic_skill': {
        'name': 'Fireball',
        'damage': 100,
        'level': 1,
        'element': 'fire'
    },
    # 高级技能数据  
    'advanced_skill': {
        'name': 'Meteor Shower',
        'damage': 500,
        'level': 10,
        'element': 'fire',
        'area_effect': True
    }
}
```

## 测试模板

### 标准测试模板
```python
class TestSkillManager:
    """技能管理器测试类"""
    
    def setup_method(self):
        """测试方法初始化"""
        self.skill_manager = SkillManager()
        self.test_skill_data = {...}
    
    def test_create_skill_success(self):
        """测试创建技能成功"""
        # 准备测试数据
        skill_data = self.test_skill_data
        
        # 执行测试
        result = self.skill_manager.create_skill(skill_data)
        
        # 验证结果
        assert result is True
        assert self.skill_manager.skill_exists(skill_data['name'])
    
    def test_create_skill_duplicate_name(self):
        """测试创建重复名称技能"""
        # 准备测试数据
        skill_data = self.test_skill_data
        
        # 执行测试
        self.skill_manager.create_skill(skill_data)  # 第一次创建
        result = self.skill_manager.create_skill(skill_data)  # 第二次创建
        
        # 验证结果
        assert result is False
        assert self.skill_manager.get_error_message() == "技能名称已存在"
```

### 性能测试模板
```python
def test_large_dataset_performance():
    """测试大数据集处理性能"""
    # 准备测试数据
    large_dataset = generate_large_test_data(10000)
    
    # 执行性能测试
    start_time = time.time()
    result = excel_processor.process_large_dataset(large_dataset)
    end_time = time.time()
    
    # 验证结果
    processing_time = end_time - start_time
    assert processing_time < PERFORMANCE_THRESHOLD
    assert len(result) == len(large_dataset)
```

---

*本文档是测试指南系列专题之一，更多内容请查看相关专题文档*