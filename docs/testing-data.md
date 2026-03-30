# 📊 Excel MCP Server 测试数据管理

> **测试数据管理专题文档**

## 测试数据结构

### 标准测试数据目录
```
tests/
├── data/                           # 测试数据根目录
│   ├── normal/                     # 正常测试数据
│   │   ├── skills.xlsx            # 技能测试数据
│   │   ├── equipment.xlsx         # 装备测试数据
│   │   ├── monsters.xlsx          # 怪物测试数据
│   │   └── classes.xlsx           # 职业测试数据
│   ├── edge/                      # 边界测试数据
│   │   ├── empty.xlsx             # 空文件
│   │   ├── large.xlsx             # 大文件(10万+行)
│   │   └── invalid_format.xlsx     # 格式错误文件
│   ├── invalid/                   # 无效测试数据
│   │   ├── duplicate_keys.xlsx    # 重复键数据
│   │   ├── missing_columns.xlsx    # 缺失列数据
│   │   └── wrong_types.xlsx       # 类型错误数据
│   └── templates/                 # 测试数据模板
│       ├── skill_template.xlsx     # 技能模板
│       ├── equipment_template.xlsx # 装备模板
│       └── monster_template.xlsx   # 怪物模板
└── fixtures/                       # 测试固件
    ├── excel_reader.py             # Excel读取器固件
    ├── mcp_client.py              # MCP客户端固件
    └── database_client.py         # 数据库客户端固件
```

### 测试数据分类

#### 1. 功能测试数据
```python
# 标准功能数据
FUNCTIONAL_TEST_DATA = {
    'skills': {
        'normal': [
            {'id': 'skill_001', 'name': 'Fireball', 'damage': 100, 'level': 1},
            {'id': 'skill_002', 'name': 'Ice Arrow', 'damage': 80, 'level': 1},
            {'id': 'skill_003', 'name': 'Lightning Strike', 'damage': 120, 'level': 2}
        ],
        'edge': [
            {'id': 'skill_empty', 'name': '', 'damage': 0, 'level': 0},
            {'id': 'skill_max', 'name': 'Ultimate Skill', 'damage': 9999, 'level': 99}
        ]
    },
    'equipment': {
        'normal': [
            {'id': 'equip_001', 'name': 'Iron Sword', 'rarity': 'common', 'power': 50},
            {'id': 'equip_002', 'name': 'Magic Staff', 'rarity': 'rare', 'power': 100}
        ]
    }
}
```

#### 2. 性能测试数据
```python
# 大数据集生成
def generate_large_skill_dataset(size=10000):
    """生成大量技能测试数据"""
    return [
        {
            'id': f'skill_{i:03d}',
            'name': f'Skill {i}',
            'damage': random.randint(10, 500),
            'level': random.randint(1, 50),
            'element': random.choice(['fire', 'water', 'earth', 'air']),
            'cooldown': random.randint(1, 30)
        }
        for i in range(size)
    ]

# 压力测试数据
STRESS_TEST_DATA = {
    'concurrent_users': 100,
    'requests_per_user': 50,
    'data_size_per_request': 1000
}
```

#### 3. 边界测试数据
```python
# 边界值数据
BOUNDARY_TEST_DATA = {
    'skills': {
        'min_values': {'damage': 0, 'level': 0, 'cooldown': 0},
        'max_values': {'damage': 9999, 'level': 99, 'cooldown': 999},
        'null_values': {'name': None, 'damage': None, 'level': None}
    },
    'equipment': {
        'empty_strings': {'name': '', 'description': ''},
        'max_length': {'name': 'a' * 255, 'description': 'a' * 1000}
    }
}
```

## 测试数据管理工具

### 数据生成器
```python
class TestDataGenerator:
    """测试数据生成器"""
    
    @staticmethod
    def generate_skill_data(count=10):
        """生成技能测试数据"""
        skills = []
        for i in range(count):
            skill = {
                'id': f'skill_{i+1:03d}',
                'name': f'技能_{i+1}',
                'damage': random.randint(50, 200),
                'level': random.randint(1, 10),
                'element': random.choice(['fire', 'water', 'earth', 'wind']),
                'mana_cost': random.randint(10, 100),
                'cooldown': random.randint(1, 10)
            }
            skills.append(skill)
        return skills
    
    @staticmethod
    def generate_equipment_data(count=10):
        """生成装备测试数据"""
        equipment = []
        rarities = ['common', 'uncommon', 'rare', 'epic', 'legendary']
        
        for i in range(count):
            item = {
                'id': f'equip_{i+1:03d}',
                'name': f'装备_{i+1}',
                'rarity': random.choice(rarities),
                'power': random.randint(10, 500),
                'durability': random.randint(50, 200),
                'price': random.randint(100, 10000)
            }
            equipment.append(item)
        return equipment
```

### 数据验证器
```python
class DataValidator:
    """测试数据验证器"""
    
    @staticmethod
    def validate_skill_data(data):
        """验证技能数据"""
        errors = []
        
        for skill in data:
            # 验证必需字段
            if 'name' not in skill or not skill['name']:
                errors.append(f"技能缺少名称: {skill}")
            
            if 'damage' not in skill or not isinstance(skill['damage'], (int, float)):
                errors.append(f"技能伤害值无效: {skill}")
            
            if 'level' not in skill or not isinstance(skill['level'], int):
                errors.append(f"技能等级无效: {skill}")
            
            # 验证数值范围
            if skill.get('damage', 0) < 0 or skill.get('damage', 0) > 9999:
                errors.append(f"技能伤害超出范围: {skill['damage']}")
            
            if skill.get('level', 0) < 1 or skill.get('level', 0) > 99:
                errors.append(f"技能等级超出范围: {skill['level']}")
        
        return errors
    
    @staticmethod
    def validate_equipment_data(data):
        """验证装备数据"""
        errors = []
        valid_rarities = ['common', 'uncommon', 'rare', 'epic', 'legendary']
        
        for item in data:
            if 'rarity' not in item or item['rarity'] not in valid_rarities:
                errors.append(f"装备稀有度无效: {item.get('rarity', 'unknown')}")
            
            if 'power' not in item or not isinstance(item['power'], (int, float)):
                errors.append(f"装备战力值无效: {item}")
        
        return errors
```

### 数据清理工具
```python
class TestDataCleaner:
    """测试数据清理工具"""
    
    @staticmethod
    def clean_test_data(data, data_type):
        """清理测试数据"""
        if data_type == 'skills':
            return TestDataCleaner._clean_skill_data(data)
        elif data_type == 'equipment':
            return TestDataCleaner._clean_equipment_data(data)
        else:
            return data
    
    @staticmethod
    def _clean_skill_data(data):
        """清理技能数据"""
        cleaned = []
        for skill in data:
            cleaned_skill = {
                'id': skill.get('id', ''),
                'name': skill.get('name', ''),
                'damage': skill.get('damage', 0),
                'level': skill.get('level', 1),
                'element': skill.get('element', ''),
                'mana_cost': skill.get('mana_cost', 0),
                'cooldown': skill.get('cooldown', 1)
            }
            cleaned.append(cleaned_skill)
        return cleaned
    
    @staticmethod
    def _clean_equipment_data(data):
        """清理装备数据"""
        cleaned = []
        for item in data:
            cleaned_item = {
                'id': item.get('id', ''),
                'name': item.get('name', ''),
                'rarity': item.get('rarity', 'common'),
                'power': item.get('power', 0),
                'durability': item.get('durability', 100),
                'price': item.get('price', 0)
            }
            cleaned.append(cleaned_item)
        return cleaned
```

## 测试数据加载

### 数据加载器
```python
class TestDataLoader:
    """测试数据加载器"""
    
    def __init__(self, data_dir="tests/data"):
        self.data_dir = Path(data_dir)
    
    def load_excel_data(self, filename, sheet_name=0):
        """加载Excel测试数据"""
        filepath = self.data_dir / filename
        if not filepath.exists():
            raise FileNotFoundError(f"测试数据文件不存在: {filepath}")
        
        try:
            df = pd.read_excel(filepath, sheet_name=sheet_name)
            return df.to_dict('records')
        except Exception as e:
            raise ValueError(f"加载Excel数据失败: {e}")
    
    def load_json_data(self, filename):
        """加载JSON测试数据"""
        filepath = self.data_dir / filename
        if not filepath.exists():
            raise FileNotFoundError(f"测试数据文件不存在: {filepath}")
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            raise ValueError(f"加载JSON数据失败: {e}")
    
    def load_yaml_data(self, filename):
        """加载YAML测试数据"""
        filepath = self.data_dir / filename
        if not filepath.exists():
            raise FileNotFoundError(f"测试数据文件不存在: {filepath}")
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            raise ValueError(f"加载YAML数据失败: {e}")
```

### 数据缓存管理
```python
class TestDataCache:
    """测试数据缓存管理"""
    
    def __init__(self):
        self.cache = {}
        self.cache_ttl = 3600  # 1小时缓存
    
    def get_data(self, data_type, data_id=None):
        """获取缓存数据"""
        cache_key = f"{data_type}_{data_id}" if data_id else data_type
        
        if cache_key in self.cache:
            cached_data, timestamp = self.cache[cache_key]
            if time.time() - timestamp < self.cache_ttl:
                return cached_data
            else:
                del self.cache[cache_key]
        
        return None
    
    def set_data(self, data_type, data, data_id=None):
        """设置缓存数据"""
        cache_key = f"{data_type}_{data_id}" if data_id else data_type
        self.cache[cache_key] = (data, time.time())
    
    def clear_cache(self):
        """清空缓存"""
        self.cache.clear()
```

## 测试数据使用示例

### 单元测试数据使用
```python
class TestSkillManager:
    def setup_method(self):
        """测试初始化"""
        self.loader = TestDataLoader()
        self.test_skills = self.loader.load_excel_data("normal/skills.xlsx")
    
    def test_skill_creation(self):
        """测试技能创建"""
        # 加载测试数据
        skill_data = self.test_skills[0]
        
        # 执行测试
        result = self.skill_manager.create_skill(skill_data)
        
        # 验证结果
        assert result is True
        assert self.skill_manager.get_skill(skill_data['id']) is not None
    
    def test_skill_validation(self):
        """测试技能验证"""
        # 加载边界测试数据
        edge_skills = self.loader.load_excel_data("edge/skills.xlsx")
        
        for skill in edge_skills:
            result = self.skill_manager.validate_skill(skill)
            # 根据预期结果验证
            if skill['name'] == '':
                assert result is False
```

### 集成测试数据使用
```python
class TestMCPIntegration:
    def setup_class(self):
        """类级别初始化"""
        self.loader = TestDataLoader()
        self.test_skills = self.loader.load_excel_data("normal/skills.xlsx")
        self.test_equipment = self.loader.load_excel_data("normal/equipment.xlsx")
    
    def test_cross_table_query(self):
        """测试跨表查询"""
        # 构建测试数据
        test_db = {
            'skills': self.test_skills,
            'equipment': self.test_equipment
        }
        
        # 执行查询
        result = self.mcp_client.query(
            "SELECT s.*, e.name as equipment_name "
            "FROM skills s "
            "LEFT JOIN equipment e ON s.equipment_id = e.id "
            "WHERE s.damage > 50"
        )
        
        # 验证结果
        assert len(result) > 0
        assert 'equipment_name' in result[0]
```

### 性能测试数据使用
```python
class TestPerformance:
    def setup_method(self):
        """性能测试初始化"""
        self.generator = TestDataGenerator()
        self.large_dataset = self.generator.generate_skill_data(10000)
    
    def test_large_dataset_processing(self):
        """测试大数据集处理"""
        # 加载大数据集
        large_data = self.large_dataset
        
        # 执行性能测试
        start_time = time.time()
        result = self.skill_processor.batch_process_skills(large_data)
        end_time = time.time()
        
        # 验证性能
        processing_time = end_time - start_time
        assert processing_time < 10.0  # 10秒内完成
        assert len(result) == len(large_data)
```

## 测试数据维护

### 数据更新策略
```python
class TestDataMaintenance:
    """测试数据维护工具"""
    
    @staticmethod
    def update_test_data(data_type, new_data):
        """更新测试数据"""
        data_path = Path(f"tests/data/{data_type}")
        
        # 备份现有数据
        backup_path = data_path.with_suffix('.backup.xlsx')
        if data_path.exists():
            shutil.copy2(data_path, backup_path)
        
        # 写入新数据
        df = pd.DataFrame(new_data)
        df.to_excel(data_path, index=False)
    
    @staticmethod
    def validate_test_data_integrity():
        """验证测试数据完整性"""
        data_dir = Path("tests/data")
        issues = []
        
        for file in data_dir.glob("*.xlsx"):
            try:
                df = pd.read_excel(file)
                # 检查必需列
                if 'id' not in df.columns:
                    issues.append(f"文件 {file.name} 缺少id列")
                
                # 检查数据完整性
                if df.isnull().sum().sum() > 0:
                    issues.append(f"文件 {file.name} 包含空值")
                    
            except Exception as e:
                issues.append(f"文件 {file.name} 读取失败: {e}")
        
        return issues
```

---

*本文档是测试指南系列专题之一，更多内容请查看相关专题文档*