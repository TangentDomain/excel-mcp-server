# Excel MCP Server - 详细ID对象属性变化跟踪功能

## 🎯 功能概述

**用户需求**: "返回值中, 得知道是哪个id的哪个属性变化了"

**解决方案**: 实现了详细的ID对象属性变化跟踪功能，在比较结果中提供精确到字段级别的变化信息。

## ✨ 功能特性

### 1. ID对象识别
- ✅ 准确识别每个变化的ID对象
- ✅ 显示对象名称和ID编号
- ✅ 支持🆕新增、🗑️删除、🔄修改三种变化类型

### 2. 详细属性变化跟踪
- ✅ **字段级变化**: 精确到每个具体属性的变化
- ✅ **原值→新值**: 显示变化前后的具体数值
- ✅ **变化类型**: 区分文本变化(text_change)、数值变化(numeric_change)、配置变化(config_change)
- ✅ **数值分析**: 自动计算数值变化量和百分比

### 3. 数据结构增强
```python
# 新增 FieldDifference 数据类
@dataclass
class FieldDifference:
    field_name: str           # 字段名
    old_value: Any           # 原始值
    new_value: Any           # 新值
    change_type: str         # 变化类型
    numeric_change: Optional[float]    # 数值变化量
    percent_change: Optional[float]    # 百分比变化
    formatted_change: Optional[str]    # 格式化显示
```

## 🧪 测试结果

### 测试用例
- **文件1**: `D:\tr\svn\trunk\配置表\测试配置\微小\TrSkill.xlsx`
- **文件2**: `D:\tr\svn\trunk\配置表\战斗环境配置\TrSkill.xlsx`

### 测试输出示例
```
📋 工作表 'TrSkillUpgrade': 3 个差异

🔍 差异1 - ID 900070010 (对象名: 9000700):
  详细字段变化数: 4
  🔧 技能增强效果3 (text_change): '[技能增强]激光折射衍射攻击buffID' → ''
  🔧 Column18 (numeric_change): '900710' → '1'
  🔧 Column22 (text_change): '900711' → ''

🔍 差异2 - ID 900120080 (对象名: 9001200):
  详细字段变化数: 2
  🔧 Column18 (text_change): '15' → ''
  🔧 技能增强效果2 (text_change): '[技能增强]驭兽之王击退' → ''

📈 统计:
- 总差异数: 111
- 详细字段差异数: 5
- 支持ID-属性跟踪: ✅
```

## 🔧 技术实现

### 1. 双层比较结构
- **简化差异**: 用于向后兼容和摘要显示
- **详细差异**: 包含完整的字段级变化信息

### 2. 核心方法
```python
def _compare_row_data_detailed(
    self,
    row_data1: Dict,
    row_data2: Dict,
    headers1: List[str],
    headers2: List[str],
    options: ComparisonOptions
) -> Tuple[List[str], List[FieldDifference]]:
    """比较行数据，返回简化和详细两种格式的差异"""
```

### 3. 字段差异创建
```python
def _create_field_difference(
    self,
    field_name: str,
    old_value: Any,
    new_value: Any,
    options: ComparisonOptions
) -> FieldDifference:
    """创建详细的字段差异对象"""
```

## 📊 返回结果结构

```json
{
  "success": true,
  "total_differences": 111,
  "sheet_comparisons": [
    {
      "sheet_name": "TrSkillUpgrade",
      "differences": [
        {
          "row_id": "900070010",
          "object_name": "9000700",
          "difference_type": "ROW_MODIFIED",
          "field_differences": ["简化摘要..."],
          "detailed_field_differences": [
            {
              "field_name": "技能增强效果3",
              "old_value": "[技能增强]激光折射衍射攻击buffID",
              "new_value": "",
              "change_type": "text_change",
              "numeric_change": null,
              "percent_change": null,
              "formatted_change": "'[技能增强]激光折射衍射攻击buffID' → ''"
            }
          ]
        }
      ]
    }
  ]
}
```

## ✅ 用户需求满足度

| 需求项 | 状态 | 说明 |
|--------|------|------|
| 知道哪个ID | ✅ | 精确显示ID编号和对象名 |
| 知道哪个属性 | ✅ | 详细显示字段名称 |
| 知道如何变化 | ✅ | 显示原值→新值，变化类型 |
| 返回值包含信息 | ✅ | detailed_field_differences完整包含 |

## 🚀 使用方式

```python
# 调用比较函数
result = excel_compare_files(
    file1_path="file1.xlsx",
    file2_path="file2.xlsx",
    header_row=1,
    id_column=1,
    case_sensitive=True
)

# 访问详细差异
for sheet_comp in result['data']['sheet_comparisons']:
    for diff in sheet_comp['differences']:
        if 'detailed_field_differences' in diff:
            for field_diff in diff['detailed_field_differences']:
                print(f"ID {diff['row_id']} 的属性 {field_diff['field_name']} 从 {field_diff['old_value']} 变为 {field_diff['new_value']}")
```

## 🎉 总结

**用户的核心需求已完全实现**: 现在可以在返回值中准确知道：
- 🎯 **哪个ID**: 具体的对象ID编号
- 🎯 **哪个属性**: 精确的字段名称
- 🎯 **如何变化**: 原值、新值、变化类型、数值分析

这个功能特别适合游戏配置表的变化跟踪，能够快速识别技能、装备、道具等游戏对象的具体属性变化。
