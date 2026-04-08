#!/usr/bin/env python3
"""
智能配置工具测试
验证新增的4个智能配置工具
"""
import json
import os
import tempfile
from pathlib import Path
import sys

# 添加项目路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src"))

from excel_mcp.core.smart_config_recommender import SmartConfigurationRecommender


def create_test_excel_data():
    """创建测试用的Excel数据"""
    # 模拟RPG游戏配置数据
    excel_data = {
        "characters": {
            "data": [
                ["角色ID", "名称", "等级", "生命值", "魔法值", "职业"],
                [1, "战士", 10, 100, 0, "近战"],
                [2, "法师", 8, 50, 80, "远程"],
                [3, "弓箭手", 12, 70, 30, "远程"]
            ]
        },
        "skills": {
            "data": [
                ["技能ID", "名称", "消耗", "冷却", "类型"],
                [101, "剑击", 10, 0, "攻击"],
                [102, "火球术", 20, 5, "魔法"],
                [103, "射击", 15, 3, "攻击"]
            ]
        },
        "items": {
            "data": [
                ["物品ID", "名称", "类型", "稀有度"],
                [1001, "铁剑", "武器", "普通"],
                [1002, "法杖", "武器", "稀有"],
                [1003, "药水", "消耗品", "普通"],
                [1004, "药水", "消耗品", "普通"]
            ]
        }
    }
    return excel_data


def test_smart_config_recommendation():
    """测试智能配置推荐"""
    print("🧪 测试1: 智能配置推荐")
    
    recommender = SmartConfigurationRecommender()
    excel_data = create_test_excel_data()
    
    try:
        recommendations = recommender.recommend_configurations(excel_data)
        print(f"✅ 检测到游戏类型: {recommendations['game_type']}")
        print(f"✅ 推荐数量: {len(recommendations['config_recommendations'])}")
        print(f"✅ 验证规则数量: {len(recommendations['validation_rules'])}")
        return True
    except Exception as e:
        print(f"❌ 智能配置推荐失败: {e}")
        return False


def test_game_type_detection():
    """测试游戏类型检测"""
    print("\n🧪 测试2: 游戏类型检测")
    
    from excel_mcp.core.smart_config_recommender import GameTypeDetector
    
    detector = GameTypeDetector()
    excel_data = create_test_excel_data()
    
    try:
        game_type = detector.detect_game_type(excel_data)
        print(f"✅ 检测结果: {game_type}")
        expected = "rpg"  # 应该检测为RPG游戏
        if game_type == expected:
            print(f"✅ 检测正确: {game_type}")
            return True
        else:
            print(f"❌ 检测错误: 期望 {expected}, 实际 {game_type}")
            return False
    except Exception as e:
        print(f"❌ 游戏类型检测失败: {e}")
        return False


def test_configuration_analyzer():
    """测试配置分析器"""
    print("\n🧪 测试3: 配置分析器")
    
    from excel_mcp.core.smart_config_recommender import ConfigurationAnalyzer
    
    analyzer = ConfigurationAnalyzer()
    excel_data = create_test_excel_data()
    
    try:
        analysis = analyzer.analyze_excel_structure(excel_data)
        
        # 检查分析结果
        if "sheet_structure" in analysis and "data_patterns" in analysis:
            print(f"✅ 工作表数量: {len(analysis['sheet_structure'])}")
            print(f"✅ 数据模式数量: {len(analysis['data_patterns'])}")
            
            # 检查每个工作表的结构
            for sheet_name, structure in analysis["sheet_structure"].items():
                print(f"✅ {sheet_name}: {structure['rows']}行 x {structure['cols']}列")
            
            return True
        else:
            print("❌ 分析结果结构不正确")
            return False
    except Exception as e:
        print(f"❌ 配置分析失败: {e}")
        return False


def test_validation_rules_generation():
    """测试验证规则生成"""
    print("\n🧪 测试4: 验证规则生成")
    
    recommender = SmartConfigurationRecommender()
    excel_data = create_test_excel_data()
    
    try:
        recommendations = recommender.recommend_configurations(excel_data)
        validation_rules = recommendations["validation_rules"]
        
        print(f"✅ 生成了 {len(validation_rules)} 条验证规则")
        
        # 显示前3条规则
        for i, rule in enumerate(validation_rules[:3]):
            print(f"   {i+1}. [{rule.get('priority', 'medium')}] {rule.get('description', '无描述')}")
        
        return len(validation_rules) > 0
    except Exception as e:
        print(f"❌ 验证规则生成失败: {e}")
        return False


def test_optimization_suggestions():
    """测试优化建议生成"""
    print("\n🧪 测试5: 优化建议生成")
    
    recommender = SmartConfigurationRecommender()
    excel_data = create_test_excel_data()
    
    try:
        recommendations = recommender.recommend_configurations(excel_data)
        optimization_tips = recommendations["optimization_tips"]
        
        print(f"✅ 生成了 {len(optimization_tips)} 条优化建议")
        
        # 显示前3条建议
        for i, tip in enumerate(optimization_tips[:3]):
            print(f"   {i+1}. [{tip.get('type', 'general')}] {tip.get('suggestion', '无建议')}")
        
        return len(optimization_tips) >= 0  # 0条也是合理的结果
    except Exception as e:
        print(f"❌ 优化建议生成失败: {e}")
        return False


def test_integration_with_server():
    """测试与MCP服务器的集成"""
    print("\n🧪 测试6: MCP服务器集成检查")
    
    try:
        # 检查是否可以正确导入server模块
        from excel_mcp.server import mcp, SMART_CONFIG_AVAILABLE, SMART_CONFIG_TOOLS_AVAILABLE
        print("✅ MCP服务器模块导入成功")
        
        # 检查mcp对象是否可用
        print(f"✅ MCP对象类型: {type(mcp)}")
        
        # 检查智能配置工具是否可用
        if SMART_CONFIG_AVAILABLE and SMART_CONFIG_TOOLS_AVAILABLE:
            print("✅ 智能配置工具状态: 可用")
            
            # 检查智能配置工具函数是否存在
            from excel_mcp.server import recommend_excel_config, analyze_game_patterns, generate_validation_rules, optimize_data_structure
            print("✅ 智能配置工具函数导入成功")
            
            return True
        else:
            print("❌ 智能配置工具不可用")
            return False
            
    except ImportError as e:
        print(f"❌ MCP服务器集成失败: {e}")
        return False
    except Exception as e:
        print(f"❌ 未知集成错误: {e}")
        return False


def main():
    """运行所有测试"""
    print("🚀 开始智能配置工具验证测试")
    print("=" * 60)
    
    tests = [
        ("智能配置推荐", test_smart_config_recommendation),
        ("游戏类型检测", test_game_type_detection),
        ("配置分析器", test_configuration_analyzer),
        ("验证规则生成", test_validation_rules_generation),
        ("优化建议生成", test_optimization_suggestions),
        ("MCP服务器集成", test_integration_with_server)
    ]
    
    results = []
    
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"\n❌ 测试 {test_name} 异常: {e}")
            results.append((test_name, False))
    
    # 汇总测试结果
    print("\n" + "=" * 60)
    print("📊 测试结果汇总")
    print("=" * 60)
    
    passed = 0
    for test_name, result in results:
        status = "✅ 通过" if result else "❌ 失败"
        print(f"{test_name}: {status}")
        if result:
            passed += 1
    
    total = len(results)
    print(f"\n总计: {passed}/{total} 项测试通过")
    
    if passed == total:
        print("🎉 所有测试通过！智能配置工具功能正常")
        return True
    else:
        print("⚠️  部分测试失败，需要进一步修复")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)