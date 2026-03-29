"""
MCP工具定义 - 智能配置推荐系统
提供基于游戏类型的Excel配置分析和优化建议
"""

from typing import Dict, List, Any, Optional
from server import ExcelMCPServer
from core.smart_config_recommender import SmartConfigurationRecommender
import json
import logging

# 智能配置推荐器实例
smart_recommender = SmartConfigurationRecommender()


def create_smart_config_tools(server: ExcelMCPServer):
    """创建智能配置推荐相关工具"""
    
    @server.tool
    def recommend_excel_config(
        file_path: str,
        game_type: Optional[str] = None,
        optimization_level: str = "balanced"
    ) -> str:
        """
        智能推荐Excel配置结构
        
        Args:
            file_path: Excel文件路径
            game_type: 游戏类型 (rpg/strategy/action/puzzle/simulation)，如不指定自动检测
            optimization_level: 优化级别 (basic/balanced/aggressive)
        
        Returns:
            智能配置推荐结果JSON字符串
        """
        try:
            # 读取Excel数据
            excel_data = server.api.read_excel_file(file_path)
            
            # 进行智能配置推荐
            recommendations = smart_recommender.recommend_configurations(excel_data)
            
            # 根据优化级别调整建议详细程度
            if optimization_level == "basic":
                # 只保留核心建议
                result = {
                    "game_type": recommendations["game_type"],
                    "core_recommendations": recommendations["config_recommendations"][:3],
                    "critical_validation_rules": [r for r in recommendations["validation_rules"] if r["priority"] == "high"]
                }
            elif optimization_level == "aggressive":
                # 添加详细的优化建议
                result = {
                    "game_type": recommendations["game_type"],
                    "full_analysis": recommendations["analysis"],
                    "all_recommendations": recommendations["config_recommendations"],
                    "all_validation_rules": recommendations["validation_rules"],
                    "optimization_tips": recommendations["optimization_tips"],
                    "summary": self._generate_summary(recommendations)
                }
            else:
                # balanced 默认模式
                result = {
                    "game_type": recommendations["game_type"],
                    "analysis_summary": self._generate_summary(recommendations),
                    "key_recommendations": recommendations["config_recommendations"][:5],
                    "important_validation_rules": recommendations["validation_rules"][:10]
                }
            
            return json.dumps(result, ensure_ascii=False, indent=2)
            
        except Exception as e:
            logging.error(f"配置推荐失败: {str(e)}")
            return json.dumps({
                "error": "配置推荐失败",
                "message": str(e)
            }, ensure_ascii=False)
    
    @server.tool  
    def analyze_game_patterns(
        file_path: str,
        target_sheet: Optional[str] = None
    ) -> str:
        """
        分析游戏模式和数据结构
        
        Args:
            file_path: Excel文件路径
            target_sheet: 指定分析特定工作表，如不分析所有工作表
        
        Returns:
            游戏模式分析结果JSON字符串
        """
        try:
            # 读取Excel数据
            excel_data = server.api.read_excel_file(file_path)
            
            # 创建分析器
            from core.smart_config_recommender import ConfigurationAnalyzer
            analyzer = ConfigurationAnalyzer()
            
            # 分析数据结构
            analysis = analyzer.analyze_excel_structure(excel_data)
            
            # 检测游戏类型
            game_type = analyzer.detector.detect_game_type(excel_data)
            
            result = {
                "detected_game_type": game_type,
                "analysis_scope": target_sheet if target_sheet else "all_sheets",
                "data_patterns": analysis.get("data_patterns", {}),
                "structure_analysis": analysis.get("sheet_structure", {}),
                "optimization_suggestions": analysis.get("optimization_suggestions", []),
                "data_quality_score": self._calculate_data_quality_score(analysis)
            }
            
            return json.dumps(result, ensure_ascii=False, indent=2)
            
        except Exception as e:
            logging.error(f"游戏模式分析失败: {str(e)}")
            return json.dumps({
                "error": "游戏模式分析失败", 
                "message": str(e)
            }, ensure_ascii=False)
    
    @server.tool
    def generate_validation_rules(
        file_path: str,
        rule_categories: Optional[List[str]] = None
    ) -> str:
        """
        基于游戏配置生成验证规则
        
        Args:
            file_path: Excel文件路径
            rule_categories: 规则类别列表，如不生成全部类别
        
        Returns:
            验证规则JSON字符串
        """
        try:
            # 读取Excel数据
            excel_data = server.api.read_excel_file(file_path)
            
            # 获取推荐
            recommendations = smart_recommender.recommend_configurations(excel_data)
            
            # 筛选规则类别
            if rule_categories:
                filtered_rules = []
                for rule in recommendations["validation_rules"]:
                    if rule["sheet"] in rule_categories:
                        filtered_rules.append(rule)
                rules = filtered_rules
            else:
                rules = recommendations["validation_rules"]
            
            result = {
                "validation_rules": rules,
                "rule_categories": list(set(rule["sheet"] for rule in rules)),
                "priority_breakdown": self._categorize_by_priority(rules),
                "rule_summary": f"生成了{len(rules)}个验证规则"
            }
            
            return json.dumps(result, ensure_ascii=False, indent=2)
            
        except Exception as e:
            logging.error(f"验证规则生成失败: {str(e)}")
            return json.dumps({
                "error": "验证规则生成失败",
                "message": str(e)
            }, ensure_ascii=False)
    
    @server.tool
    def optimize_data_structure(
        file_path: str,
        optimization_type: str = "compression"
    ) -> str:
        """
        优化Excel数据结构
        
        Args:
            file_path: Excel文件路径
            optimization_type: 优化类型 (compression/restructuring/indexing)
        
        Returns:
            优化建议JSON字符串
        """
        try:
            # 读取Excel数据
            excel_data = server.api.read_excel_file(file_path)
            
            # 分析数据结构
            from core.smart_config_recommender import ConfigurationAnalyzer
            analyzer = ConfigurationAnalyzer()
            analysis = analyzer.analyze_excel_structure(excel_data)
            
            # 生成优化建议
            optimization_suggestions = []
            
            if optimization_type == "compression":
                # 数据压缩优化
                for sheet_name, patterns in analysis.get("data_patterns", {}).items():
                    for col_name, pattern in patterns.items():
                        if pattern.get("uniqueness_ratio", 0) < 0.1:
                            optimization_suggestions.append({
                                "type": "enum_conversion",
                                "sheet": sheet_name,
                                "column": col_name,
                                "description": f"建议将{col_name}转换为枚举类型，节省存储空间",
                                "estimated_savings": f"{(1 - pattern.get('uniqueness_ratio', 0)) * 100:.1f}%"
                            })
            
            elif optimization_type == "restructuring":
                # 结构重构优化
                for sheet_name, sheet_info in analysis.get("sheet_structure", {}).items():
                    if sheet_info["rows"] > 1000:
                        optimization_suggestions.append({
                            "type": "normalization",
                            "sheet": sheet_name,
                            "description": f"{sheet_name}表数据量较大，建议考虑表结构规范化",
                            "current_rows": sheet_info["rows"],
                            "recommendation": "拆分为多个关联表"
                        })
            
            elif optimization_type == "indexing":
                # 索引优化
                for sheet_name, sheet_info in analysis.get("sheet_structure", {}).items():
                    headers = sheet_info.get("headers", [])
                    for i, header in enumerate(headers):
                        if any(keyword in header.lower() for keyword in ["id", "name", "key", "code"]):
                            optimization_suggestions.append({
                                "type": "index_recommendation",
                                "sheet": sheet_name,
                                "column": header,
                                "description": f"建议为{sheet_name}表的{header}列添加索引，提升查询性能",
                                "query_type": "primary_key" if "id" in header.lower() else "search_key"
                            })
            
            result = {
                "optimization_type": optimization_type,
                "original_structure": analysis,
                "optimization_suggestions": optimization_suggestions,
                "expected_improvements": self._estimate_improvements(optimization_suggestions)
            }
            
            return json.dumps(result, ensure_ascii=False, indent=2)
            
        except Exception as e:
            logging.error(f"数据结构优化失败: {str(e)}")
            return json.dumps({
                "error": "数据结构优化失败",
                "message": str(e)
            }, ensure_ascii=False)
    
    # 辅助方法
    def _generate_summary(self, recommendations: Dict[str, Any]) -> str:
        """生成推荐摘要"""
        game_type = recommendations["game_type"]
        key_recs = recommendations["config_recommendations"][:3]
        
        summary = f"检测到游戏类型: {game_type}\\n"
        summary += f"核心推荐: {len(key_recs)}条\\n"
        summary += "主要建议: "
        summary += "; ".join([rec["suggestion"] for rec in key_recs])
        
        return summary
    
    def _calculate_data_quality_score(self, analysis: Dict[str, Any]) -> float:
        """计算数据质量评分"""
        score = 100.0
        
        # 根据优化建议扣分
        suggestions = analysis.get("optimization_suggestions", [])
        score -= len(suggestions) * 5  # 每个建议扣5分
        
        # 确保评分在0-100之间
        return max(0, min(100, score))
    
    def _categorize_by_priority(self, rules: List[Dict[str, Any]]) -> Dict[str, int]:
        """按优先级分类规则"""
        priority_count = {"high": 0, "medium": 0, "low": 0}
        for rule in rules:
            priority = rule.get("priority", "medium")
            priority_count[priority] += 1
        return priority_count
    
    def _estimate_improvements(self, suggestions: List[Dict[str, Any]]) -> Dict[str, Any]:
        """预估优化效果"""
        improvements = {
            "performance_gain": "预估查询速度提升20-30%",
            "storage_reduction": "预估存储节省10-25%",
            "maintainability": "预估维护难度降低40%"
        }
        
        if len(suggestions) > 5:
            improvements["significant_improvement"] = "大规模优化，效果显著"
        elif len(suggestions) > 2:
            improvements["moderate_improvement"] = "中等规模优化，效果明显"
        else:
            improvements["minor_improvement"] = "小幅优化，略有改善"
        
        return improvements