"""
智能配置推荐系统
基于游戏类型自动推荐配置结构、数值平衡性AI建议、配置优化预警系统
"""

import re
import json
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path
from collections import Counter, defaultdict
import pandas as pd


class GameTypeDetector:
    """游戏类型检测器"""
    
    GAME_PATTERNS = {
        'rpg': {
            'keywords': ['level', 'exp', 'hp', 'mp', 'skill', 'attribute', 'equipment', 'quest'],
            'sheet_patterns': ['skills', 'characters', 'items', 'quests', 'attributes']
        },
        'strategy': {
            'keywords': ['unit', 'building', 'resource', 'tech', 'civilization', 'army'],
            'sheet_patterns': ['units', 'buildings', 'resources', 'technologies']
        },
        'action': {
            'keywords': ['combo', 'animation', 'damage', 'health', 'power', 'weapon'],
            'sheet_patterns': ['characters', 'weapons', 'combos', 'animations']
        },
        'puzzle': {
            'keywords': ['level', 'score', 'achievement', 'star', 'objective'],
            'sheet_patterns': ['levels', 'scores', 'achievements', 'objectives']
        },
        'simulation': {
            'keywords': ['factory', 'production', 'resource', 'efficiency', 'capacity'],
            'sheet_patterns': ['buildings', 'production', 'resources', 'upgrades']
        }
    }
    
    def detect_game_type(self, excel_data: Dict[str, Any]) -> str:
        """检测游戏类型"""
        all_text = []
        
        # 收集所有工作表的文本数据
        for sheet_name, sheet_data in excel_data.items():
            if isinstance(sheet_data, dict):
                all_text.append(sheet_name)
                if 'data' in sheet_data:
                    for row in sheet_data['data']:
                        all_text.extend(str(cell) for cell in row if cell)
        
        text_content = ' '.join(all_text).lower()
        
        # 计算每种游戏类型的匹配度
        scores = {}
        for game_type, patterns in self.GAME_PATTERNS.items():
            score = 0
            for keyword in patterns['keywords']:
                score += text_content.count(keyword.lower())
            for sheet_pattern in patterns['sheet_patterns']:
                if sheet_pattern.lower() in text_content:
                    score += 2
            scores[game_type] = score
        
        # 返回得分最高的游戏类型
        return max(scores, key=scores.get) if scores else 'unknown'


class ConfigurationAnalyzer:
    """配置分析器"""
    
    def __init__(self):
        self.detector = GameTypeDetector()
    
    def analyze_excel_structure(self, excel_data: Dict[str, Any]) -> Dict[str, Any]:
        """分析Excel配置结构"""
        analysis = {
            'sheet_structure': {},
            'data_patterns': {},
            'validation_issues': [],
            'optimization_suggestions': []
        }
        
        for sheet_name, sheet_data in excel_data.items():
            if isinstance(sheet_data, dict) and 'data' in sheet_data:
                # 分析工作表结构
                rows = len(sheet_data['data'])
                cols = len(sheet_data['data'][0]) if rows > 0 else 0
                
                # 检测数据类型
                type_analysis = self._analyze_column_types(sheet_data['data'])
                
                analysis['sheet_structure'][sheet_name] = {
                    'rows': rows,
                    'cols': cols,
                    'types': type_analysis,
                    'headers': sheet_data['data'][0] if rows > 0 else []
                }
                
                # 识别数据模式
                patterns = self._identify_data_patterns(sheet_data['data'])
                analysis['data_patterns'][sheet_name] = patterns
                
                # 生成优化建议
                suggestions = self._generate_optimization_suggestions(
                    sheet_name, sheet_data['data'], patterns
                )
                analysis['optimization_suggestions'].extend(suggestions)
        
        return analysis
    
    def _analyze_column_types(self, data: List[List[Any]]) -> Dict[str, str]:
        """分析列数据类型"""
        if not data:
            return {}
        
        headers = data[0]
        types = {}
        
        for col_idx, header in enumerate(headers):
            col_data = [row[col_idx] for row in data[1:] if len(row) > col_idx]
            
            # 检测数据类型
            if all(isinstance(x, (int, float)) for x in col_data if x is not None):
                types[header] = 'numeric'
            elif all(isinstance(x, str) for x in col_data if x is not None):
                if any(x.lower() in ['true', 'false', 'yes', 'no'] for x in col_data if x):
                    types[header] = 'boolean'
                else:
                    types[header] = 'text'
            else:
                types[header] = 'mixed'
        
        return types
    
    def _identify_data_patterns(self, data: List[List[Any]]) -> Dict[str, Any]:
        """识别数据模式"""
        if not data:
            return {}
        
        headers = data[0]
        patterns = {}
        
        for col_idx, header in enumerate(headers):
            col_data = [row[col_idx] for row in data[1:] if len(row) > col_idx and row[col_idx] is not None]
            
            if not col_data:
                continue
            
            # 统计唯一值数量
            unique_count = len(set(col_data))
            total_count = len(col_data)
            
            # 识别常见模式
            patterns[header] = {
                'unique_values': unique_count,
                'total_values': total_count,
                'uniqueness_ratio': unique_count / total_count if total_count > 0 else 0,
                'sample_values': col_data[:5] if len(col_data) > 5 else col_data
            }
        
        return patterns
    
    def _generate_optimization_suggestions(self, sheet_name: str, data: List[List[Any]], patterns: Dict[str, Any]) -> List[Dict[str, Any]]:
        """生成优化建议"""
        suggestions = []
        
        if not data:
            return suggestions
        
        headers = data[0]
        
        for col_idx, header in enumerate(headers):
            if header in patterns:
                pattern = patterns[header]
                
                # 如果唯一值比例很低，建议转换为枚举类型
                if pattern['uniqueness_ratio'] < 0.1 and pattern['unique_values'] > 1:
                    suggestions.append({
                        'sheet': sheet_name,
                        'column': header,
                        'type': 'enum_optimization',
                        'suggestion': f'建议将"{header}"列转换为枚举类型，当前有{pattern["unique_values"]}个重复值',
                        'priority': 'medium'
                    })
                
                # 如果数值范围过大，建议添加范围验证
                if pattern.get('type') == 'numeric':
                    numeric_values = [float(x) for x in pattern['sample_values'] if isinstance(x, (int, float))]
                    if numeric_values:
                        min_val, max_val = min(numeric_values), max(numeric_values)
                        if max_val - min_val > 10000:
                            suggestions.append({
                                'sheet': sheet_name,
                                'column': header,
                                'type': 'validation_enhancement',
                                'suggestion': f'建议为"{header}"列添加数值范围验证({min_val:.2f} - {max_val:.2f})',
                                'priority': 'high'
                            })
        
        return suggestions


class SmartConfigurationRecommender:
    """智能配置推荐器"""
    
    def __init__(self):
        self.analyzer = ConfigurationAnalyzer()
    
    def recommend_configurations(self, excel_data: Dict[str, Any]) -> Dict[str, Any]:
        """智能推荐配置"""
        # 检测游戏类型
        game_type = self.analyzer.detector.detect_game_type(excel_data)
        
        # 分析配置结构
        analysis = self.analyzer.analyze_excel_structure(excel_data)
        
        # 生成配置建议
        recommendations = {
            'game_type': game_type,
            'analysis': analysis,
            'config_recommendations': self._generate_type_specific_recommendations(game_type, analysis),
            'validation_rules': self._generate_validation_rules(game_type, analysis),
            'optimization_tips': self._generate_optimization_tips(analysis)
        }
        
        return recommendations
    
    def _generate_type_specific_recommendations(self, game_type: str, analysis: Dict[str, Any]) -> List[Dict[str, Any]]:
        """生成类型特定的配置建议"""
        recommendations = []
        
        if game_type == 'rpg':
            recommendations.extend([
                {
                    'type': 'character_optimization',
                    'suggestion': '建议添加角色成长曲线表，记录等级与属性的关系',
                    'priority': 'high'
                },
                {
                    'type': 'balance_optimization', 
                    'suggestion': '建议添加技能平衡性检查工具，确保技能伤害比合理',
                    'priority': 'medium'
                }
            ])
        
        elif game_type == 'strategy':
            recommendations.extend([
                {
                    'type': 'resource_optimization',
                    'suggestion': '建议添加资源产出平衡性分析，避免经济失衡',
                    'priority': 'high'
                },
                {
                    'type': 'unit_optimization',
                    'suggestion': '建议添加单位克制关系表，完善战斗系统平衡性',
                    'priority': 'medium'
                }
            ])
        
        elif game_type == 'action':
            recommendations.extend([
                {
                    'type': 'combo_optimization',
                    'suggestion': '建议添加连招动作表，优化战斗流畅度',
                    'priority': 'medium'
                }
            ])
        
        return recommendations
    
    def _generate_validation_rules(self, game_type: str, analysis: Dict[str, Any]) -> List[Dict[str, Any]]:
        """生成验证规则"""
        rules = []
        
        # 基础验证规则
        for sheet_name, sheet_info in analysis.get('sheet_structure', {}).items():
            rules.append({
                'sheet': sheet_name,
                'rule': 'required_headers',
                'description': f'确保{sheet_name}表包含必要的基础字段',
                'priority': 'high'
            })
        
        # 游戏类型特定验证规则
        if game_type == 'rpg':
            rules.extend([
                {
                    'sheet': 'characters',
                    'rule': 'stat_validation',
                    'description': '角色属性必须为正数，且总和不超过最大值限制',
                    'priority': 'high'
                },
                {
                    'sheet': 'skills',
                    'rule': 'cost_validation', 
                    'description': '技能消耗必须合理，避免数值溢出',
                    'priority': 'medium'
                }
            ])
        
        return rules
    
    def _generate_optimization_tips(self, analysis: Dict[str, Any]) -> List[Dict[str, Any]]:
        """生成优化建议"""
        tips = []
        
        # 检查重复数据
        for sheet_name, patterns in analysis.get('data_patterns', {}).items():
            for col_name, pattern in patterns.items():
                if pattern.get('uniqueness_ratio', 0) < 0.1:
                    tips.append({
                        'type': 'data_compression',
                        'sheet': sheet_name,
                        'column': col_name,
                        'suggestion': f'{col_name}列重复率高，建议使用枚举或引用优化',
                        'priority': 'low'
                    })
        
        return tips