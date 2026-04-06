#!/usr/bin/env python3
"""
边缘案例自动发现脚本
从 Stack Overflow 和 GitHub Issues 搜索真实用户遇到的 Excel 奇怪问题
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from typing import Dict, List, Optional
from urllib.parse import quote, urlencode


class EdgeCaseCollector:
    """边缘案例收集器"""

    def __init__(self, output_path: str = "docs/edge_cases.json"):
        """初始化收集器
        
        Args:
            output_path: 输出 JSON 文件路径
        """
        self.output_path = output_path
        self.existing_cases = self._load_existing_cases()
        self.search_keywords = [
            "Excel strange bug",
            "Excel error unexpected",
            "Excel weird behavior",
            "Excel edge case",
            "Excel unexpected behavior",
            "Excel formula bug",
            "Excel data corruption",
            "Excel formatting issue",
            "Excel calculation error"
        ]

    def _load_existing_cases(self) -> Dict[str, Dict]:
        """加载已存在的边缘案例
        
        Returns:
            已有案例的字典，以标题为键
        """
        if os.path.exists(self.output_path):
            try:
                with open(self.output_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return {case['title']: case for case in data.get('edge_cases', [])}
            except Exception as e:
                print(f"加载现有案例失败: {e}")
        return {}

    def _save_cases(self, cases: List[Dict]) -> None:
        """保存边缘案例到 JSON 文件
        
        Args:
            cases: 边缘案例列表
        """
        output_data = {
            'last_updated': datetime.now().isoformat(),
            'total_cases': len(cases),
            'edge_cases': cases
        }

        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        with open(self.output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)

    def _is_duplicate(self, title: str) -> bool:
        """检查案例是否重复
        
        Args:
            title: 案例标题
            
        Returns:
            如果重复返回 True
        """
        return title.lower() in [k.lower() for k in self.existing_cases.keys()]

    def _calculate_priority(self, views: int, answers: int, score: int) -> str:
        """根据流行度和严重程度计算优先级
        
        Args:
            views: 浏览次数
            answers: 回答数量
            score: 评分
            
        Returns:
            优先级等级（high/medium/low）
        """
        if views > 5000 or answers > 5 or score > 10:
            return "high"
        elif views > 1000 or answers > 2 or score > 5:
            return "medium"
        else:
            return "low"

    def _parse_stackoverflow_item(self, item: Dict) -> Optional[Dict]:
        """解析 Stack Overflow 问答项
        
        Args:
            item: Stack Overflow API 返回的项
            
        Returns:
            解析后的边缘案例字典，如果无效则返回 None
        """
        if item.get('answer_count', 0) == 0:
            return None

        title = item.get('title', '')
        if not title or self._is_duplicate(title):
            return None

        tags = item.get('tags', [])
        if not any('excel' in tag.lower() for tag in tags):
            return None

        priority = self._calculate_priority(
            item.get('view_count', 0),
            item.get('answer_count', 0),
            item.get('score', 0)
        )

        case = {
            'title': title,
            'description': self._clean_text(item.get('body', '')),
            'steps': self._extract_steps(item.get('body', '')),
            'expected': '',
            'actual': '',
            'source': 'stackoverflow',
            'source_url': f"https://stackoverflow.com/questions/{item.get('question_id', '')}",
            'views': item.get('view_count', 0),
            'answers': item.get('answer_count', 0),
            'score': item.get('score', 0),
            'tags': tags,
            'priority': priority,
            'discovered_at': datetime.now().isoformat()
        }

        return case

    def _clean_text(self, text: str) -> str:
        """清理 HTML 文本
        
        Args:
            text: 原始文本（可能包含 HTML）
            
        Returns:
            清理后的纯文本
        """
        text = re.sub(r'<[^>]+>', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()[:500]

    def _extract_steps(self, text: str) -> List[str]:
        """从文本中提取操作步骤
        
        Args:
            text: 问题描述文本
            
        Returns:
            操作步骤列表
        """
        steps = []
        patterns = [
            r'steps?:\s*([^.]+)',
            r'to reproduce:\s*([^.]+)',
            r'first,\s*([^.]+)',
            r'when i\s*([^.]+)',
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                step = match.group(1).strip()
                if len(step) > 10:
                    steps.append(step[:200])
        
        return steps[:5]

    def search_stackoverflow(self, max_results: int = 20) -> List[Dict]:
        """搜索 Stack Overflow 获取 Excel 相关问题
        
        Args:
            max_results: 最大返回结果数
            
        Returns:
            边缘案例列表
        """
        cases = []
        
        for keyword in self.search_keywords:
            try:
                query = f"{keyword} [excel]"
                url = f"https://api.stackexchange.com/2.3/search/advanced"
                
                params = {
                    'order': 'desc',
                    'sort': 'votes',
                    'q': query,
                    'accepted': 'True',
                    'answers': '1',
                    'pagesize': min(max_results, 100),
                    'site': 'stackoverflow',
                    'filter': 'withbody'
                }
                
                import requests
                response = requests.get(url, params=params, timeout=30)
                response.raise_for_status()
                data = response.json()
                
                for item in data.get('items', []):
                    case = self._parse_stackoverflow_item(item)
                    if case:
                        cases.append(case)
                        self.existing_cases[case['title']] = case
                        
            except Exception as e:
                print(f"搜索 Stack Overflow 失败 ({keyword}): {e}")
                continue

        return cases

    def search_github_issues(self, repos: List[str] = None, max_per_repo: int = 5) -> List[Dict]:
        """搜索 GitHub Issues 获取 Excel 相关问题
        
        Args:
            repos: 要搜索的仓库列表
            max_per_repo: 每个仓库返回的最大问题数
            
        Returns:
            边缘案例列表
        """
        if repos is None:
            repos = [
                "microsoft/vscode",
                "SheetJS/sheetjs",
                "closedxml/closedxml",
                "microsoft/Excel-JS"
            ]
        
        cases = []
        
        for repo in repos:
            try:
                url = f"https://api.github.com/repos/{repo}/issues"
                params = {
                    'state': 'closed',
                    'sort': 'comments',
                    'direction': 'desc',
                    'per_page': max_per_repo,
                    'labels': 'bug'
                }
                
                import requests
                response = requests.get(url, params=params, timeout=30)
                response.raise_for_status()
                data = response.json()
                
                for issue in data:
                    if 'pull_request' in issue:
                        continue
                    
                    title = issue.get('title', '')
                    if not title or self._is_duplicate(title):
                        continue
                    
                    priority = self._calculate_priority(
                        issue.get('comments', 0) * 10,
                        issue.get('comments', 0),
                        issue.get('reactions', {}).get('+1', 0)
                    )
                    
                    case = {
                        'title': title,
                        'description': self._clean_text(issue.get('body', '')),
                        'steps': self._extract_steps(issue.get('body', '')),
                        'expected': '',
                        'actual': '',
                        'source': 'github',
                        'source_url': issue.get('html_url', ''),
                        'views': issue.get('comments', 0) * 10,
                        'answers': issue.get('comments', 0),
                        'score': issue.get('reactions', {}).get('+1', 0),
                        'tags': [repo],
                        'priority': priority,
                        'discovered_at': datetime.now().isoformat()
                    }
                    
                    cases.append(case)
                    self.existing_cases[case['title']] = case
                    
            except Exception as e:
                print(f"搜索 GitHub Issues 失败 ({repo}): {e}")
                continue

        return cases

    def discover_edge_cases(self) -> List[Dict]:
        """发现并收集边缘案例
        
        Returns:
            所有发现的边缘案例列表
        """
        all_cases = []
        
        print("正在搜索 Stack Overflow...")
        so_cases = self.search_stackoverflow()
        all_cases.extend(so_cases)
        print(f"从 Stack Overflow 发现 {len(so_cases)} 个案例")
        
        print("正在搜索 GitHub Issues...")
        gh_cases = self.search_github_issues()
        all_cases.extend(gh_cases)
        print(f"从 GitHub 发现 {len(gh_cases)} 个案例")
        
        return all_cases

    def get_recent_high_priority_cases(self, limit: int = 10) -> List[Dict]:
        """获取最近发现的高优先级边缘案例
        
        Args:
            limit: 返回的最大案例数
            
        Returns:
            高优先级边缘案例列表，按发现时间倒序排列
        """
        cases = list(self.existing_cases.values())
        high_priority = [c for c in cases if c['priority'] == 'high']
        
        # 按发现时间倒序排序
        high_priority.sort(
            key=lambda x: x.get('discovered_at', ''),
            reverse=True
        )
        
        return high_priority[:limit]

    def convert_to_test_case(self, case: Dict) -> str:
        """将边缘案例转换为测试用例格式
        
        Args:
            case: 边缘案例字典
            
        Returns:
            格式化的测试用例字符串
        """
        test_case = f"""# 测试用例: {case['title']}

## 优先级: {case['priority']}

## 描述
{case['description']}

## 来源
{case['source']}: {case['source_url']}

## 操作步骤
"""
        if case.get('steps'):
            for i, step in enumerate(case['steps'], 1):
                test_case += f"{i}. {step}\n"
        else:
            test_case += "暂无详细步骤\n"
        
        test_case += f"""
## 预期结果
{case.get('expected', '待补充')}

## 实际结果
{case.get('actual', '待补充')}

## 元数据
- 浏览量: {case['views']}
- 回答数: {case['answers']}
- 评分: {case['score']}
- 发现时间: {case['discovered_at']}
"""
        return test_case

    def run(self) -> None:
        """运行边缘案例发现流程"""
        print("=" * 50)
        print("边缘案例自动发现工具")
        print("=" * 50)
        
        new_cases = self.discover_edge_cases()
        
        if not new_cases:
            print("未发现新的边缘案例")
            return
        
        # 加载现有案例
        existing_list = list(self.existing_cases.values())
        
        # 合并新旧案例
        all_cases = existing_list + new_cases
        
        # 去重（以标题为键）
        unique_cases = {case['title']: case for case in all_cases}
        final_cases = list(unique_cases.values())
        
        # 按优先级排序
        priority_order = {'high': 0, 'medium': 1, 'low': 2}
        final_cases.sort(key=lambda x: (priority_order.get(x['priority'], 3), -x['score']))
        
        # 保存
        self._save_cases(final_cases)
        
        print("=" * 50)
        print(f"总计: {len(final_cases)} 个边缘案例")
        print(f"新增: {len(new_cases)} 个案例")
        print(f"高优先级: {len([c for c in final_cases if c['priority'] == 'high'])}")
        print(f"中优先级: {len([c for c in final_cases if c['priority'] == 'medium'])}")
        print(f"低优先级: {len([c for c in final_cases if c['priority'] == 'low'])}")
        print(f"输出文件: {self.output_path}")
        print("=" * 50)


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='边缘案例自动发现工具',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        '--recent', '-r',
        action='store_true',
        help='只返回最近发现的高优先级案例，不执行完整搜索'
    )
    parser.add_argument(
        '--limit', '-l',
        type=int,
        default=10,
        help='返回的案例数量限制（默认: 10）'
    )
    parser.add_argument(
        '--format-test-cases', '-f',
        action='store_true',
        help='将案例转换为测试用例格式输出'
    )
    parser.add_argument(
        '--output', '-o',
        type=str,
        default='docs/edge_cases.json',
        help='输出文件路径（默认: docs/edge_cases.json）'
    )
    
    args = parser.parse_args()
    
    collector = EdgeCaseCollector(output_path=args.output)
    
    if args.recent:
        cases = collector.get_recent_high_priority_cases(limit=args.limit)
        print(f"\n最近发现的高优先级边缘案例 ({len(cases)} 个):\n")
        print("=" * 50)
        
        for i, case in enumerate(cases, 1):
            if args.format_test_cases:
                print(collector.convert_to_test_case(case))
                print("\n" + "-" * 50 + "\n")
            else:
                print(f"{i}. [{case['priority'].upper()}] {case['title']}")
                print(f"   来源: {case['source']}")
                print(f"   URL: {case['source_url']}")
                print(f"   发现时间: {case['discovered_at']}")
                print(f"   评分: {case['score']} | 浏览: {case['views']} | 回答: {case['answers']}")
                print()
    else:
        collector.run()


if __name__ == "__main__":
    main()
