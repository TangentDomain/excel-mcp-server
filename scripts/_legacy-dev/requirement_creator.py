#!/usr/bin/env python3
"""
需求自动生成器
从 EDGE-CASE-TESTS.md 解析失败案例，自动生成 REQUIREMENTS.md 格式的需求
"""

import json
import os
import re
from datetime import datetime
from typing import Dict, List, Optional


class RequirementCreator:
    """需求自动生成类"""

    def __init__(self, test_results_path: str = "docs/EDGE-CASE-TESTS.md",
                 requirements_path: str = "REQUIREMENTS.md"):
        """初始化需求生成器

        Args:
            test_results_path: 测试结果 markdown 文件路径
            requirements_path: 需求文件路径
        """
        self.test_results_path = test_results_path
        self.requirements_path = requirements_path
        self.failed_cases: List[Dict] = []
        self.generated_requirements: List[Dict] = []

    def parse_test_results(self) -> List[Dict]:
        """解析测试结果 markdown，提取 FAIL/ERROR 案例

        Returns:
            失败案例列表，每个案例包含 title、description、root_cause 等字段
        """
        if not os.path.exists(self.test_results_path):
            print(f"警告: 测试结果文件不存在 {self.test_results_path}")
            return []

        try:
            with open(self.test_results_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 解析 markdown 中的测试用例
            failed_cases = []
            sections = re.split(r'## \d{4}-\d{2}-\d{2} 第\d+轮', content)

            for section in sections[1:]:  # 跳过标题前的空部分
                # 提取测试用例
                test_blocks = re.split(r'### 测试\d+:', section)

                for test_block in test_blocks[1:]:  # 跳过第一个空块
                    case = self._parse_single_test_case(test_block)
                    if case and case.get('status') in ['FAIL', 'ERROR']:
                        failed_cases.append(case)

            self.failed_cases = failed_cases
            print(f"从测试结果中提取 {len(failed_cases)} 个失败案例")
            return failed_cases

        except Exception as e:
            print(f"解析测试结果失败: {e}")
            return []

    def _parse_single_test_case(self, test_block: str) -> Optional[Dict]:
        """解析单个测试用例

        Args:
            test_block: 测试用例文本块

        Returns:
            测试用例字典
        """
        try:
            # 提取标题
            title_match = re.search(r'^(.+?)(?:\n|$)', test_block.strip())
            title = title_match.group(1).strip() if title_match else "未知测试"

            # 提取各个字段
            description_match = re.search(r'- \*\*操作步骤\*\*:\s*(.+?)(?:\n|$)', test_block)
            expected_match = re.search(r'- \*\*预期结果\*\*:\s*(.+?)(?:\n|$)', test_block)
            actual_match = re.search(r'- \*\*实际结果\*\*:\s*(.+?)(?:\n|$)', test_block)
            status_match = re.search(r'- \*\*是否通过\*\*:\s*(.+?)(?:\n|$)', test_block)
            root_cause_match = re.search(r'- \*\*根因\*\*:\s*(.+?)(?:\n|$)', test_block)
            suggestion_match = re.search(r'- \*\*建议\*\*:\s*(.+?)(?:\n|$)', test_block)

            case = {
                'title': title,
                'description': description_match.group(1).strip() if description_match else "",
                'expected': expected_match.group(1).strip() if expected_match else "",
                'actual': actual_match.group(1).strip() if actual_match else "",
                'status': status_match.group(1).strip() if status_match else "UNKNOWN",
                'root_cause': root_cause_match.group(1).strip() if root_cause_match else "",
                'suggestion': suggestion_match.group(1).strip() if suggestion_match else ""
            }

            return case

        except Exception as e:
            print(f"解析单个测试用例失败: {e}")
            return None

    def convert_to_requirement(self, case: Dict) -> Dict:
        """将失败案例转换为 REQUIREMENTS.md 格式

        Args:
            case: 失败案例字典

        Returns:
            需求字典
        """
        # 生成需求编号
        req_num = self._get_next_req_number()
        req_id = f"REQ-{req_num:03d}"

        # 确定优先级
        priority = self._determine_priority(case)

        # 构建需求描述
        description = self._build_description(case)

        requirement = {
            req_id: {
                "title": self._build_title(case),
                "type": "fix",
                "priority": priority,
                "status": "TODO",
                "source": f"边缘测试 - {case.get('title', '未知测试')}",
                "created": datetime.now().strftime("%Y-%m-%d"),
                "description": description,
                "notes": self._build_notes(case)
            }
        }

        return requirement

    def _get_next_req_number(self) -> int:
        """获取下一个需求编号

        Returns:
            下一个需求编号
        """
        try:
            if os.path.exists(self.requirements_path):
                with open(self.requirements_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                requirements = data.get('REQUIREMENTS', {})
                if requirements:
                    req_ids = list(requirements.keys())
                    # 提取编号并找到最大值
                    nums = []
                    for req_id in req_ids:
                        match = re.search(r'REQ-(\d+)', req_id)
                        if match:
                            nums.append(int(match.group(1)))

                    if nums:
                        return max(nums) + 1

        except Exception as e:
            print(f"获取需求编号失败: {e}")

        # 默认从 66 开始（避免与现有需求冲突）
        return 66

    def _determine_priority(self, case: Dict) -> str:
        """根据失败案例确定优先级

        Args:
            case: 失败案例字典

        Returns:
            优先级 (P0, P1, P2, P3)
        """
        root_cause = case.get('root_cause', '').lower()
        description = case.get('description', '').lower()

        # P0: 严重数据丢失或功能完全不可用
        if any(keyword in root_cause + description for keyword in
               ['静默替换', '静默截断', '丢失', '崩溃', '不可用', '返回空', '返回0']):
            return "P0"

        # P1: 功能错误或行为不符合预期
        if any(keyword in root_cause + description for keyword in
               ['错误', '失败', '不一致', '不准确', '缺失']):
            return "P1"

        # P2: 体验问题或信息不准确
        if any(keyword in root_cause + description for keyword in
               ['警告', '不明确', '不友好', '改进']):
            return "P2"

        # P3: 其他
        return "P3"

    def _build_title(self, case: Dict) -> str:
        """构建需求标题

        Args:
            case: 失败案例字典

        Returns:
            需求标题
        """
        title = case.get('title', '未知测试')

        # 根据失败原因添加前缀
        root_cause = case.get('root_cause', '')
        if '静默替换' in root_cause:
            title = f"修复{title}中的静默替换问题"
        elif '静默截断' in root_cause:
            title = f"修复{title}中的静默截断问题"
        elif '缺失' in root_cause:
            title = f"完善{title}功能"
        else:
            title = f"修复{title}"

        # 限制标题长度
        if len(title) > 50:
            title = title[:47] + "..."

        return title

    def _build_description(self, case: Dict) -> str:
        """构建需求描述

        Args:
            case: 失败案例字典

        Returns:
            需求描述
        """
        parts = []

        # 添加操作步骤
        if case.get('description'):
            parts.append(f"**测试场景**: {case['description']}")

        # 添加预期和实际结果
        if case.get('expected'):
            parts.append(f"**预期结果**: {case['expected']}")
        if case.get('actual'):
            parts.append(f"**实际结果**: {case['actual']}")

        # 添加根因
        if case.get('root_cause'):
            parts.append(f"**问题根因**: {case['root_cause']}")

        # 添加建议
        if case.get('suggestion'):
            parts.append(f"**修复建议**: {case['suggestion']}")

        return "\n".join(parts)

    def _build_notes(self, case: Dict) -> str:
        """构建需求备注

        Args:
            case: 失败案例字典

        Returns:
            需求备注
        """
        notes = []

        # 添加测试用例信息
        notes.append(f"来源测试: {case.get('title', '未知测试')}")

        # 添加状态
        notes.append(f"当前状态: {case.get('status', 'UNKNOWN')}")

        # 如果有根因，添加修复方向
        if case.get('root_cause'):
            notes.append(f"需要修复的文件/位置: {case['root_cause']}")

        return "\n".join(notes)

    def update_requirements(self, preview: bool = False) -> bool:
        """更新 REQUIREMENTS.md 文件

        Args:
            preview: 是否只预览而不实际写入

        Returns:
            是否成功
        """
        # 检查是否有失败案例
        if not self.failed_cases:
            print("没有失败案例可供生成需求")
            return False

        # 为每个失败案例生成需求
        new_requirements = {}
        for case in self.failed_cases:
            req = self.convert_to_requirement(case)
            new_requirements.update(req)

        # 检查是否已存在相同需求
        if os.path.exists(self.requirements_path):
            try:
                with open(self.requirements_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                existing_requirements = data.get('REQUIREMENTS', {})

                # 过滤已存在的需求
                filtered_requirements = {}
                for req_id, req_data in new_requirements.items():
                    title = req_data.get('title', '')
                    source = req_data.get('source', '')

                    # 检查是否已存在相同标题或来源的需求
                    exists = False
                    for existing_id, existing_data in existing_requirements.items():
                        existing_title = existing_data.get('title', '')
                        existing_source = existing_data.get('source', '')

                        if (title and existing_title and title == existing_title) or \
                           (source and existing_source and source == existing_source):
                            exists = True
                            print(f"需求 {req_id} 已存在（与 {existing_id} 重复），跳过创建")
                            break

                    if not exists:
                        filtered_requirements[req_id] = req_data

                new_requirements = filtered_requirements

            except Exception as e:
                print(f"检查现有需求失败: {e}，将添加所有新需求")

        if not new_requirements:
            print("没有新需求需要添加")
            return False

        # 预览模式
        if preview:
            print("\n=== 预览模式 - 不会实际写入文件 ===")
            print(f"将生成 {len(new_requirements)} 个新需求:")
            for req_id, req_data in new_requirements.items():
                print(f"\n{req_id}: {req_data['title']}")
                print(f"  优先级: {req_data['priority']}")
                print(f"  状态: {req_data['status']}")
                print(f"  来源: {req_data['source']}")
                print(f"  描述: {req_data['description'][:100]}...")
            print("=" * 60)
            return True

        # 实际更新文件
        try:
            # 读取现有需求
            if os.path.exists(self.requirements_path):
                with open(self.requirements_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            else:
                data = {"REQUIREMENTS": {}}

            # 添加新需求
            data['REQUIREMENTS'].update(new_requirements)

            # 写入文件
            with open(self.requirements_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            print(f"成功更新 {self.requirements_path}")
            print(f"添加 {len(new_requirements)} 个新需求")
            for req_id in new_requirements.keys():
                print(f"  - {req_id}")

            self.generated_requirements = new_requirements
            return True

        except Exception as e:
            print(f"更新需求文件失败: {e}")
            return False

    def generate_from_test_results(self, preview: bool = False) -> int:
        """从测试结果生成需求的便捷方法

        Args:
            preview: 是否只预览而不实际写入

        Returns:
            生成的需求数量
        """
        # 解析测试结果
        failed_cases = self.parse_test_results()
        if not failed_cases:
            return 0

        # 更新需求
        if self.update_requirements(preview=preview):
            return len(self.generated_requirements) if self.generated_requirements else len(failed_cases)

        return 0


def main():
    """主函数 - 演示使用流程"""
    import argparse

    parser = argparse.ArgumentParser(description='需求自动生成器')
    parser.add_argument('--test-results', default='docs/EDGE-CASE-TESTS.md',
                       help='测试结果 markdown 文件路径')
    parser.add_argument('--requirements', default='REQUIREMENTS.md',
                       help='需求文件路径')
    parser.add_argument('--preview', '-p', action='store_true',
                       help='预览模式，不实际写入文件')
    parser.add_argument('--parse-only', action='store_true',
                       help='只解析测试结果，不生成需求')

    args = parser.parse_args()

    # 创建需求生成器
    creator = RequirementCreator(
        test_results_path=args.test_results,
        requirements_path=args.requirements
    )

    if args.parse_only:
        # 只解析测试结果
        failed_cases = creator.parse_test_results()
        print(f"\n提取到 {len(failed_cases)} 个失败案例:")
        for i, case in enumerate(failed_cases, 1):
            print(f"\n{i}. {case['title']}")
            print(f"   状态: {case['status']}")
            print(f"   根因: {case['root_cause'][:50]}..." if case['root_cause'] else "   根因: 未知")
    else:
        # 生成需求
        count = creator.generate_from_test_results(preview=args.preview)
        if count > 0:
            print(f"\n✅ 成功生成 {count} 个需求")
        else:
            print("\n❌ 没有生成新需求")


if __name__ == "__main__":
    main()
