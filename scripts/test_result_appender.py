#!/usr/bin/env python3
"""
测试结果自动追加器
加载测试结果 markdown 并追加到 docs/EDGE-CASE-TESTS.md
"""

import os
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple


class TestResultAppender:
    """测试结果追加类"""

    def __init__(self, target_file: str = "docs/EDGE-CASE-TESTS.md"):
        """初始化测试结果追加器

        Args:
            target_file: 目标 markdown 文件路径（默认为 docs/EDGE-CASE-TESTS.md）
        """
        self.target_file = target_file
        self.loaded_content: Optional[str] = None

    def load_markdown(self, source_path: str) -> str:
        """加载测试结果 markdown 文件

        Args:
            source_path: 源 markdown 文件路径

        Returns:
            加载的 markdown 内容字符串

        Raises:
            FileNotFoundError: 源文件不存在
            IOError: 读取文件失败
        """
        if not os.path.exists(source_path):
            raise FileNotFoundError(f"源文件不存在: {source_path}")

        try:
            with open(source_path, 'r', encoding='utf-8') as f:
                content = f.read()
                self.loaded_content = content
                print(f"成功加载测试结果: {source_path} ({len(content)} 字符)")
                return content
        except Exception as e:
            raise IOError(f"读取文件失败: {e}")

    def append_to_file(self, content: str, preview: bool = False) -> bool:
        """将内容追加到目标文件

        Args:
            content: 要追加的 markdown 内容
            preview: 是否只预览而不实际写入

        Returns:
            是否成功追加
        """
        # 预览模式
        if preview:
            print("\n=== 预览模式 - 不会实际写入文件 ===")
            print("将追加以下内容到:")
            print(f"  {self.target_file}")
            print("\n内容预览 (前500字符):")
            print(content[:500])
            if len(content) > 500:
                print(f"\n... (共 {len(content)} 字符)")
            print("=" * 60)
            return True

        # 验证内容格式
        if not self._validate_markdown_format(content):
            print("❌ Markdown 格式验证失败，拒绝追加")
            return False

        try:
            # 确保目标目录存在
            os.makedirs(os.path.dirname(self.target_file), exist_ok=True)

            # 追加内容
            with open(self.target_file, 'a', encoding='utf-8') as f:
                # 如果文件不为空，添加两个换行符分隔
                if os.path.exists(self.target_file) and os.path.getsize(self.target_file) > 0:
                    f.write("\n\n")
                f.write(content)

            print(f"✅ 成功追加内容到: {self.target_file}")
            return True
        except Exception as e:
            print(f"❌ 追加内容失败: {e}")
            return False

    def _validate_markdown_format(self, content: str) -> bool:
        """验证 markdown 格式是否符合要求

        Args:
            content: 要验证的 markdown 内容

        Returns:
            是否符合格式要求
        """
        errors = []

        # 检查是否有内容
        if not content or not content.strip():
            errors.append("内容为空")
            return False

        # 检查必需的标题格式（## YYYY-MM-DD 第X轮 或 ### 测试X:）
        has_round_header = re.search(r'## \d{4}-\d{2}-\d{2} 第\d+轮', content)
        has_test_header = re.search(r'### 测试\d+:', content)

        if not has_round_header and not has_test_header:
            errors.append("缺少必需的标题格式（## YYYY-MM-DD 第X轮 或 ### 测试X:）")

        # 检查是否有测试用例字段（操作步骤、预期结果、实际结果、是否通过）
        has_operation = re.search(r'- \*\*操作步骤\*\*:', content)
        has_expected = re.search(r'- \*\*预期结果\*\*:', content)
        has_actual = re.search(r'- \*\*实际结果\*\*:', content)
        has_status = re.search(r'- \*\*是否通过\*\*:', content)

        if has_test_header:
            # 如果有测试用例标题，检查必需字段
            if not (has_operation and has_expected and has_actual and has_status):
                missing = []
                if not has_operation:
                    missing.append("操作步骤")
                if not has_expected:
                    missing.append("预期结果")
                if not has_actual:
                    missing.append("实际结果")
                if not has_status:
                    missing.append("是否通过")
                errors.append(f"测试用例缺少必需字段: {', '.join(missing)}")

        # 如果有错误，打印并返回 False
        if errors:
            print("Markdown 格式验证失败:")
            for error in errors:
                print(f"  - {error}")
            return False

        print("✅ Markdown 格式验证通过")
        return True

    def get_latest_round_number(self) -> Optional[int]:
        """获取目标文件中的最新轮次编号

        Returns:
            最新轮次编号，如果文件不存在或无法解析则返回 None
        """
        if not os.path.exists(self.target_file):
            return None

        try:
            with open(self.target_file, 'r', encoding='utf-8') as f:
                content = f.read()

            # 查找所有轮次标题
            matches = re.findall(r'## \d{4}-\d{2}-\d{2} 第(\d+)轮', content)

            if matches:
                return max(map(int, matches))

        except Exception as e:
            print(f"解析轮次编号失败: {e}")

        return None

    def append_round_results(self, round_num: int, test_results: List[Dict],
                             preview: bool = False) -> bool:
        """追加一轮测试结果到文件

        Args:
            round_num: 轮次编号
            test_results: 测试结果列表，每个测试结果包含:
                         - title: 测试标题
                         - steps: 操作步骤
                         - expected: 预期结果
                         - actual: 实际结果
                         - status: 状态（PASS/FAIL/INFO）
                         - root_cause: 根因（可选）
                         - suggestion: 建议（可选）
            preview: 是否只预览而不实际写入

        Returns:
            是否成功追加
        """
        # 生成轮次标题
        markdown = f"## {datetime.now().strftime('%Y-%m-%d')} 第{round_num}轮\n\n"

        # 为每个测试结果生成测试用例
        for i, result in enumerate(test_results, start=1):
            # 计算测试编号（基于已有轮次）
            latest_round = self.get_latest_round_number()
            if latest_round:
                # 假设每轮最多20个测试，计算测试编号
                test_num = (latest_round * 20) + i
            else:
                test_num = i

            markdown += f"### 测试{test_num}: {result.get('title', '未知测试')}\n"
            markdown += f"- **操作步骤**: {result.get('steps', '')}\n"
            markdown += f"- **预期结果**: {result.get('expected', '')}\n"
            markdown += f"- **实际结果**: {result.get('actual', '')}\n"
            markdown += f"- **是否通过**: {result.get('status', 'UNKNOWN')}\n"

            # 可选字段
            if result.get('root_cause'):
                markdown += f"- **根因**: {result['root_cause']}\n"
            if result.get('suggestion'):
                markdown += f"- **建议**: {result['suggestion']}\n"

            markdown += "\n"

        # 添加统计信息
        total = len(test_results)
        passed = sum(1 for r in test_results if r.get('status') == 'PASS')
        failed = sum(1 for r in test_results if r.get('status') == 'FAIL')
        info = sum(1 for r in test_results if r.get('status') == 'INFO')

        markdown += f"### 第{round_num}轮统计\n"
        markdown += f"- **总计**: {total}个边缘案例\n"
        markdown += f"- **通过**: {passed}个\n"
        if failed > 0:
            markdown += f"- **失败**: {failed}个\n"
        if info > 0:
            markdown += f"- **信息**: {info}个\n"

        markdown += "\n"

        # 追加到文件
        return self.append_to_file(markdown, preview=preview)


def main():
    """主函数 - 演示使用流程"""
    import argparse

    parser = argparse.ArgumentParser(description='测试结果自动追加器')
    parser.add_argument('source', help='源 markdown 文件路径')
    parser.add_argument('--target', '-t', default='docs/EDGE-CASE-TESTS.md',
                       help='目标文件路径（默认: docs/EDGE-CASE-TESTS.md）')
    parser.add_argument('--preview', '-p', action='store_true',
                       help='预览模式，不实际写入文件')

    args = parser.parse_args()

    # 创建追加器实例
    appender = TestResultAppender(target_file=args.target)

    # 加载源文件
    try:
        content = appender.load_markdown(args.source)

        # 追加到目标文件
        success = appender.append_to_file(content, preview=args.preview)

        if success:
            print("\n✅ 操作完成")
            if args.preview:
                print("（预览模式，未实际写入）")
        else:
            print("\n❌ 操作失败")
    except Exception as e:
        print(f"\n❌ 错误: {e}")


if __name__ == "__main__":
    main()
