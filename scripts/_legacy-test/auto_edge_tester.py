#!/usr/bin/env python3
"""
边缘测试自动化脚本
从 edge_case_discovery.py 的输出加载案例，生成测试脚本并执行
"""

import json
import os
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional


class EdgeCaseAutomation:
    """边缘案例测试自动化类"""

    def __init__(self, discovery_output: str = "docs/edge_cases.json",
                 test_output_dir: str = "scripts",
                 markdown_output: str = "docs/EDGE-CASE-TESTS.md"):
        """初始化自动化类

        Args:
            discovery_output: edge_case_discovery.py 输出的 JSON 文件路径
            test_output_dir: 测试脚本输出目录
            markdown_output: Markdown 报告输出路径
        """
        self.discovery_output = discovery_output
        self.test_output_dir = test_output_dir
        self.markdown_output = markdown_output
        self.cases: List[Dict] = []

    def load_from_discovery(self, json_path: Optional[str] = None) -> List[Dict]:
        """从 edge_case_discovery.py 的输出加载案例

        Args:
            json_path: JSON 文件路径，如果为 None 则使用初始化时的路径

        Returns:
            加载的边缘案例列表
        """
        path = json_path or self.discovery_output

        if not os.path.exists(path):
            print(f"警告: 发现文件不存在 {path}")
            return []

        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.cases = data.get('edge_cases', [])
                print(f"成功加载 {len(self.cases)} 个边缘案例")
                return self.cases
        except Exception as e:
            print(f"加载边缘案例失败: {e}")
            return []

    def _get_next_test_number(self) -> int:
        """获取下一个测试编号

        Returns:
            下一个测试编号
        """
        max_num = 0
        test_files = list(Path(self.test_output_dir).glob("edge_case_tests_round*.py"))

        for test_file in test_files:
            try:
                with open(test_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                    import re
                    matches = list(map(int, re.findall(r'self\.test\((\d+)', content)))
                    if matches:
                        max_num = max(max_num, max(matches))
            except Exception:
                continue

        return max_num + 1

    def _get_next_round_number(self) -> int:
        """获取下一个轮次编号

        Returns:
            下一个轮次编号
        """
        max_round = 0
        test_files = list(Path(self.test_output_dir).glob("edge_case_tests_round*.py"))

        for test_file in test_files:
            import re
            match = re.search(r'round(\d+)', test_file.name)
            if match:
                max_round = max(max_round, int(match.group(1)))

        return max_round + 1

    def _generate_test_case_method(self, case: Dict, test_num: int) -> str:
        """生成测试用例方法

        Args:
            case: 边缘案例字典
            test_num: 测试编号

        Returns:
            生成的测试方法代码字符串
        """
        title = case.get('title', '未知测试')
        description = case.get('description', '')
        steps = case.get('steps', [])
        priority = case.get('priority', 'low')
        source = case.get('source', 'unknown')
        source_url = case.get('source_url', '')

        method_code = f'''    def _test_case_{test_num}(self):
        """测试用例 T{test_num}: {title}

        来源: {source}
        优先级: {priority}
        URL: {source_url}
        """
        # TODO: 根据描述和步骤实现具体测试逻辑
        # {title}
        # {description[:200]}
'''

        if steps:
            method_code += "        # 操作步骤:\n"
            for i, step in enumerate(steps[:3], 1):
                method_code += f"        # {i}. {step}\n"

        method_code += "        return {'success': False, 'message': '待实现测试'}\n\n"
        return method_code

    def _generate_test_function_call(self, case: Dict, test_num: int) -> str:
        """生成测试函数调用

        Args:
            case: 边缘案例字典
            test_num: 测试编号

        Returns:
            生成的测试函数调用代码字符串
        """
        title = case.get('title', '未知测试')
        priority = case.get('priority', 'low')

        # 根据优先级决定是否预期通过
        expected_pass = priority == 'high'

        test_code = f'''        # T{test_num}: {title}
        self.test({test_num}, "{title[:50]}", lambda: self._test_case_{test_num}(), expected_pass={expected_pass.lower()})
'''
        return test_code

    def generate_test_round(self, round_num: Optional[int] = None) -> str:
        """生成当前轮次的测试脚本

        Args:
            round_num: 轮次编号，如果为 None 则自动生成

        Returns:
            生成的测试脚本文件路径
        """
        if not self.cases:
            print("没有边缘案例可供生成测试脚本")
            return ""

        # 确定轮次编号
        if round_num is None:
            round_num = self._get_next_round_number()

        # 获取测试编号范围
        start_num = self._get_next_test_number()
        end_num = start_num + len(self.cases) - 1

        # 生成测试脚本
        script_content = f'''#!/usr/bin/env python3
"""
边缘案例测试脚本 - 第{round_num}轮 - T{start_num}-T{end_num}
自动生成时间: {datetime.now().isoformat()}
"""

import subprocess
import json
import tempfile
import os
from openpyxl import Workbook


class EdgeCaseTester:
    """边缘案例测试器"""

    def __init__(self):
        """初始化测试器"""
        self.results = []

    def create_test_file(self, data=None, sheet_names=None):
        """用 openpyxl 创建测试文件

        Args:
            data: 测试数据（二维列表）
            sheet_names: 工作表名称列表

        Returns:
            临时测试文件路径
        """
        filepath = tempfile.mktemp(suffix=".xlsx")
        wb = Workbook()

        if sheet_names and len(sheet_names) > 1:
            wb.remove(wb.active)
            for name in sheet_names:
                wb.create_sheet(name)
        else:
            ws = wb.active
            if sheet_names:
                ws.title = sheet_names[0]
            if data:
                for row in data:
                    ws.append(row)

        wb.save(filepath)
        return filepath

    def call_mcp(self, method, params):
        """调用 MCP 工具

        Args:
            method: MCP 方法名
            params: 参数字典

        Returns:
            工具调用结果
        """
        proc = subprocess.Popen(
            ["uvx", "--from", ".", "excel-mcp-server-fastmcp"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        init = json.dumps({{
            "jsonrpc": "2.0", "id": 0,
            "method": "initialize",
            "params": {{
                "protocolVersion": "2024-11-05",
                "capabilities": {{}},
                "clientInfo": {{"name": "test", "version": "1.0"}}
            }}
        }})
        proc.stdin.write(init + "\\n")
        proc.stdin.flush()

        request = json.dumps({{
            "jsonrpc": "2.0", "id": 1,
            "method": "tools/call",
            "params": {{"name": method, "arguments": params}}
        }})

        try:
            proc.stdin.write(request + "\\n")
            proc.stdin.flush()

            # 读取多行响应直到找到工具调用的结果
            while True:
                response = proc.stdout.readline()
                if not response:
                    break

                try:
                    result = json.loads(response)
                    # 检查是否是工具调用的响应（id=1）
                    if result.get("id") == 1:
                        # 找到工具调用响应
                        if "error" in result:
                            proc.terminate()
                            return {{"success": False, "message": result["error"].get("message", str(result["error"]))}}

                        r = result.get("result", {{}})
                        content = r.get("content", [])
                        if content and isinstance(content, list):
                            text = content[0].get("text", "")
                            try:
                                proc.terminate()
                                return json.loads(text)
                            except json.JSONDecodeError:
                                proc.terminate()
                                return {{"success": bool(text), "data": text}}
                        proc.terminate()
                        return r
                except json.JSONDecodeError:
                    continue

            proc.terminate()
            return {{"success": False, "message": "No valid response received"}}
        except Exception as e:
            proc.terminate()
            return {{"success": False, "message": str(e)}}

    def test(self, num, description, operation, expected_pass=True):
        """执行单个测试

        Args:
            num: 测试编号
            description: 测试描述
            operation: 测试操作函数
            expected_pass: 预期是否通过
        """
        print(f"测试T{{num}}: {{description}}")
        try:
            result = operation()
            success = result.get("success", False)
            has_data = result.get("data") is not None

            if expected_pass:
                if success or has_data:
                    status = "PASS"
                else:
                    status = "FAIL"
            else:
                if not success:
                    status = "PASS"
                else:
                    status = "FAIL"

            self.results.append({{
                "num": num, "desc": description,
                "status": status, "result": str(result)[:200]
            }})
            print(f"  结果: {{status}} - {{self.results[-1]['result'][:100]}}")
        except Exception as e:
            self.results.append({{
                "num": num, "desc": description,
                "status": "ERROR", "result": str(e)
            }})
            print(f"  结果: ERROR - {{e}}")

'''

        # 生成所有测试用例方法
        for i, case in enumerate(self.cases):
            test_num = start_num + i
            script_content += self._generate_test_case_method(case, test_num)

        # 生成运行所有测试的方法
        script_content += f'''
    def run_tests(self):
        """运行所有测试"""
        print("=" * 60)
        print(f"开始执行第{round_num}轮边缘案例测试")
        print("=" * 60)

'''

        # 生成测试调用
        for i, case in enumerate(self.cases):
            test_num = start_num + i
            script_content += self._generate_test_function_call(case, test_num)

        # 生成统计和输出
        script_content += '''
        self.print_summary()

    def print_summary(self):
        """打印测试摘要"""
        print("\\n" + "=" * 60)
        print("测试摘要")
        print("=" * 60)

        passed = sum(1 for r in self.results if r["status"] == "PASS")
        failed = sum(1 for r in self.results if r["status"] == "FAIL")
        errors = sum(1 for r in self.results if r["status"] == "ERROR")

        print(f"总计: {len(self.results)} 个测试")
        print(f"通过: {passed} 个")
        print(f"失败: {failed} 个")
        print(f"错误: {errors} 个")

        if failed > 0:
            print("\\n失败的测试:")
            for r in self.results:
                if r["status"] == "FAIL":
                    print(f"  T{r['num']}: {r['desc']}")

        if errors > 0:
            print("\\n错误的测试:")
            for r in self.results:
                if r["status"] == "ERROR":
                    print(f"  T{r['num']}: {r['desc']}")

        return {{
            "total": len(self.results),
            "pass": passed,
            "fail": failed,
            "error": errors
        }}


if __name__ == "__main__":
    tester = EdgeCaseTester()
    tester.run_tests()
    stats = tester.print_summary()
'''

        # 保存测试脚本
        os.makedirs(self.test_output_dir, exist_ok=True)
        script_path = os.path.join(self.test_output_dir, f"edge_case_tests_round{round_num}.py")

        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script_content)

        print(f"成功生成测试脚本: {script_path}")
        print(f"包含 {len(self.cases)} 个测试用例 (T{start_num}-T{end_num})")
        return script_path

    def run_tests(self, round_num: Optional[int] = None) -> List[Dict]:
        """执行测试并返回结果

        Args:
            round_num: 轮次编号，如果为 None 则使用最新的测试脚本

        Returns:
            测试结果列表
        """
        if round_num is None:
            # 查找最新的测试脚本
            test_files = list(Path(self.test_output_dir).glob("edge_case_tests_round*.py"))
            if not test_files:
                print("未找到测试脚本")
                return []
            test_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            script_path = str(test_files[0])
        else:
            script_path = os.path.join(self.test_output_dir, f"edge_case_tests_round{round_num}.py")

        if not os.path.exists(script_path):
            print(f"测试脚本不存在: {script_path}")
            return []

        print(f"执行测试脚本: {script_path}")
        try:
            result = subprocess.run(
                ["python3", script_path],
                capture_output=True,
                text=True,
                timeout=600  # 10分钟超时
            )

            test_results = [{
                'script': script_path,
                'success': result.returncode == 0,
                'stdout': result.stdout,
                'stderr': result.stderr,
                'timestamp': datetime.now().isoformat()
            }]

            if result.returncode == 0:
                print("测试执行完成")
            else:
                print(f"测试执行失败，返回码: {result.returncode}")
                if result.stderr:
                    print(f"错误输出:\\n{result.stderr}")

            return test_results
        except subprocess.TimeoutExpired:
            print("测试执行超时")
            return [{
                'script': script_path,
                'success': False,
                'error': 'timeout',
                'timestamp': datetime.now().isoformat()
            }]
        except Exception as e:
            print(f"执行测试时出错: {e}")
            return [{
                'script': script_path,
                'success': False,
                'error': str(e),
                'timestamp': datetime.now().isoformat()
            }]

    def generate_markdown(self, results: Optional[List[Dict]] = None) -> str:
        """生成符合 EDGE-CASE-TESTS.md 格式的 markdown

        Args:
            results: 测试结果列表，如果为 None 则生成模板

        Returns:
            生成的 markdown 字符串
        """
        if not self.cases:
            print("没有边缘案例可供生成报告")
            return ""

        # 确定轮次编号
        round_num = self._get_next_round_number() - 1
        start_num = self._get_next_test_number() - len(self.cases)
        end_num = start_num + len(self.cases) - 1

        markdown = f"""
## {datetime.now().strftime('%Y-%m-%d')} 第{round_num}轮

"""

        # 为每个测试用例生成报告
        for i, case in enumerate(self.cases):
            test_num = start_num + i
            title = case.get('title', '未知测试')
            description = case.get('description', '')
            steps = case.get('steps', [])
            priority = case.get('priority', 'low')
            source = case.get('source', 'unknown')
            source_url = case.get('source_url', '')

            if not results:
                status = "TODO"
                actual_result = "待测试"
            else:
                status = "PASS"
                actual_result = "待实现"

            # 生成操作步骤描述
            steps_text = ""
            if steps:
                steps_text = "\\n".join([f"{i+1}. {step}" for i, step in enumerate(steps[:3])])
            else:
                steps_text = description[:200]

            markdown += f"""### 测试T{test_num}: {title}
- **操作步骤**: {steps_text}
- **优先级**: {priority}
- **来源**: {source}
- **来源URL**: {source_url}
- **预期结果**: 待补充
- **实际结果**: {actual_result}
- **是否通过**: {status}
- **备注**: 自动生成的测试用例，需要补充具体操作步骤和验证逻辑

"""

        # 生成统计信息
        if results:
            success = results[0].get('success', False)
            markdown += f"""
### 第{round_num}轮统计
- **总计**: {len(self.cases)} 个边缘案例 (T{start_num}-T{end_num})
- **生成时间**: {datetime.now().isoformat()}
- **测试脚本**: {results[0].get('script', 'N/A')}
- **执行状态**: {"成功" if success else "失败"}"""
        else:
            markdown += f"""
### 第{round_num}轮统计
- **总计**: {len(self.cases)} 个边缘案例 (T{start_num}-T{end_num})
- **生成时间**: {datetime.now().isoformat()}
- **测试脚本**: 自动生成，待执行
- **执行状态**: 未执行

- **建议**: 需要为每个测试用例补充具体的操作步骤和验证逻辑"""

        return markdown

    def save_to_markdown(self, markdown_content: str) -> bool:
        """将生成的 markdown 追加到 EDGE-CASE-TESTS.md 文件

        Args:
            markdown_content: 要追加的 markdown 内容

        Returns:
            是否成功
        """
        try:
            os.makedirs(os.path.dirname(self.markdown_output), exist_ok=True)
            if os.path.exists(self.markdown_output):
                with open(self.markdown_output, 'a', encoding='utf-8') as f:
                    f.write("\\n\\n" + markdown_content)
            else:
                with open(self.markdown_output, 'w', encoding='utf-8') as f:
                    f.write(markdown_content)

            print(f"成功保存报告到: {self.markdown_output}")
            return True
        except Exception as e:
            print(f"保存 markdown 报告失败: {e}")
            return False


def main():
    """主函数 - 演示使用流程"""
    import argparse

    parser = argparse.ArgumentParser(description='边缘测试自动化脚本')
    parser.add_argument('--load', '-l', action='store_true', help='加载发现的边缘案例')
    parser.add_argument('--generate', '-g', action='store_true', help='生成测试脚本')
    parser.add_argument('--run', '-r', action='store_true', help='运行测试')
    parser.add_argument('--report', '-m', action='store_true', help='生成 Markdown 报告')
    parser.add_argument('--round', type=int, help='指定轮次编号')
    parser.add_argument('--discovery-output', default='docs/edge_cases.json',
                        help='边缘案例 JSON 文件路径')
    parser.add_argument('--test-dir', default='scripts', help='测试脚本输出目录')
    parser.add_argument('--markdown-output', default='docs/EDGE-CASE-TESTS.md',
                        help='Markdown 报告输出路径')
    parser.add_argument('--all', action='store_true', help='执行完整流程（加载+生成+运行+报告）')

    args = parser.parse_args()

    # 创建自动化实例
    automation = EdgeCaseAutomation(
        discovery_output=args.discovery_output,
        test_output_dir=args.test_dir,
        markdown_output=args.markdown_output
    )

    # 执行完整流程
    if args.all:
        automation.load_from_discovery()
        script_path = automation.generate_test_round(round_num=args.round)
        test_results = automation.run_tests()
        markdown_content = automation.generate_markdown(test_results)
        automation.save_to_markdown(markdown_content)
        return

    # 步骤 1: 加载发现的边缘案例
    if args.load:
        automation.load_from_discovery()

    # 步骤 2: 生成测试脚本
    script_path = ""
    if args.generate:
        script_path = automation.generate_test_round(round_num=args.round)

    # 步骤 3: 执行测试
    test_results = []
    if args.run:
        test_results = automation.run_tests(round_num=args.round)

    # 步骤 4: 生成 Markdown 报告
    if args.report:
        markdown_content = automation.generate_markdown(test_results)
        automation.save_to_markdown(markdown_content)

    # 如果没有指定任何操作，执行完整流程
    if not any([args.load, args.generate, args.run, args.report]):
        automation.load_from_discovery()
        script_path = automation.generate_test_round()
        test_results = automation.run_tests()
        markdown_content = automation.generate_markdown(test_results)
        automation.save_to_markdown(markdown_content)


if __name__ == "__main__":
    main()
