#!/usr/bin/env python3
"""
生成TrSkill配置表比较报告
"""

from src.server import excel_compare_files
import json
from datetime import datetime

def generate_report():
    """生成详细的配置表比较报告"""

    print("🚀 开始生成 TrSkill 配置表比较报告...")
    print("文件1: D:/tr/svn/trunk/配置表/测试配置/微小/TrSkill.xlsx")
    print("文件2: D:/tr/svn/trunk/配置表/战斗环境配置/TrSkill.xlsx")
    print("="*80)

    try:
        # 执行比较
        result = excel_compare_files(
            r'D:\tr\svn\trunk\配置表\测试配置\微小\TrSkill.xlsx',
            r'D:\tr\svn\trunk\配置表\战斗环境配置\TrSkill.xlsx'
        )

        if not result.get('success'):
            print(f"❌ 比较失败: {result.get('message', 'Unknown error')}")
            return None

        data = result.get('data', {})

        # 生成Markdown报告
        report_lines = []
        report_lines.append("# TrSkill配置表比较报告")
        report_lines.append("")
        report_lines.append(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append("")

        # 文件信息
        report_lines.append("## 📁 比较文件")
        report_lines.append("| 类型 | 文件路径 |")
        report_lines.append("|------|----------|")
        report_lines.append(f"| 源文件 | `{data.get('file1_path', 'N/A')}` |")
        report_lines.append(f"| 目标文件 | `{data.get('file2_path', 'N/A')}` |")
        report_lines.append("")

        # 总体概览
        total_diff = data.get('total_differences', 0)
        is_identical = data.get('identical', False)
        report_lines.append("## 📊 总体概览")
        report_lines.append(f"- **文件是否相同**: {'✅ 是' if is_identical else '❌ 否'}")
        report_lines.append(f"- **总差异数量**: {total_diff:,}")
        report_lines.append(f"- **比较工作表数**: {len(data.get('sheet_comparisons', []))}")
        report_lines.append("")

        # 工作表详情
        sheets = data.get('sheet_comparisons', [])
        if sheets:
            report_lines.append("## 📋 工作表比较详情")

            for i, sheet in enumerate(sheets, 1):
                sheet_name = sheet.get('sheet_name', f'工作表{i}')
                report_lines.append(f"### {i}. 📄 {sheet_name}")

                # 基本信息
                if 'summary' in sheet:
                    summary = sheet['summary']
                    report_lines.append("#### 📈 统计信息")
                    report_lines.append(f"- **新增对象**: {summary.get('added_rows', 0)}")
                    report_lines.append(f"- **删除对象**: {summary.get('removed_rows', 0)}")
                    report_lines.append(f"- **修改对象**: {summary.get('modified_rows', 0)}")
                    report_lines.append(f"- **总差异数**: {summary.get('total_differences', 0)}")
                    report_lines.append("")

                # ID对象变化详情
                if 'row_differences' in sheet:
                    row_diffs = sheet['row_differences']
                    if row_diffs:
                        report_lines.append("#### 🔍 ID对象变化详情")

                        # 按变化类型分组
                        added_items = []
                        removed_items = []
                        modified_items = []

                        for diff in row_diffs:
                            change_type = diff.get('change_type', 'unknown')
                            row_id = diff.get('row_id', 'N/A')
                            summary = diff.get('id_based_summary', f"{change_type}: ID {row_id}")

                            if 'added' in change_type.lower() or '新增' in summary:
                                added_items.append(summary)
                            elif 'removed' in change_type.lower() or '删除' in summary:
                                removed_items.append(summary)
                            else:
                                modified_items.append(summary)

                        # 新增对象
                        if added_items:
                            report_lines.append("##### 🆕 新增对象")
                            for item in added_items[:10]:  # 限制显示前10个
                                report_lines.append(f"- {item}")
                            if len(added_items) > 10:
                                report_lines.append(f"- ... 还有 {len(added_items) - 10} 个新增对象")
                            report_lines.append("")

                        # 删除对象
                        if removed_items:
                            report_lines.append("##### 🗑️ 删除对象")
                            for item in removed_items[:10]:  # 限制显示前10个
                                report_lines.append(f"- {item}")
                            if len(removed_items) > 10:
                                report_lines.append(f"- ... 还有 {len(removed_items) - 10} 个删除对象")
                            report_lines.append("")

                        # 修改对象
                        if modified_items:
                            report_lines.append("##### 🔄 修改对象")
                            for item in modified_items[:10]:  # 限制显示前10个
                                report_lines.append(f"- {item}")
                            if len(modified_items) > 10:
                                report_lines.append(f"- ... 还有 {len(modified_items) - 10} 个修改对象")
                            report_lines.append("")
                    else:
                        report_lines.append("该工作表无ID对象变化")
                        report_lines.append("")

                report_lines.append("---")
                report_lines.append("")

        # 总结
        report_lines.append("## 💡 比较总结")
        if is_identical:
            report_lines.append("两个配置表文件完全相同。")
        else:
            report_lines.append(f"检测到 **{total_diff:,}** 处差异，涉及 **{len(sheets)}** 个工作表。")

            # 差异分布
            if sheets:
                report_lines.append("")
                report_lines.append("### 📊 差异分布")
                report_lines.append("| 工作表 | 新增 | 删除 | 修改 | 总差异 |")
                report_lines.append("|--------|------|------|------|--------|")

                for sheet in sheets:
                    sheet_name = sheet.get('sheet_name', 'N/A')
                    if 'summary' in sheet:
                        summary = sheet['summary']
                        added = summary.get('added_rows', 0)
                        removed = summary.get('removed_rows', 0)
                        modified = summary.get('modified_rows', 0)
                        total = summary.get('total_differences', 0)
                        report_lines.append(f"| {sheet_name} | {added} | {removed} | {modified} | {total} |")
                    else:
                        report_lines.append(f"| {sheet_name} | - | - | - | - |")

        report_lines.append("")
        report_lines.append("---")
        report_lines.append(f"*报告生成于 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")

        # 保存报告
        report_content = "\n".join(report_lines)
        report_filename = f"TrSkill配置表比较报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"

        with open(report_filename, 'w', encoding='utf-8') as f:
            f.write(report_content)

        print(f"✅ 报告生成完成: {report_filename}")
        print(f"📄 报告包含 {len(report_lines)} 行内容")

        # 显示关键摘要
        print("\n" + "="*80)
        print("📋 关键摘要:")
        print(f"• 总差异数: {total_diff:,}")
        print(f"• 工作表数: {len(sheets)}")

        if sheets:
            total_added = sum(sheet.get('summary', {}).get('added_rows', 0) for sheet in sheets)
            total_removed = sum(sheet.get('summary', {}).get('removed_rows', 0) for sheet in sheets)
            total_modified = sum(sheet.get('summary', {}).get('modified_rows', 0) for sheet in sheets)

            print(f"• 🆕 总新增对象: {total_added}")
            print(f"• 🗑️ 总删除对象: {total_removed}")
            print(f"• 🔄 总修改对象: {total_modified}")

        return report_filename

    except Exception as e:
        print(f"❌ 生成报告失败: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    generate_report()
