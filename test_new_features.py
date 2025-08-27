#!/usr/bin/env python3
"""
新功能验证脚本
测试新添加的Excel MCP Server功能：文件操作、格式化等
"""
import os
import tempfile
from pathlib import Path

# 设置导入路径
import sys
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.server import (
    excel_create_file,
    excel_export_to_csv,
    excel_import_from_csv,
    excel_convert_format,
    excel_merge_files,
    excel_get_file_info,
    excel_merge_cells,
    excel_unmerge_cells,
    excel_set_borders,
    excel_set_row_height,
    excel_set_column_width,
    excel_update_range
)

def test_new_features():
    """测试新添加的功能"""
    print("🧪 开始测试Excel MCP Server新功能...")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        # 1. 测试文件创建和基本操作
        print("\n1️⃣ 测试文件创建...")
        excel_file = temp_path / "test_new_features.xlsx"
        result = excel_create_file(str(excel_file), ["数据表", "测试表"])
        print(f"   ✅ 创建文件: {result['success']}")

        # 添加一些测试数据
        test_data = [
            ["姓名", "年龄", "部门"],
            ["张三", 25, "技术部"],
            ["李四", 30, "销售部"],
            ["王五", 28, "运营部"]
        ]
        result = excel_update_range(str(excel_file), "A1:C4", test_data, "数据表")
        print(f"   ✅ 添加测试数据: {result['success']}")

        # 2. 测试文件信息获取
        print("\n2️⃣ 测试文件信息获取...")
        result = excel_get_file_info(str(excel_file))
        if result['success']:
            info = result['data']
            print(f"   ✅ 文件大小: {info['file_size_mb']} MB")
            print(f"   ✅ 工作表数量: {info['sheet_count']}")
            print(f"   ✅ 工作表名称: {info['sheet_names']}")

        # 3. 测试CSV导出
        print("\n3️⃣ 测试CSV导出...")
        csv_file = temp_path / "export_test.csv"
        result = excel_export_to_csv(str(excel_file), str(csv_file), "数据表")
        print(f"   ✅ CSV导出: {result['success']}")
        if result['success']:
            print(f"   ✅ 导出行数: {result['data']['row_count']}")

        # 4. 测试CSV导入
        print("\n4️⃣ 测试CSV导入...")
        imported_excel = temp_path / "imported_from_csv.xlsx"
        result = excel_import_from_csv(str(csv_file), str(imported_excel), "导入数据")
        print(f"   ✅ CSV导入: {result['success']}")
        if result['success']:
            print(f"   ✅ 导入行数: {result['data']['row_count']}")

        # 5. 测试格式转换
        print("\n5️⃣ 测试格式转换...")
        json_file = temp_path / "converted.json"
        result = excel_convert_format(str(excel_file), str(json_file), "json")
        print(f"   ✅ JSON转换: {result['success']}")

        # 6. 测试文件合并
        print("\n6️⃣ 测试文件合并...")
        merged_file = temp_path / "merged.xlsx"
        files_to_merge = [str(excel_file), str(imported_excel)]
        result = excel_merge_files(files_to_merge, str(merged_file), "sheets")
        print(f"   ✅ 文件合并: {result['success']}")
        if result['success']:
            print(f"   ✅ 合并文件数: {len(result['data']['merged_files'])}")
            print(f"   ✅ 总工作表数: {result['data']['total_sheets']}")

        # 7. 测试单元格格式化功能
        print("\n7️⃣ 测试单元格格式化...")

        # 测试合并单元格
        result = excel_merge_cells(str(excel_file), "数据表", "A1:C1")
        print(f"   ✅ 合并单元格: {result['success']}")

        # 测试设置边框
        result = excel_set_borders(str(excel_file), "数据表", "A1:C4", "thick")
        print(f"   ✅ 设置边框: {result['success']}")

        # 测试设置行高
        result = excel_set_row_height(str(excel_file), "数据表", 1, 25)
        print(f"   ✅ 设置行高: {result['success']}")

        # 测试设置列宽
        result = excel_set_column_width(str(excel_file), "数据表", 1, 15)
        print(f"   ✅ 设置列宽: {result['success']}")

        # 测试取消合并单元格
        result = excel_unmerge_cells(str(excel_file), "数据表", "A1:C1")
        print(f"   ✅ 取消合并单元格: {result['success']}")

        print(f"\n🎉 所有新功能测试完成！测试文件保存在: {temp_dir}")

        # 显示最终文件信息
        result = excel_get_file_info(str(excel_file))
        if result['success']:
            print(f"\n📊 最终文件信息:")
            info = result['data']
            for key, value in info.items():
                if key not in ['sheet_names']:  # 跳过长列表
                    print(f"   {key}: {value}")


if __name__ == "__main__":
    try:
        test_new_features()
        print("\n✅ 测试成功完成！")
    except Exception as e:
        print(f"\n❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
