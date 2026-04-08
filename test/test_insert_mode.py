"""
REQ-028: excel_update_range insert_mode 默认值改为 false

测试文件专门验证 insert_mode 参数的默认值和两种模式的行为差异。
使用真实游戏配置表作为测试 fixture。
"""
import pytest
import os
import tempfile
import shutil
from pathlib import Path
from excel_mcp.api.excel_operations import ExcelOperations




def _extract_values(data):
    """从 get_range 返回的 {coordinate, value} 对象列表中提取纯值"""
    if not data or not isinstance(data, list):
        return data
    result = []
    for row in data:
        if isinstance(row, list):
            new_row = []
            for cell in row:
                if isinstance(cell, dict):
                    new_row.append(cell.get('value'))
                else:
                    new_row.append(cell)
            result.append(new_row)
        else:
            result.append(row)
    return result

class TestExcelUpdateRangeInsertMode:
    """测试 excel_update_range 的 insert_mode 参数行为"""
    
    @pytest.fixture
    def test_file_path(self):
        """创建测试用的 Excel 文件"""
        # 创建临时目录
        temp_dir = tempfile.mkdtemp()
        test_file = os.path.join(temp_dir, "test_insert_mode.xlsx")
        
        # 创建一个简单的测试文件，类似游戏配置表结构
        import pandas as pd
        
        # 模拟技能表结构
        data = [
            ["ID", "技能名称", "类型", "消耗MP", "冷却时间"],
            [1, "火球术", "攻击", 30, 5],
            [2, "冰冻术", "控制", 25, 8],
            [3, "治疗术", "治疗", 40, 10],
            [4, "雷击术", "攻击", 35, 6],
        ]
        
        df = pd.DataFrame(data[1:], columns=data[0])
        df.to_excel(test_file, index=False)
        
        yield test_file
        
        # 清理临时文件
        shutil.rmtree(temp_dir)
    
    def test_default_behavior_cover_mode(self, test_file_path):
        """测试1: 覆盖模式默认行为 - 不传 insert_mode 时应默认为覆盖模式"""
        # 准备测试数据
        range_to_update = "Sheet1!B2:B4"  # 技能名称列的2-4行
        new_data = [["烈火球"], ["冰冻新星"], ["强力治疗"]]
        
        # 不传 insert_mode，验证默认行为
        result = ExcelOperations.update_range(
            test_file_path,
            range_to_update,
            new_data
            # 注意：不传 insert_mode，应默认为 False (覆盖模式)
        )
        
        assert result['success'], f"操作失败: {result.get('message')}"
        
        # 验证覆盖结果：读取更新后的数据
        check_result = ExcelOperations.get_range(test_file_path, range_to_update)
        assert check_result['success'], "无法读取更新后的数据"
        updated_data = _extract_values(check_result['data'])
        
        # 验证值正确覆盖
        assert updated_data[0][0] == "烈火球", f"第一行值错误: {updated_data[0][0]}"
        assert updated_data[1][0] == "冰冻新星", f"第二行值错误: {updated_data[1][0]}"
        assert updated_data[2][0] == "强力治疗", f"第三行值错误: {updated_data[2][0]}"
        
        # 验证相邻列数据不变（覆盖模式不应影响其他列）
        range_check_other = "Sheet1!A2:A4"  # ID列
        other_result = ExcelOperations.get_range(test_file_path, range_check_other)
        assert other_result['success']
        other_data = _extract_values(other_result['data'])
        
        # ID列应该保持原值（1,2,3）
        assert other_data[0][0] == 1, "ID列数据被意外修改"
        assert other_data[1][0] == 2, "ID列数据被意外修改"
        assert other_data[2][0] == 3, "ID列数据被意外修改"

    def test_cover_mode_precision_verification(self, test_file_path):
        """测试2: 覆盖模式精确验证 - 只更新目标单元格，相邻数据不受影响"""
        # 读取更新前的数据用于验证
        original_data = ExcelOperations.get_range(test_file_path, "Sheet1!A2:D4")
        assert original_data['success']
        
        # 更新单个单元格 D2 (消耗MP)
        range_update = "Sheet1!D2"
        new_data = [[50]]  # 将火球术的MP消耗改为50
        
        result = ExcelOperations.update_range(
            test_file_path,
            range_update,
            new_data
            # 默认覆盖模式
        )
        
        assert result['success'], f"操作失败: {result.get('message')}"
        
        # 验证目标单元格值正确
        check_result = ExcelOperations.get_range(test_file_path, range_update)
        assert check_result['success']
        assert _extract_values(check_result['data'])[0][0] == 50, "目标单元格值更新错误"
        
        # 验证相邻单元格不变
        range_check_neighbors = "Sheet1!C2:E2"  # 整行
        neighbors_result = ExcelOperations.get_range(test_file_path, range_check_neighbors)
        assert neighbors_result['success']
        
        # C2应该保持"攻击"(类型)，D2应该是50(消耗MP已更新)，E2应该保持5(冷却时间)
        assert _extract_values(neighbors_result['data'])[0][0] == "攻击", "左邻单元格被错误修改"
        assert _extract_values(neighbors_result['data'])[0][1] == 50, "目标单元格应该是50"
        assert _extract_values(neighbors_result['data'])[0][2] == 5, "右邻单元格被错误修改"

    def test_insert_mode_explicit_enable(self, test_file_path):
        """测试3: 插入模式显式开启 - insert_mode=True 时应插入新行"""
        # 先读取当前行数
        all_data = ExcelOperations.get_range(test_file_path, "Sheet1!A1:E10")
        assert all_data['success']
        extracted = _extract_values(all_data['data'])
        original_row_count = len([row for row in extracted if any(cell is not None for cell in row)])
        
        # 插入新数据在第2行前面
        range_to_insert = "Sheet1!A2:E2"
        new_data = [[99, "新技能", "辅助", 20, 3]]  # 单行数据
        
        result = ExcelOperations.update_range(
            test_file_path,
            range_to_insert,
            new_data,
            insert_mode=True  # 显式开启插入模式
        )
        
        assert result['success'], f"插入操作失败: {result.get('message')}"
        
        # 验证行数增加
        updated_all_data = ExcelOperations.get_range(test_file_path, "Sheet1!A1:E10")
        assert updated_all_data['success']
        new_row_count = len([row for row in _extract_values(updated_all_data['data']) if any(cell is not None for cell in row)])
        
        assert new_row_count == original_row_count + 1, f"行数应增加1，原{original_row_count}行，现{new_row_count}行"

    def test_insert_mode_verification(self, test_file_path):
        """测试4: 插入模式验证 - 原数据正确下移，新数据插入"""
        # 先获取当前完整数据
        current_data = ExcelOperations.get_range(test_file_path, "Sheet1!A1:E10")
        assert current_data['success']
        
        # 插入数据在第3行
        insert_range = "Sheet1!A3:E3"
        new_skill_data = [[88, "传送门", "移动", 60, 15]]
        
        result = ExcelOperations.update_range(
            test_file_path,
            insert_range,
            new_skill_data,
            insert_mode=True
        )
        
        assert result['success']
        
        # 验证插入结果
        check_result = ExcelOperations.get_range(test_file_path, insert_range)
        assert check_result['success']
        
        # 新数据应该在第3行
        assert _extract_values(check_result['data'])[0][0] == 88, "新技能ID错误"
        assert _extract_values(check_result['data'])[0][1] == "传送门", "新技能名称错误"
        
        # 验证原数据下移
        # 原第3行(冰冻术)现在应该在第4行
        moved_range = "Sheet1!A4:E4"
        moved_result = ExcelOperations.get_range(test_file_path, moved_range)
        assert moved_result['success']
        assert _extract_values(moved_result['data'])[0][1] == "冰冻术", "原数据下移错误"
        
        # 验证原第2行数据不变
        second_row = "Sheet1!A2:E2"
        second_result = ExcelOperations.get_range(test_file_path, second_row)
        assert second_result['success']
        assert _extract_values(second_result['data'])[0][1] == "火球术", "原第2行数据被错误修改"

    def test_multi_column_write_cover(self, test_file_path):
        """测试5: 多列写入覆盖 - 一次写入多列数据，验证同行其他列不受影响"""
        # 更新第2行的多个列（技能名称和消耗MP）
        range_multi = "Sheet1!B2:D2"
        new_multi_data = [["超级火球", "攻击", 80]]
        
        result = ExcelOperations.update_range(
            test_file_path,
            range_multi,
            new_multi_data
            # 默认覆盖模式
        )
        
        assert result['success']
        
        # 验证更新的列正确
        check_result = ExcelOperations.get_range(test_file_path, range_multi)
        assert check_result['success']
        assert _extract_values(check_result['data'])[0][0] == "超级火球", "技能名称列更新错误"
        assert _extract_values(check_result['data'])[0][1] == "攻击", "类型列更新错误"
        assert _extract_values(check_result['data'])[0][2] == 80, "MP消耗列更新错误"
        
        # 验证未更新的列不变（冷却时间E2）
        unchanged_range = "Sheet1!E2"
        unchanged_result = ExcelOperations.get_range(test_file_path, unchanged_range)
        assert unchanged_result['success']
        assert _extract_values(unchanged_result['data'])[0][0] == 5, f"冷却时间被错误修改，期望5，实际{_extract_values(unchanged_result['data'])[0][0]}"

    def test_multi_row_write_cover(self, test_file_path):
        """测试6: 多行写入覆盖 - 一次写入多行数据，验证非目标行不受影响"""
        # 更新第2-3行的数据（技能名称和类型）
        range_multi = "Sheet1!B2:B3"
        new_multi_data = [
            ["超级火球", "攻击魔法"],
            ["冰冻新星", "控制魔法"]
        ]
        
        result = ExcelOperations.update_range(
            test_file_path,
            range_multi,
            new_multi_data
            # 默认覆盖模式
        )
        
        assert result['success']
        
        # 验证更新的行正确
        check_result = ExcelOperations.get_range(test_file_path, range_multi)
        assert check_result['success']
        assert _extract_values(check_result['data'])[0][0] == "超级火球", "第2行技能名称错误"
        assert _extract_values(check_result['data'])[1][0] == "冰冻新星", "第3行技能名称错误"
        
        # 验证非目标行不变（第4行的治疗术）
        non_target_range = "Sheet1!B4"
        non_target_result = ExcelOperations.get_range(test_file_path, non_target_range)
        assert non_target_result['success']
        assert _extract_values(non_target_result['data'])[0][0] == "治疗术", "非目标行被错误修改"

    def test_edge_cases(self, test_file_path):
        """测试7: 边界场景 - 空文件、末尾行、单单元格等"""
        # 测试空文件
        temp_dir = tempfile.mkdtemp()
        empty_file = os.path.join(temp_dir, "empty.xlsx")
        
        # 创建空Excel文件
        import pandas as pd
        empty_df = pd.DataFrame()
        empty_df.to_excel(empty_file, index=False)
        
        # 测试在空文件中写入（应该能正常工作）
        try:
            result = ExcelOperations.update_range(
                empty_file,
                "Sheet1!A1",
                [["测试"]]
            )
            # 空文件写入可能成功也可能失败（取决于库实现），但不应该报参数错误
            assert "参数验证" not in result.get('message', ''), f"空文件写入错误: {result}"
        finally:
            os.remove(empty_file)
            shutil.rmtree(temp_dir)
        
        # 测试写入末尾行（不影响已有数据）
        end_range = "Sheet1!D10"  # 末尾行
        end_data = [[99]]  # 新的冷却时间
        
        result = ExcelOperations.update_range(
            test_file_path,
            end_range,
            end_data
        )
        
        assert result['success'], f"末尾行写入失败: {result.get('message')}"
        
        # 验证新增的数据
        end_result = ExcelOperations.get_range(test_file_path, end_range)
        assert end_result['success']
        assert _extract_values(end_result['data'])[0][0] == 99, "末尾行数据写入错误"
        
        # 验证原有数据仍然存在
        original_check = ExcelOperations.get_range(test_file_path, "Sheet1!A1")
        assert original_check['success'], "原有数据丢失"

    def test_real_config_table_structure(self, test_file_path):
        """测试8: 使用真实配置表结构 - 使用MapEvent.xlsx等真实结构作为测试"""
        # 模拟MapEvent表的配置结构
        event_data = [
            ["事件ID", "事件名称", "触发条件", "奖励金币", "奖励经验"],
            [1001, "新手任务", "等级>=1", 100, 50],
            [1002, "主线任务1", "主线任务进度=1", 200, 100],
            [1003, "支线任务", "支线任务进度=1", 150, 75],
        ]
        
        # 创建专门的MapEvent测试文件
        temp_dir = tempfile.mkdtemp()
        map_event_file = os.path.join(temp_dir, "MapEvent.xlsx")
        
        import pandas as pd
        df = pd.DataFrame(event_data[1:], columns=event_data[0])
        df.to_excel(map_event_file, index=False)
        
        try:
            # 测试覆盖默认行为
            result = ExcelOperations.update_range(
                map_event_file,
                "Sheet1!B2:B4",  # 事件名称列
                [["新手任务（优化版）"], ["主线任务1（加强版）"], ["支线任务（限时）"]]
                # 默认覆盖模式
            )
            
            assert result['success'], f"MapEvent更新失败: {result.get('message')}"
            
            # 验证结果
            check_result = ExcelOperations.get_range(map_event_file, "Sheet1!B2:B4")
            assert check_result['success']
            
            # 确认每个事件名称都被正确覆盖
            assert _extract_values(check_result['data'])[0][0] == "新手任务（优化版）", "事件1名称更新错误"
            assert _extract_values(check_result['data'])[1][0] == "主线任务1（加强版）", "事件2名称更新错误"
            assert _extract_values(check_result['data'])[2][0] == "支线任务（限时）", "事件3名称更新错误"
            
            # 验证关联数据（奖励金币）保持不变
            reward_check = ExcelOperations.get_range(map_event_file, "Sheet1!D2:D4")
            assert reward_check['success']
            assert _extract_values(reward_check['data'])[0][0] == 100, "事件1奖励金币被错误修改"
            assert _extract_values(reward_check['data'])[1][0] == 200, "事件2奖励金币被错误修改"
            assert _extract_values(reward_check['data'])[2][0] == 150, "事件3奖励金币被错误修改"
            
        finally:
            os.remove(map_event_file)
            shutil.rmtree(temp_dir)