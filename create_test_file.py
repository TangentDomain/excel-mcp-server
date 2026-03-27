import tempfile
import os
from pathlib import Path

# 创建真实的测试xlsx文件
test_data = {
    '角色': {
        'A1': 'ID', 'B1': '名称', 'C1': '职业', 'D1': '等级', 'E1': '属性',
        'A2': 1, 'B2': '战士', 'C2': '战士', 'D2': 50, 'E2': '火',
        'A3': 2, 'B3': '法师', 'C3': '法师', 'D3': 45, 'E3': '水',
        'A4': 3, 'B4': '刺客', 'C4': '刺客', 'D4': 48, 'E4': '暗',
        'A5': 4, 'B5': '牧师', 'C5': '牧师', 'D5': 42, 'E5': '光',
        'A6': 5, 'B6': '战士', 'C6': '战士', 'D6': 55, 'E6': '火',
    },
    '技能': {
        'A1': 'ID', 'B1': '名称', 'C1': '职业限制', 'D1': '伤害',
        'A2': 101, 'B2': '火球术', 'C2': '法师', 'D2': 100,
        'A3': 102, 'B3': '治愈术', 'C3': '牧师', 'D3': 80,
        'A4': 103, 'B4': '斩击', 'C4': '战士', 'D4': 120,
        'A5': 104, 'B5': '隐身', 'C5': '刺客', 'D5': 0,
        'A6': 105, 'B6': '爆炎', 'C6': '战士', 'D6': 150,
        'A7': 106, 'B7': '寒冰箭', 'C7': '法师', 'D7': 90,
    },
    '装备': {
        'A1': 'ID', 'B1': '名称', 'C1': '类型', 'D1': '职业要求', 'E1': '等级要求',
        'A2': 1001, 'B2': '火焰剑', 'C2': '武器', 'D2': '战士', 'E2': 40,
        'A3': 1002, 'B3': '冰杖', 'C3': '武器', 'D3': '法师', 'E4': 35,
        'A4': 1003, 'B4': '暗影匕首', 'C4': '武器', 'D4': '刺客', 'E5': 38,
        'A5': 1004, 'B5': '圣杖', 'C5': '武器', 'D5': '牧师', 'E6': 30,
        'A6': 1006, 'B6': '火焰盾', 'C6': '防具', 'D6': '战士', 'E7': 45,
    }
}

def create_test_xlsx():
    try:
        import openpyxl
        from openpyxl import Workbook
        
        # 创建临时目录
        temp_dir = Path(tempfile.mkdtemp(prefix='excel_mcp_test_'))
        print(f"创建测试文件目录: {temp_dir}")
        
        # 创建测试文件
        test_file = temp_dir / "test_game_data.xlsx"
        
        wb = Workbook()
        
        # 写入每个工作表
        for sheet_name, data in test_data.items():
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
                
            ws = wb.create_sheet(sheet_name)
            
            # 写入数据
            for cell_ref, value in data.items():
                ws[cell_ref] = value
                
        # 保存文件
        wb.save(test_file)
        print(f"测试文件创建成功: {test_file}")
        
        return str(test_file)
        
    except ImportError:
        print("openpyxl未安装，使用备选方案")
        # 使用csv作为备选
        import csv
        
        temp_dir = Path(tempfile.mkdtemp(prefix='excel_mcp_test_'))
        test_file = temp_dir / "test_game_data.csv"
        
        with open(test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # 简单的CSV格式
            writer.writerow(['Sheet', 'Cell', 'Value'])
            for sheet_name, data in test_data.items():
                for cell_ref, value in data.items():
                    writer.writerow([sheet_name, cell_ref, value])
                    
        print(f"CSV测试文件创建成功: {test_file}")
        return str(test_file)

if __name__ == "__main__":
    test_file = create_test_xlsx()
    print(f"测试文件路径: {test_file}")