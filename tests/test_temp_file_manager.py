# -*- coding: utf-8 -*-
"""
Temp File Manager 完整测试套件

覆盖 temp_file_manager.py 中未被测试的功能
"""

import pytest
import tempfile
import os
import time
from pathlib import Path

from src.utils.temp_file_manager import TempFileManager


class TestTempFileManager:
    """TempFileManager 完整测试"""

    def test_create_temp_excel_file_default(self):
        """测试创建默认临时Excel文件"""
        temp_file = TempFileManager.create_temp_excel_file()
        
        assert temp_file is not None
        assert temp_file.endswith('.xlsx')
        
        # 清理
        TempFileManager.cleanup_temp_file(temp_file)

    def test_create_temp_excel_file_custom_prefix(self):
        """测试自定义前缀"""
        temp_file = TempFileManager.create_temp_excel_file(prefix="custom_prefix")
        
        assert "custom_prefix" in temp_file
        assert temp_file.endswith('.xlsx')
        
        # 清理
        TempFileManager.cleanup_temp_file(temp_file)

    def test_create_temp_excel_file_custom_suffix(self):
        """测试自定义后缀"""
        temp_file = TempFileManager.create_temp_excel_file(suffix=".xlsm")
        
        assert temp_file.endswith('.xlsm')
        
        # 清理
        TempFileManager.cleanup_temp_file(temp_file)

    def test_create_temp_csv_file(self):
        """测试创建临时CSV文件"""
        temp_file = TempFileManager.create_temp_csv_file()
        
        assert temp_file is not None
        assert temp_file.endswith('.csv')
        
        # 清理
        TempFileManager.cleanup_temp_file(temp_file)

    def test_create_temp_csv_file_custom_prefix(self):
        """测试自定义前缀的CSV文件"""
        temp_file = TempFileManager.create_temp_csv_file(prefix="my_csv")
        
        assert "my_csv" in temp_file
        assert temp_file.endswith('.csv')
        
        # 清理
        TempFileManager.cleanup_temp_file(temp_file)

    def test_create_temp_json_file(self):
        """测试创建临时JSON文件"""
        temp_file = TempFileManager.create_temp_json_file()
        
        assert temp_file is not None
        assert temp_file.endswith('.json')
        
        # 清理
        TempFileManager.cleanup_temp_file(temp_file)

    def test_create_temp_json_file_custom_prefix(self):
        """测试自定义前缀的JSON文件"""
        temp_file = TempFileManager.create_temp_json_file(prefix="my_json")
        
        assert "my_json" in temp_file
        assert temp_file.endswith('.json')
        
        # 清理
        TempFileManager.cleanup_temp_file(temp_file)

    def test_cleanup_temp_file_existing(self):
        """测试清理存在的临时文件"""
        # 创建临时文件
        temp_file = TempFileManager.create_temp_excel_file()
        
        # 先创建一个实际的文件
        with open(temp_file, 'w') as f:
            f.write("test content")
        
        # 验证文件存在
        assert os.path.exists(temp_file)
        
        # 清理
        result = TempFileManager.cleanup_temp_file(temp_file)
        
        # 验证清理结果
        assert result is True
        assert not os.path.exists(temp_file)

    def test_cleanup_temp_file_not_existing(self):
        """测试清理不存在的文件"""
        result = TempFileManager.cleanup_temp_file("/nonexistent/file.xlsx")
        
        assert result is False

    def test_get_temp_directory(self):
        """测试获取系统临时目录"""
        temp_dir = TempFileManager.get_temp_directory()
        
        assert temp_dir is not None
        assert os.path.exists(temp_dir)
        assert os.path.isdir(temp_dir)
        
        # 验证是系统临时目录
        system_temp = tempfile.gettempdir()
        assert temp_dir == system_temp

    def test_create_temp_dir_default(self):
        """测试创建临时目录"""
        temp_dir = TempFileManager.create_temp_dir()
        
        assert temp_dir is not None
        assert os.path.exists(temp_dir)
        assert os.path.isdir(temp_dir)
        
        # 清理
        import shutil
        shutil.rmtree(temp_dir)

    def test_create_temp_dir_custom_prefix(self):
        """测试自定义前缀的临时目录"""
        temp_dir = TempFileManager.create_temp_dir(prefix="my_custom_dir_")
        
        assert "my_custom_dir_" in temp_dir
        assert os.path.exists(temp_dir)
        assert os.path.isdir(temp_dir)
        
        # 清理
        import shutil
        shutil.rmtree(temp_dir)

    def test_temp_file_uniqueness(self):
        """测试临时文件唯一性"""
        # 创建多个临时文件
        files = []
        for _ in range(5):
            temp_file = TempFileManager.create_temp_excel_file(prefix="unique_test_")
            files.append(temp_file)
        
        # 验证文件名唯一
        unique_files = set(files)
        assert len(unique_files) == len(files)
        
        # 清理所有文件
        for f in files:
            TempFileManager.cleanup_temp_file(f)

    def test_temp_file_in_system_temp_dir(self):
        """测试临时文件在系统临时目录中"""
        temp_file = TempFileManager.create_temp_excel_file()
        
        system_temp = tempfile.gettempdir()
        assert temp_file.startswith(system_temp)
        
        # 清理
        TempFileManager.cleanup_temp_file(temp_file)

    def test_temp_dir_in_system_temp_dir(self):
        """测试临时目录在系统临时目录中"""
        temp_dir = TempFileManager.create_temp_dir(prefix="temp_dir_test_")
        
        system_temp = tempfile.gettempdir()
        assert temp_dir.startswith(system_temp)
        
        # 清理
        import shutil
        shutil.rmtree(temp_dir)
