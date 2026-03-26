"""
Excel MCP Server - 临时文件管理器

提供安全的临时文件创建和管理功能
"""

import tempfile
import os
import time
import random
import string
from typing import Optional


class TempFileManager:
    """临时文件管理器"""

    @staticmethod
    def create_temp_excel_file(prefix: str = "excel_mcp", suffix: str = ".xlsx") -> str:
        """
        创建临时Excel文件，返回文件路径

        Args:
            prefix: 文件名前缀
            suffix: 文件扩展名

        Returns:
            str: 临时文件路径
        """
        temp_dir = tempfile.gettempdir()

        # 生成唯一文件名：前缀 + 进程ID + 时间戳 + 随机字符串
        timestamp = int(time.time())
        process_id = os.getpid()
        random_str = ''.join(random.choices(string.ascii_lowercase + string.digits, k=6))

        filename = f"{prefix}_{process_id}_{timestamp}_{random_str}{suffix}"
        temp_file_path = os.path.join(temp_dir, filename)

        # 确保目录存在
        os.makedirs(os.path.dirname(temp_file_path), exist_ok=True)

        return temp_file_path

    @staticmethod
    def create_temp_csv_file(prefix: str = "excel_mcp") -> str:
        """
        创建临时CSV文件，返回文件路径

        Args:
            prefix: 文件名前缀

        Returns:
            str: 临时文件路径
        """
        return TempFileManager.create_temp_excel_file(prefix, ".csv")

    @staticmethod
    def create_temp_json_file(prefix: str = "excel_mcp") -> str:
        """
        创建临时JSON文件，返回文件路径

        Args:
            prefix: 文件名前缀

        Returns:
            str: 临时文件路径
        """
        return TempFileManager.create_temp_excel_file(prefix, ".json")

    @staticmethod
    def cleanup_temp_file(file_path: str) -> bool:
        """
        清理临时文件

        Args:
            file_path: 要删除的文件路径

        Returns:
            bool: 是否成功删除
        """
        try:
            if os.path.exists(file_path):
                os.unlink(file_path)
                return True
            return False
        except Exception:
            return False

    @staticmethod
    def get_temp_directory() -> str:
        """
        获取系统临时目录

        Returns:
            str: 系统临时目录路径
        """
        return tempfile.gettempdir()

    @staticmethod
    def create_temp_dir(prefix: str = "excel_mcp") -> str:
        """
        创建临时目录

        Args:
            prefix: 目录名前缀

        Returns:
            str: 临时目录路径
        """
        temp_dir = tempfile.mkdtemp(prefix=prefix)
        return temp_dir


if __name__ == "__main__":
    # 测试代码
    print("测试临时文件管理器...")

    # 测试Excel文件创建
    temp_excel = TempFileManager.create_temp_excel_file()
    print(f"Excel临时文件: {temp_excel}")

    # 测试CSV文件创建
    temp_csv = TempFileManager.create_temp_csv_file()
    print(f"CSV临时文件: {temp_csv}")

    # 测试JSON文件创建
    temp_json = TempFileManager.create_temp_json_file()
    print(f"JSON临时文件: {temp_json}")

    # 验证在系统临时目录中
    system_temp = tempfile.gettempdir()
    print(f"系统临时目录: {system_temp}")
    print(f"Excel文件在系统临时目录: {temp_excel.startswith(system_temp)}")

    # 清理测试文件
    TempFileManager.cleanup_temp_file(temp_excel)
    TempFileManager.cleanup_temp_file(temp_csv)
    TempFileManager.cleanup_temp_file(temp_json)
    print("测试文件已清理")
    print("测试完成！")