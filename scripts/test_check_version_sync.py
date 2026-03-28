#!/usr/bin/env python3
"""
测试自动化版本检查脚本
"""

import pytest
import tempfile
import os
from pathlib import Path
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import check_version_sync

def test_get_version_from_pyproject():
    """测试从pyproject.toml获取版本"""
    version = check_version_sync.get_version_from_pyproject()
    assert version == "1.6.35", f"Expected 1.6.35, got {version}"

def test_get_version_from_init():
    """测试从__init__.py获取版本"""
    version = check_version_sync.get_version_from_init()
    assert version == "1.6.35", f"Expected 1.6.35, got {version}"

def test_get_version_from_readme():
    """测试从README获取版本"""
    version = check_version_sync.get_version_from_readme(Path("README.md"))
    assert version == "1.6.35", f"Expected 1.6.35, got {version}"

def test_get_version_from_changelog():
    """测试从CHANGELOG获取版本"""
    version = check_version_sync.get_version_from_changelog()
    assert version == "1.6.35", f"Expected 1.6.35, got {version}"

def test_version_consistency():
    """测试版本一致性检查"""
    result = check_version_sync.check_version_consistency()
    assert result == True, "版本应该是一致的"

def test_latest_version_from_files():
    """测试从文件获取最新版本"""
    version = check_version_sync.get_latest_version_from_files()
    assert version == "1.6.35", f"Expected 1.6.35, got {version}"

if __name__ == "__main__":
    pytest.main([__file__, "-v"])