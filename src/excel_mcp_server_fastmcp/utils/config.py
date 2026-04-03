"""配置常量

集中管理系统中使用的硬编码常量，便于维护和调整。
"""

# Excel 操作配置
MAX_SEARCH_FILES = 100  # 最大搜索文件数

# SQL 查询引擎配置
MAX_CACHE_SIZE = 20  # 最大DataFrame缓存文件数，防止内存泄漏
MAX_QUERY_CACHE_SIZE = 15  # 最大查询结果缓存数，防止内存泄漏
QUERY_CACHE_TTL = 300  # 查询缓存TTL（5分钟）
CACHE_TARGET_MEMORY_MB = 512.0  # 目标最大缓存内存（MB）

# 结果限制配置
MAX_RESULT_ROWS = 500  # 最大结果行数（保护AI上下文窗口）
