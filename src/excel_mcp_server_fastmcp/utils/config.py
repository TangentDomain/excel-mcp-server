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

# 安全验证配置
MAX_FILE_SIZE_MB = 50  # 最大文件大小（MB）

# 流式写入配置
STREAMING_WRITE_MIN_ROWS = 50  # 流式写入最小影响行数
STREAMING_WRITE_MIN_CHANGES = 100  # 流式写入最小单元格修改数
STREAMING_WRITE_MIN_FILE_SIZE_MB = 1  # 流式写入最小文件大小（MB）

# 数据质量评分配置
DATA_QUALITY_MAX_SCORE = 100.0  # 数据质量最大评分
DATA_QUALITY_PENALTY_PER_SUGGESTION = 5  # 每个优化建议的扣分值

# 数据密度配置
DATA_DENSITY_THRESHOLD = 50  # 数据密度阈值（百分比）
LARGE_OPERATION_CELL_THRESHOLD = 5000  # 大范围操作单元格数阈值

# Markdown表格配置
MARKDOWN_TABLE_MAX_ROWS = 50  # Markdown表格最大行数
