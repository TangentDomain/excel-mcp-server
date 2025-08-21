"""
Excel MCP Server - 公式计算缓存管理器

提供高性能的公式计算缓存机制
"""

import os
import time
import hashlib
import tempfile
import threading
from typing import Any, Dict, Optional, Tuple
from dataclasses import dataclass
from openpyxl import load_workbook, Workbook
import logging

logger = logging.getLogger(__name__)


@dataclass
class CacheEntry:
    """缓存条目"""
    value: Any
    timestamp: float
    access_count: int
    file_mtime: float
    formula_hash: str


@dataclass 
class WorkbookCache:
    """工作簿缓存"""
    workbook: Workbook
    temp_file_path: str
    file_mtime: float
    timestamp: float
    access_count: int


class FormulaCalculationCache:
    """公式计算缓存管理器"""
    
    def __init__(self, max_size: int = 100, ttl: int = 3600):
        """
        初始化缓存管理器
        
        Args:
            max_size: 最大缓存条目数量
            ttl: 缓存生存时间（秒）
        """
        self.max_size = max_size
        self.ttl = ttl
        self._cache: Dict[str, CacheEntry] = {}
        self._workbook_cache: Dict[str, WorkbookCache] = {}
        self._lock = threading.RLock()
        
        # 统计信息
        self.hit_count = 0
        self.miss_count = 0
        
    def _generate_cache_key(
        self,
        file_path: str,
        formula: str,
        context_sheet: Optional[str] = None
    ) -> str:
        """生成缓存键"""
        key_data = f"{file_path}:{formula}:{context_sheet or ''}"
        return hashlib.md5(key_data.encode()).hexdigest()
    
    def _generate_formula_hash(self, formula: str) -> str:
        """生成公式哈希"""
        return hashlib.md5(formula.encode()).hexdigest()
    
    def _get_file_mtime(self, file_path: str) -> float:
        """获取文件修改时间"""
        try:
            return os.path.getmtime(file_path)
        except OSError:
            return 0.0
    
    def _is_cache_valid(self, entry: CacheEntry, file_path: str) -> bool:
        """检查缓存是否有效"""
        current_time = time.time()
        current_mtime = self._get_file_mtime(file_path)
        
        # 检查TTL
        if current_time - entry.timestamp > self.ttl:
            return False
        
        # 检查文件是否被修改
        if current_mtime > entry.file_mtime:
            return False
        
        return True
    
    def _cleanup_expired_entries(self):
        """清理过期缓存条目"""
        current_time = time.time()
        expired_keys = []
        
        for key, entry in self._cache.items():
            if current_time - entry.timestamp > self.ttl:
                expired_keys.append(key)
        
        for key in expired_keys:
            del self._cache[key]
        
        # 清理工作簿缓存
        expired_workbook_keys = []
        for key, wb_cache in self._workbook_cache.items():
            if current_time - wb_cache.timestamp > self.ttl:
                expired_workbook_keys.append(key)
        
        for key in expired_workbook_keys:
            wb_cache = self._workbook_cache[key]
            try:
                # 清理临时文件
                if os.path.exists(wb_cache.temp_file_path):
                    os.unlink(wb_cache.temp_file_path)
            except OSError as e:
                logger.warning(f"清理临时文件失败: {e}")
            del self._workbook_cache[key]
    
    def _evict_lru_entries(self):
        """驱逐最少使用的缓存条目"""
        if len(self._cache) <= self.max_size:
            return
        
        # 按访问次数和时间排序，移除最少使用的
        sorted_entries = sorted(
            self._cache.items(),
            key=lambda x: (x[1].access_count, x[1].timestamp)
        )
        
        # 移除最老和最少使用的条目
        entries_to_remove = len(self._cache) - self.max_size + 10  # 多删除10个为后续留空间
        for i in range(min(entries_to_remove, len(sorted_entries))):
            key = sorted_entries[i][0]
            del self._cache[key]
    
    def get_cached_workbook(self, file_path: str) -> Optional[Tuple[Workbook, str]]:
        """获取缓存的工作簿"""
        with self._lock:
            file_mtime = self._get_file_mtime(file_path)
            cache_key = f"wb:{file_path}:{file_mtime}"
            
            if cache_key in self._workbook_cache:
                wb_cache = self._workbook_cache[cache_key]
                
                # 检查缓存是否仍然有效
                if (time.time() - wb_cache.timestamp <= self.ttl and 
                    wb_cache.file_mtime >= file_mtime):
                    wb_cache.access_count += 1
                    wb_cache.timestamp = time.time()
                    return wb_cache.workbook, wb_cache.temp_file_path
                else:
                    # 清理过期缓存
                    try:
                        if os.path.exists(wb_cache.temp_file_path):
                            os.unlink(wb_cache.temp_file_path)
                    except OSError:
                        pass
                    del self._workbook_cache[cache_key]
            
            return None
    
    def cache_workbook(
        self, 
        file_path: str, 
        workbook: Workbook, 
        temp_file_path: str
    ) -> None:
        """缓存工作簿"""
        with self._lock:
            file_mtime = self._get_file_mtime(file_path)
            cache_key = f"wb:{file_path}:{file_mtime}"
            
            self._workbook_cache[cache_key] = WorkbookCache(
                workbook=workbook,
                temp_file_path=temp_file_path,
                file_mtime=file_mtime,
                timestamp=time.time(),
                access_count=1
            )
    
    def get(
        self,
        file_path: str,
        formula: str,
        context_sheet: Optional[str] = None
    ) -> Optional[Any]:
        """获取缓存的计算结果"""
        with self._lock:
            cache_key = self._generate_cache_key(file_path, formula, context_sheet)
            
            if cache_key in self._cache:
                entry = self._cache[cache_key]
                
                # 检查缓存是否有效
                if self._is_cache_valid(entry, file_path):
                    # 更新访问统计
                    entry.access_count += 1
                    self.hit_count += 1
                    
                    logger.debug(f"缓存命中: {formula}")
                    return entry.value
                else:
                    # 移除无效缓存
                    del self._cache[cache_key]
            
            self.miss_count += 1
            logger.debug(f"缓存未命中: {formula}")
            return None
    
    def put(
        self,
        file_path: str,
        formula: str,
        value: Any,
        context_sheet: Optional[str] = None
    ) -> None:
        """存储计算结果到缓存"""
        with self._lock:
            # 清理过期条目
            self._cleanup_expired_entries()
            
            # 如果缓存已满，驱逐旧条目
            self._evict_lru_entries()
            
            cache_key = self._generate_cache_key(file_path, formula, context_sheet)
            file_mtime = self._get_file_mtime(file_path)
            formula_hash = self._generate_formula_hash(formula)
            
            self._cache[cache_key] = CacheEntry(
                value=value,
                timestamp=time.time(),
                access_count=1,
                file_mtime=file_mtime,
                formula_hash=formula_hash
            )
            
            logger.debug(f"缓存已存储: {formula}")
    
    def clear(self) -> None:
        """清空所有缓存"""
        with self._lock:
            self._cache.clear()
            
            # 清理工作簿缓存和临时文件
            for wb_cache in self._workbook_cache.values():
                try:
                    if os.path.exists(wb_cache.temp_file_path):
                        os.unlink(wb_cache.temp_file_path)
                except OSError as e:
                    logger.warning(f"清理临时文件失败: {e}")
            
            self._workbook_cache.clear()
            
            # 重置统计
            self.hit_count = 0
            self.miss_count = 0
    
    def invalidate_file(self, file_path: str) -> None:
        """使指定文件的所有缓存失效"""
        with self._lock:
            keys_to_remove = []
            for key, entry in self._cache.items():
                if key.startswith(hashlib.md5(f"{file_path}:".encode()).hexdigest()[:8]):
                    keys_to_remove.append(key)
            
            for key in keys_to_remove:
                del self._cache[key]
            
            # 清理工作簿缓存
            wb_keys_to_remove = []
            for key in self._workbook_cache.keys():
                if file_path in key:
                    wb_keys_to_remove.append(key)
            
            for key in wb_keys_to_remove:
                wb_cache = self._workbook_cache[key]
                try:
                    if os.path.exists(wb_cache.temp_file_path):
                        os.unlink(wb_cache.temp_file_path)
                except OSError as e:
                    logger.warning(f"清理临时文件失败: {e}")
                del self._workbook_cache[key]
    
    def get_stats(self) -> Dict[str, Any]:
        """获取缓存统计信息"""
        with self._lock:
            total_requests = self.hit_count + self.miss_count
            hit_rate = (self.hit_count / total_requests * 100) if total_requests > 0 else 0
            
            return {
                'cache_size': len(self._cache),
                'workbook_cache_size': len(self._workbook_cache),
                'max_size': self.max_size,
                'hit_count': self.hit_count,
                'miss_count': self.miss_count,
                'hit_rate': round(hit_rate, 2),
                'ttl': self.ttl
            }


# 全局缓存实例
_formula_cache = FormulaCalculationCache()


def get_formula_cache() -> FormulaCalculationCache:
    """获取全局公式缓存实例"""
    return _formula_cache
