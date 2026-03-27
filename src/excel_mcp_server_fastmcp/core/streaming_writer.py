"""
Excel MCP Server - 流式写入模块

copy-modify-write 方案：用 calamine 读取 + write_only 模式写入，
显著降低大文件修改操作的内存占用和耗时。

适用场景：批量插入、大批量数据修改等数据密集型操作。
权衡：write_only 模式不保留单元格级格式（字体/填充/边框/合并），
但保留列宽、行高、数据值。对游戏配置表场景（数据优先）友好。
"""

import logging
import os
import shutil
import tempfile
from typing import List, Dict, Any, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

try:
    from python_calamine import CalamineWorkbook
    _HAS_CALAMINE = True
except ImportError:
    _HAS_CALAMINE = False
    logger.debug("python-calamine未安装，流式写入将降级到openpyxl")


class StreamingWriter:
    """流式写入器：calamine读取 + write_only写入

    内存占用与文件大小无关（只缓存当前行），
    适合大数据量文件的修改操作。
    """

    def __init__(self, file_path: str):
        self.file_path = file_path

    @staticmethod
    def is_available() -> bool:
        """检查流式写入是否可用（需要calamine）"""
        return _HAS_CALAMINE

    @classmethod
    def batch_insert_rows(
        cls,
        file_path: str,
        sheet_name: str,
        data: List[Dict[str, Any]],
        header_row: int = 1,
        preserve_col_widths: bool = True,
    ) -> Tuple[bool, str, Dict[str, Any]]:
        """流式批量插入行（calamine读取 + write_only写入）

        Args:
            file_path: Excel文件路径
            sheet_name: 目标工作表名
            data: 行数据列表，每行为{列名: 值}字典
            header_row: 表头所在行号（默认1）
            preserve_col_widths: 是否保留列宽（需额外openpyxl读取，但比全量加载快）

        Returns:
            (success, message, metadata)
        """
        if not _HAS_CALAMINE:
            return False, "calamine未安装，无法使用流式写入", {}

        try:
            # 1. 用 calamine 读取所有工作表数据
            cal_wb = CalamineWorkbook.from_path(file_path)
            all_sheets_data = {}
            target_sheet_data = None
            col_map = {}

            for sn in cal_wb.sheet_names:
                rows = cal_wb.get_sheet_by_name(sn).to_python()
                all_sheets_data[sn] = rows

                if sn == sheet_name and rows:
                    # 提取表头映射
                    if header_row <= len(rows):
                        for col_idx, cell_val in enumerate(rows[header_row - 1], 1):
                            if cell_val is not None:
                                col_map[str(cell_val).strip()] = col_idx
                    target_sheet_data = rows

            if sheet_name not in all_sheets_data:
                available = ', '.join(all_sheets_data.keys())
                return False, f"工作表不存在: {sheet_name}（可用: {available}）", {}

            if not col_map:
                return False, f"未找到表头（行{header_row}）", {}

            # 2. 构建新数据（原有数据 + 新行）
            new_rows = list(target_sheet_data)
            unknown_cols = set()

            for row_data in data:
                if not isinstance(row_data, dict):
                    return False, f"行数据必须是字典，实际类型: {type(row_data).__name__}", {}

                # 计算最大列数
                max_col = max(col_map.values()) if col_map else 0
                new_row = [None] * max_col

                for col_name, value in row_data.items():
                    col_name_stripped = col_name.strip()
                    if col_name_stripped in col_map:
                        new_row[col_map[col_name_stripped] - 1] = value
                    else:
                        unknown_cols.add(col_name_stripped)

                new_rows.append(new_row)

            # 3. 读取列宽（用非只读模式，需要column_dimensions）
            col_widths = {}
            if preserve_col_widths:
                try:
                    wb_meta = load_workbook(file_path)
                    if sheet_name in wb_meta.sheetnames:
                        ws_meta = wb_meta[sheet_name]
                        col_widths = {
                            col_letter: dim.width
                            for col_letter, dim in ws_meta.column_dimensions.items()
                            if dim.width
                        }
                    wb_meta.close()
                except Exception as e:
                    logger.warning(f"读取列宽失败，跳过: {e}")

            # 4. 用 write_only 模式写入临时文件
            wb_out = Workbook(write_only=True)

            # 更新目标工作表数据为含新行的版本
            all_sheets_data[sheet_name] = new_rows

            for sn in cal_wb.sheet_names:
                ws = wb_out.create_sheet(title=sn)
                sheet_rows = all_sheets_data[sn]

                if not sheet_rows:
                    continue

                # 设置列宽（write_only支持）
                if sn == sheet_name and col_widths:
                    for col_letter, width in col_widths.items():
                        try:
                            ws.column_dimensions[col_letter].width = width
                        except Exception:
                            pass

                # 逐行写入（流式，不积累内存）
                for row in sheet_rows:
                    ws.append(row)

            # 5. 原子替换：写入临时文件 → 替换原文件
            fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(file_path))
            os.close(fd)

            try:
                wb_out.save(tmp_path)
                wb_out.close()

                # 原子替换
                shutil.move(tmp_path, file_path)
            except Exception:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
                raise

            start_row = len(target_sheet_data) + 1
            end_row = start_row + len(data) - 1

            return True, (
                f"流式批量插入 {len(data)} 行"
                f"（第{start_row}-{end_row}行）"
            ), {
                'action': 'batch_insert_streaming',
                'start_row': start_row,
                'end_row': end_row,
                'inserted_count': len(data),
                'total_rows_after': len(new_rows),
                'unknown_columns': sorted(unknown_cols)[:5] if unknown_cols else [],
                'mode': 'streaming',
                'col_widths_preserved': len(col_widths) > 0,
            }

        except Exception as e:
            logger.error(f"流式批量插入失败: {e}")
            return False, f"流式写入失败: {str(e)}", {}

    @classmethod
    def upsert_row(
        cls,
        file_path: str,
        sheet_name: str,
        key_column: str,
        key_value: Any,
        updates: Dict[str, Any],
        header_row: int = 1,
        preserve_col_widths: bool = True,
    ) -> Tuple[bool, str, Dict[str, Any]]:
        """流式upsert行：按键列查找并更新/插入

        Args:
            file_path: Excel文件路径
            sheet_name: 目标工作表名
            key_column: 匹配列名
            key_value: 匹配值
            updates: 列值字典
            header_row: 表头行号
            preserve_col_widths: 是否保留列宽

        Returns:
            (success, message, metadata)
        """
        if not _HAS_CALAMINE:
            return False, "calamine未安装，无法使用流式写入", {}

        try:
            cal_wb = CalamineWorkbook.from_path(file_path)
            all_sheets_data = {}
            target_rows = None
            col_map = {}

            for sn in cal_wb.sheet_names:
                rows = cal_wb.get_sheet_by_name(sn).to_python()
                all_sheets_data[sn] = rows

                if sn == sheet_name and rows and header_row <= len(rows):
                    for col_idx, cell_val in enumerate(rows[header_row - 1], 1):
                        if cell_val is not None:
                            col_map[str(cell_val).strip()] = col_idx
                    target_rows = list(rows)  # 深拷贝

            if sheet_name not in all_sheets_data:
                available = ', '.join(all_sheets_data.keys())
                return False, f"工作表不存在: {sheet_name}（可用: {available}）", {}

            if key_column not in col_map:
                actual = list(col_map.keys())[:10]
                return False, f"键列 '{key_column}' 不存在。实际列名: {', '.join(actual)}", {}

            key_col_idx = col_map[key_column]

            # 查找匹配行
            target_row = None
            for row_num, row in enumerate(target_rows[header_row:], start=header_row + 1):
                if key_col_idx <= len(row):
                    cell_val = row[key_col_idx - 1]
                    if cell_val is not None:
                        # calamine 把整数读成浮点数（2→2.0），需要标准化比较
                        a, b = str(cell_val).strip(), str(key_value).strip()
                        try:
                            a_norm = float(a)
                            b_norm = float(b)
                            matched = a_norm == b_norm
                        except (ValueError, TypeError):
                            matched = a == b
                        if matched:
                            target_row = row_num
                            break

            action = 'update'
            if target_row is not None:
                # UPDATE: 修改已有行
                updated_cols = []
                row = target_rows[target_row - 1]
                for col_name, value in updates.items():
                    col_name_stripped = col_name.strip()
                    if col_name_stripped in col_map:
                        row[col_map[col_name_stripped] - 1] = value
                        updated_cols.append(col_name_stripped)
            else:
                # INSERT: 追加新行
                action = 'insert'
                max_col = max(col_map.values()) if col_map else 0
                new_row = [None] * max_col
                updated_cols = []

                for col_name, col_idx in sorted(col_map.items(), key=lambda x: x[1]):
                    if col_name in updates:
                        new_row[col_idx - 1] = updates[col_name]
                        updated_cols.append(col_name)
                    elif col_name == key_column:
                        new_row[col_idx - 1] = key_value
                        updated_cols.append(col_name)

                target_rows.append(new_row)
                target_row = len(target_rows)

            # 更新 all_sheets_data
            all_sheets_data[sheet_name] = target_rows

            # 读取列宽
            col_widths = {}
            if preserve_col_widths:
                try:
                    wb_meta = load_workbook(file_path)
                    if sheet_name in wb_meta.sheetnames:
                        ws_meta = wb_meta[sheet_name]
                        col_widths = {
                            col_letter: dim.width
                            for col_letter, dim in ws_meta.column_dimensions.items()
                            if dim.width
                        }
                    wb_meta.close()
                except Exception as e:
                    logger.warning(f"读取列宽失败，跳过: {e}")

            # write_only 写入
            wb_out = Workbook(write_only=True)
            for sn in cal_wb.sheet_names:
                ws = wb_out.create_sheet(title=sn)
                sheet_rows = all_sheets_data[sn]

                if sn == sheet_name and col_widths:
                    for col_letter, width in col_widths.items():
                        try:
                            ws.column_dimensions[col_letter].width = width
                        except Exception:
                            pass

                for row in sheet_rows:
                    ws.append(row)

            fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(file_path))
            os.close(fd)

            try:
                wb_out.save(tmp_path)
                wb_out.close()
                shutil.move(tmp_path, file_path)
            except Exception:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
                raise

            return True, (
                f"流式{'更新' if action == 'update' else '插入'}行 {target_row}"
                f"（键列 '{key_column}'='{key_value}'），"
                f"修改了 {len(updated_cols)} 列"
            ), {
                'action': action,
                'row': target_row,
                'updated_columns': updated_cols,
                'mode': 'streaming',
                'col_widths_preserved': len(col_widths) > 0,
            }

        except Exception as e:
            logger.error(f"流式upsert失败: {e}")
            return False, f"流式写入失败: {str(e)}", {}
