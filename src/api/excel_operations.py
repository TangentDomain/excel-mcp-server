"""
Excel MCP Server - Excel操作API模块

提供高内聚的Excel业务操作功能，包含完整的参数验证、业务逻辑、错误处理和结果格式化

@intention: 将Excel操作的具体实现从server.py中分离，提高代码内聚性和可维护性
"""

import logging
from typing import Dict, Any, List, Optional, Union

from ..core.excel_reader import ExcelReader
from ..core.excel_writer import ExcelWriter
from ..core.excel_manager import ExcelManager
from ..utils.formatter import format_operation_result

logger = logging.getLogger(__name__)


class ExcelOperations:
    """
    @class ExcelOperations
    @brief Excel业务操作的高内聚封装
    @intention 提供完整的Excel操作功能，包含参数验证、错误处理、结果格式化
    """

    # ==================== 日志系统 ====================
    DEBUG_LOG_ENABLED: bool = False
    _LOG_PREFIX = '[API][ExcelOperations]'

    # ==================== 主干API ====================

    @classmethod
    def get_range(
        cls,
        file_path: str,
        range_expression: str,
        include_formatting: bool = False
    ) -> Dict[str, Any]:
        """
        @intention 获取Excel文件中指定范围的数据，提供完整的业务逻辑处理

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            range_expression: 范围表达式，必须包含工作表名
            include_formatting: 是否包含格式信息

        Returns:
            Dict: 标准化的操作结果

        Example:
            result = ExcelOperations.get_range("data.xlsx", "Sheet1!A1:C10")
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始获取范围数据: {range_expression}")

        try:
            # 步骤1: 验证参数格式
            validation_result = cls._validate_range_format(range_expression)
            if not validation_result['valid']:
                return cls._format_error_result(validation_result['error'])

            # 步骤2: 执行数据读取
            reader = ExcelReader(file_path)
            result = reader.get_range(range_expression, include_formatting)
            reader.close()

            # 步骤3: 格式化结果
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"获取范围数据失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def update_range(
        cls,
        file_path: str,
        range_expression: str,
        data: List[List[Any]],
        preserve_formulas: bool = True,
        insert_mode: bool = True
    ) -> Dict[str, Any]:
        """
        @intention 更新Excel文件中指定范围的数据，支持插入和覆盖模式

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            range_expression: 范围表达式，必须包含工作表名
            data: 二维数组数据 [[row1], [row2], ...]
            preserve_formulas: 是否保留现有公式
            insert_mode: 数据写入模式 (默认值: True)
                - True: 插入模式，在指定位置插入新行然后写入数据（更安全）
                - False: 覆盖模式，直接覆盖目标范围的现有数据

        Returns:
            Dict: 标准化的操作结果

        Example:
            data = [["姓名", "年龄"], ["张三", 25]]
            # 插入模式（默认，更安全）
            result = ExcelOperations.update_range("test.xlsx", "Sheet1!A1:B2", data)
            # 覆盖模式（显式指定）
            result = ExcelOperations.update_range("test.xlsx", "Sheet1!A1:B2", data, insert_mode=False)
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始更新范围数据: {range_expression}, 模式: {'插入' if insert_mode else '覆盖'}")

        try:
            # 步骤1: 验证参数格式
            validation_result = cls._validate_range_format(range_expression)
            if not validation_result['valid']:
                return cls._format_error_result(validation_result['error'])

            # 步骤2: 执行数据写入
            writer = ExcelWriter(file_path)
            result = writer.update_range(range_expression, data, preserve_formulas, insert_mode)

            # 步骤3: 格式化结果
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"更新范围数据失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def list_sheets(cls, file_path: str) -> Dict[str, Any]:
        """
        @intention 获取Excel文件中所有工作表信息，提供完整的文件结构概览

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)

        Returns:
            Dict: 包含工作表列表、总数量、活动工作表等信息

        Example:
            result = ExcelOperations.list_sheets("data.xlsx")
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始获取工作表列表: {file_path}")

        try:
            # 步骤1: 读取工作表信息
            reader = ExcelReader(file_path)
            result = reader.list_sheets()

            # 步骤2: 提取和格式化数据
            sheets = [sheet.name for sheet in result.data] if result.data else []

            response = {
                'success': True,
                'sheets': sheets,
                'file_path': file_path,
                'total_sheets': result.metadata.get('total_sheets', len(sheets)) if result.metadata else len(sheets)
            }

            # 步骤3: 清理资源
            reader.close()

            return response

        except Exception as e:
            error_msg = f"获取工作表列表失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def get_headers(
        cls,
        file_path: str,
        sheet_name: str,
        header_row: int = 1,
        max_columns: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        @intention 获取指定工作表的表头信息，支持游戏开发双行模式（字段描述+字段名）

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            sheet_name: 工作表名称
            header_row: 表头起始行号 (1-based，默认从第1行开始获取两行)
            max_columns: 最大列数限制，None表示自动截取到空列

        Returns:
            Dict: 包含双行表头信息
            {
                'success': bool,
                'data': List[str],  # 字段名列表（兼容性）
                'headers': List[str],  # 字段名列表（兼容性）
                'descriptions': List[str],  # 字段描述列表（第1行）
                'field_names': List[str],   # 字段名列表（第2行）
                'header_count': int,
                'sheet_name': str,
                'header_row': int,
                'message': str
            }

        Example:
            result = ExcelOperations.get_headers("data.xlsx", "Sheet1")
            # 第1行：['技能ID描述', '技能名称描述', '技能类型描述']
            # 第2行：['skill_id', 'skill_name', 'skill_type']
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始获取双行表头: {sheet_name}")

        try:
            # 步骤1: 构建双行范围表达式
            range_expression = cls._build_header_range(sheet_name, header_row, max_columns, dual_row=True)

            # 步骤2: 读取表头数据（两行）
            reader = ExcelReader(file_path)
            result = reader.get_range(range_expression)
            reader.close()

            if not result.success:
                return cls._format_error_result(f"无法读取表头数据: {result.message}")

            # 步骤3: 解析双行表头信息
            header_info = cls._parse_dual_header_data(result.data, max_columns)

            return {
                'success': True,
                'data': header_info['field_names'],  # 兼容性字段，返回字段名
                'headers': header_info['field_names'],  # 兼容性字段，返回字段名
                'descriptions': header_info['descriptions'],  # 字段描述（第1行）
                'field_names': header_info['field_names'],    # 字段名（第2行）
                'header_count': len(header_info['field_names']),
                'sheet_name': sheet_name,
                'header_row': header_row,
                'message': f"成功获取{len(header_info['field_names'])}个表头字段（描述+字段名）"
            }

        except Exception as e:
            error_msg = f"获取表头失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def create_file(
        cls,
        file_path: str,
        sheet_names: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """
        @intention 创建新的Excel文件，支持自定义工作表配置

        Args:
            file_path: 新文件路径 (必须以.xlsx或.xlsm结尾)
            sheet_names: 工作表名称列表，None表示默认工作表

        Returns:
            Dict: 包含创建结果和文件信息

        Example:
            result = ExcelOperations.create_file("new_file.xlsx", ["数据", "分析"])
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始创建文件: {file_path}")

        try:
            # 步骤1: 执行文件创建
            result = ExcelManager.create_file(file_path, sheet_names)

            # 步骤2: 格式化结果
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"创建文件失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    # ==================== 分支实现 ====================

    # --- 参数验证 ---
    @classmethod
    def _validate_range_format(cls, range_expression: str) -> Dict[str, Any]:
        """验证范围表达式格式"""
        if not range_expression or not range_expression.strip():
            return {'valid': False, 'error': 'range参数不能为空'}

        if '!' not in range_expression:
            return {
                'valid': False,
                'error': f"range必须包含工作表名。当前格式: '{range_expression}'，正确格式示例: 'Sheet1!A1:B2'"
            }

        return {'valid': True}

    @classmethod
    def _build_header_range(cls, sheet_name: str, header_row: int, max_columns: Optional[int], dual_row: bool = False) -> str:
        """构建表头范围表达式，支持单行或双行模式"""
        if max_columns:
            # 如果指定了最大列数，使用具体范围
            from openpyxl.utils import get_column_letter
            end_column = get_column_letter(max_columns)
            if dual_row:
                # 双行模式：获取连续两行
                end_row = header_row + 1
                return f"{sheet_name}!A{header_row}:{end_column}{end_row}"
            else:
                # 单行模式（保持兼容性）
                return f"{sheet_name}!A{header_row}:{end_column}{header_row}"
        else:
            # 否则使用一个合理的默认范围（读取前100列）
            if dual_row:
                # 双行模式：获取连续两行
                end_row = header_row + 1
                return f"{sheet_name}!A{header_row}:CV{end_row}"  # CV = 第100列
            else:
                # 单行模式（保持兼容性）
                return f"{sheet_name}!A{header_row}:CV{header_row}"  # CV = 第100列

    @classmethod
    def _parse_header_data(cls, data: List[List], max_columns: Optional[int]) -> List[str]:
        """解析表头数据"""
        headers = []
        if data and len(data) > 0:
            first_row = data[0]
            for i, cell_info in enumerate(first_row):
                # 处理CellInfo对象和普通值
                cell_value = getattr(cell_info, 'value', cell_info) if hasattr(cell_info, 'value') else cell_info

                # 转换为字符串并清理
                if cell_value is not None:
                    str_value = str(cell_value).strip()
                    if str_value != "":
                        headers.append(str_value)
                    else:
                        # 空字符串的处理
                        if max_columns:
                            headers.append("")  # 指定max_columns时保留空字符串
                        else:
                            break  # 否则停止
                else:
                    # None值的处理
                    if max_columns:
                        headers.append("")  # 指定max_columns时将None转为空字符串
                    else:
                        break  # 否则停止

                # 如果指定了max_columns，检查是否已达到限制
                if max_columns and len(headers) >= max_columns:
                    break

        return headers

    @classmethod
    def _parse_dual_header_data(cls, data: List[List], max_columns: Optional[int]) -> Dict[str, List[str]]:
        """解析双行表头数据（字段描述 + 字段名），支持空值fallback机制"""
        descriptions = []
        field_names = []

        if not data or len(data) < 2:
            # 如果数据不足两行，返回空结果
            return {
                'descriptions': descriptions,
                'field_names': field_names
            }

        # 解析第一行（字段描述）
        first_row = data[0] if len(data) > 0 else []
        # 解析第二行（字段名）
        second_row = data[1] if len(data) > 1 else []

        # 确定实际处理的列数
        max_cols = max(len(first_row), len(second_row))  # 改为取最大值，不遗漏任何列
        if max_columns:
            max_cols = min(max_cols, max_columns)

        # 导入列名转换工具
        from openpyxl.utils import get_column_letter

        for i in range(max_cols):
            # 处理字段描述（第1行）
            desc_cell = first_row[i] if i < len(first_row) else None
            desc_value = getattr(desc_cell, 'value', desc_cell) if hasattr(desc_cell, 'value') else desc_cell
            desc_str = str(desc_value).strip() if desc_value is not None and str(desc_value).strip() else ""

            # 处理字段名（第2行）
            name_cell = second_row[i] if i < len(second_row) else None
            name_value = getattr(name_cell, 'value', name_cell) if hasattr(name_cell, 'value') else name_cell
            name_str = str(name_value).strip() if name_value is not None and str(name_value).strip() else ""

            # 🆕 智能Fallback机制
            column_letter = get_column_letter(i + 1)  # 1-based列名：A, B, C...

            # 描述为空时使用列标识作为fallback
            if not desc_str:
                desc_str = f"列{column_letter}"  # 中文：列A, 列B, 列C...

            # 字段名为空时使用列名作为fallback
            if not name_str:
                name_str = column_letter.lower()  # 小写：a, b, c...

            # 🆕 检查是否应该停止（简化的停止条件）
            # 只有在没有指定max_columns时才进行智能停止
            if not max_columns:
                # 检查原始数据是否为完全空（描述和字段名都是原始空值）
                desc_is_empty = (desc_cell is None or
                               (hasattr(desc_cell, 'value') and desc_cell.value is None) or
                               (not hasattr(desc_cell, 'value') and desc_cell is None))
                name_is_empty = (name_cell is None or
                               (hasattr(name_cell, 'value') and name_cell.value is None) or
                               (not hasattr(name_cell, 'value') and name_cell is None))

                # 如果当前列完全为空，检查接下来连续3列是否也为空
                if desc_is_empty and name_is_empty:
                    consecutive_empty = 0
                    for j in range(i, min(i + 3, max_cols)):  # 检查当前及后续2列
                        check_desc = first_row[j] if j < len(first_row) else None
                        check_name = second_row[j] if j < len(second_row) else None

                        desc_empty = (check_desc is None or
                                    (hasattr(check_desc, 'value') and check_desc.value is None) or
                                    (not hasattr(check_desc, 'value') and check_desc is None))
                        name_empty = (check_name is None or
                                    (hasattr(check_name, 'value') and check_name.value is None) or
                                    (not hasattr(check_name, 'value') and check_name is None))

                        if desc_empty and name_empty:
                            consecutive_empty += 1
                        else:
                            break

                    # 如果连续3列都为空，则停止
                    if consecutive_empty >= 3:
                        break

            descriptions.append(desc_str)
            field_names.append(name_str)

            # 如果指定了max_columns，检查是否已达到限制
            if max_columns and len(field_names) >= max_columns:
                break

        return {
            'descriptions': descriptions,
            'field_names': field_names
        }

    @classmethod
    def search(
        cls,
        file_path: str,
        pattern: str,
        sheet_name: Optional[str] = None,
        case_sensitive: bool = False,
        whole_word: bool = False,
        use_regex: bool = False,
        include_values: bool = True,
        include_formulas: bool = False,
        range: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        @intention 在Excel文件中搜索单元格内容（VSCode风格搜索选项）

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            pattern: 搜索模式（正则表达式或字面字符串）
            sheet_name: 工作表名称 (可选)
            case_sensitive: 大小写敏感
            whole_word: 全词匹配
            use_regex: 启用正则表达式
            include_values: 是否搜索单元格值
            include_formulas: 是否搜索公式内容
            range: 搜索范围表达式

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            search_type = "正则" if use_regex else ("全词" if whole_word else "字符串")
            case_info = "大小写敏感" if case_sensitive else "忽略大小写"
            logger.info(f"{cls._LOG_PREFIX} 开始{search_type}搜索({case_info}): {pattern}")

        try:
            from ..core.excel_search import ExcelSearcher
            import re

            searcher = ExcelSearcher(file_path)

            # 构建正则表达式模式
            if use_regex:
                # 直接使用用户提供的正则表达式
                regex_pattern = pattern
            else:
                # 将字面字符串转义为正则表达式
                escaped_pattern = re.escape(pattern)

                # 如果是全词匹配，添加单词边界
                if whole_word:
                    regex_pattern = r'\b' + escaped_pattern + r'\b'
                else:
                    regex_pattern = escaped_pattern

            # 构建正则表达式标志
            regex_flags = "" if case_sensitive else "i"

            result = searcher.regex_search(regex_pattern, regex_flags, include_values, include_formulas, sheet_name, range)
            return format_operation_result(result)

        except Exception as e:
            search_type = "正则" if use_regex else ("全词" if whole_word else "字符串")
            error_msg = f"{search_type}搜索失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def search_directory(
        cls,
        directory_path: str,
        pattern: str,
        case_sensitive: bool = False,
        whole_word: bool = False,
        use_regex: bool = False,
        include_values: bool = True,
        include_formulas: bool = False,
        recursive: bool = True,
        file_extensions: Optional[List[str]] = None,
        file_pattern: Optional[str] = None,
        max_files: int = 100
    ) -> Dict[str, Any]:
        """
        @intention 在目录下的所有Excel文件中搜索内容（VSCode风格搜索选项）

        Args:
            directory_path: 目录路径
            pattern: 搜索模式（正则表达式或字面字符串）
            case_sensitive: 大小写敏感
            whole_word: 全词匹配
            use_regex: 启用正则表达式
            其他参数同search方法

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            search_type = "正则" if use_regex else ("全词" if whole_word else "字符串")
            case_info = "大小写敏感" if case_sensitive else "忽略大小写"
            logger.info(f"{cls._LOG_PREFIX} 开始目录{search_type}搜索({case_info}): {directory_path}")

        try:
            from ..core.excel_search import ExcelSearcher
            import re

            # 构建正则表达式模式
            if use_regex:
                # 直接使用用户提供的正则表达式
                regex_pattern = pattern
            else:
                # 将字面字符串转义为正则表达式
                escaped_pattern = re.escape(pattern)

                # 如果是全词匹配，添加单词边界
                if whole_word:
                    regex_pattern = r'\b' + escaped_pattern + r'\b'
                else:
                    regex_pattern = escaped_pattern

            # 构建正则表达式标志
            regex_flags = "" if case_sensitive else "i"

            result = ExcelSearcher.search_directory_static(
                directory_path, regex_pattern, regex_flags, include_values, include_formulas,
                recursive, file_extensions, file_pattern, max_files
            )
            return format_operation_result(result)

        except Exception as e:
            search_type = "正则" if use_regex else ("全词" if whole_word else "字符串")
            error_msg = f"目录{search_type}搜索失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def get_sheet_headers(cls, file_path: str) -> Dict[str, Any]:
        """
        @intention 获取Excel文件中所有工作表的双行表头信息（字段描述+字段名）

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)

        Returns:
            Dict: 包含所有工作表的双行表头信息
            {
                'success': bool,
                'sheets_with_headers': [
                    {
                        'name': str,
                        'headers': List[str],       # 字段名（兼容性）
                        'descriptions': List[str],  # 字段描述（第1行）
                        'field_names': List[str],   # 字段名（第2行）
                        'header_count': int
                    }
                ],
                'file_path': str,
                'total_sheets': int
            }
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始获取所有工作表双行表头: {file_path}")

        try:
            # 步骤1: 获取所有工作表列表
            sheets_result = cls.list_sheets(file_path)
            if not sheets_result.get('success'):
                return sheets_result

            # 步骤2: 获取每个工作表的双行表头
            sheets_with_headers = []
            sheets = sheets_result.get('sheets', [])

            for sheet_name in sheets:
                try:
                    header_result = cls.get_headers(file_path, sheet_name, header_row=1)

                    if header_result.get('success'):
                        headers = header_result.get('headers', [])
                        descriptions = header_result.get('descriptions', [])
                        field_names = header_result.get('field_names', [])

                        # 如果没有获取到field_names，使用headers作为fallback
                        if not field_names and headers:
                            field_names = headers

                        sheets_with_headers.append({
                            'name': sheet_name,
                            'headers': field_names,         # 兼容性字段，使用字段名
                            'descriptions': descriptions,   # 字段描述（第1行）
                            'field_names': field_names,     # 字段名（第2行）
                            'header_count': len(field_names)
                        })
                    else:
                        sheets_with_headers.append({
                            'name': sheet_name,
                            'headers': [],
                            'descriptions': [],
                            'field_names': [],
                            'header_count': 0,
                            'error': header_result.get('error', '未知错误')
                        })

                except Exception as e:
                    sheets_with_headers.append({
                        'name': sheet_name,
                        'headers': [],
                        'descriptions': [],
                        'field_names': [],
                        'header_count': 0,
                        'error': str(e)
                    })

            return format_operation_result({
                'success': True,
                'sheets_with_headers': sheets_with_headers,
                'file_path': file_path,
                'total_sheets': len(sheets)
            })

        except Exception as e:
            error_msg = f"获取工作表表头失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def insert_rows(
        cls,
        file_path: str,
        sheet_name: str,
        row_index: int,
        count: int = 1
    ) -> Dict[str, Any]:
        """
        @intention 在指定位置插入空行

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            sheet_name: 工作表名称
            row_index: 插入位置 (1-based)
            count: 插入行数

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始插入行: {sheet_name} 第{row_index}行")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            result = writer.insert_rows(sheet_name, row_index, count)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"插入行操作失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def insert_columns(
        cls,
        file_path: str,
        sheet_name: str,
        column_index: int,
        count: int = 1
    ) -> Dict[str, Any]:
        """
        @intention 在指定位置插入空列

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            sheet_name: 工作表名称
            column_index: 插入位置 (1-based)
            count: 插入列数

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始插入列: {sheet_name} 第{column_index}列")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            result = writer.insert_columns(sheet_name, column_index, count)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"插入列操作失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def export_to_csv(
        cls,
        file_path: str,
        output_path: str,
        sheet_name: Optional[str] = None,
        encoding: str = "utf-8"
    ) -> Dict[str, Any]:
        """
        @intention 将Excel工作表导出为CSV文件

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            output_path: 输出CSV文件路径
            sheet_name: 工作表名称 (默认使用活动工作表)
            encoding: 文件编码 (默认: utf-8)

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始导出为CSV: {output_path}")

        try:
            from ..core.excel_converter import ExcelConverter
            converter = ExcelConverter(file_path)
            result = converter.export_to_csv(output_path, sheet_name, encoding)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"导出为CSV失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def import_from_csv(
        cls,
        csv_path: str,
        output_path: str,
        sheet_name: str = "Sheet1",
        encoding: str = "utf-8",
        has_header: bool = True
    ) -> Dict[str, Any]:
        """
        @intention 从CSV文件导入数据创建Excel文件

        Args:
            csv_path: CSV文件路径
            output_path: 输出Excel文件路径
            sheet_name: 工作表名称
            encoding: CSV文件编码
            has_header: 是否包含表头行

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始从CSV导入: {csv_path}")

        try:
            from ..core.excel_converter import ExcelConverter
            result = ExcelConverter.import_from_csv(csv_path, output_path, sheet_name, encoding, has_header)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"从CSV导入失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def convert_format(
        cls,
        input_path: str,
        output_path: str,
        target_format: str = "xlsx"
    ) -> Dict[str, Any]:
        """
        @intention 转换Excel文件格式

        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            target_format: 目标格式

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始格式转换: {input_path} -> {output_path}")

        try:
            from ..core.excel_converter import ExcelConverter
            result = ExcelConverter.convert_format(input_path, output_path, target_format)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"文件格式转换失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def merge_files(
        cls,
        input_files: List[str],
        output_path: str,
        merge_mode: str = "sheets"
    ) -> Dict[str, Any]:
        """
        @intention 合并多个Excel文件

        Args:
            input_files: 输入文件路径列表
            output_path: 输出文件路径
            merge_mode: 合并模式

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始合并文件: {len(input_files)}个文件")

        try:
            from ..core.excel_converter import ExcelConverter
            result = ExcelConverter.merge_files(input_files, output_path, merge_mode)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"合并Excel文件失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def get_file_info(cls, file_path: str) -> Dict[str, Any]:
        """
        @intention 获取Excel文件的详细信息

        Args:
            file_path: Excel文件路径

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始获取文件信息: {file_path}")

        try:
            from ..core.excel_manager import ExcelManager
            result = ExcelManager.get_file_info(file_path)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"获取文件信息失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def create_sheet(
        cls,
        file_path: str,
        sheet_name: str,
        index: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        @intention 在文件中创建新工作表

        Args:
            file_path: Excel文件路径
            sheet_name: 新工作表名称
            index: 插入位置

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始创建工作表: {sheet_name}")

        try:
            from ..core.excel_manager import ExcelManager
            manager = ExcelManager(file_path)
            result = manager.create_sheet(sheet_name, index)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"创建工作表失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def delete_sheet(cls, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """
        @intention 删除指定工作表

        Args:
            file_path: Excel文件路径
            sheet_name: 要删除的工作表名称

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始删除工作表: {sheet_name}")

        try:
            from ..core.excel_manager import ExcelManager
            manager = ExcelManager(file_path)
            result = manager.delete_sheet(sheet_name)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"删除工作表失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def rename_sheet(
        cls,
        file_path: str,
        old_name: str,
        new_name: str
    ) -> Dict[str, Any]:
        """
        @intention 重命名工作表

        Args:
            file_path: Excel文件路径
            old_name: 当前工作表名称
            new_name: 新工作表名称

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始重命名工作表: {old_name} -> {new_name}")

        try:
            from ..core.excel_manager import ExcelManager
            manager = ExcelManager(file_path)
            result = manager.rename_sheet(old_name, new_name)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"重命名工作表失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def delete_rows(
        cls,
        file_path: str,
        sheet_name: str,
        row_index: int,
        count: int = 1
    ) -> Dict[str, Any]:
        """
        @intention 删除指定行

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            row_index: 起始行号 (1-based)
            count: 删除行数

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始删除行: {sheet_name} 第{row_index}行")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            result = writer.delete_rows(sheet_name, row_index, count)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"删除行操作失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def delete_columns(
        cls,
        file_path: str,
        sheet_name: str,
        column_index: int,
        count: int = 1
    ) -> Dict[str, Any]:
        """
        @intention 删除指定列

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            column_index: 起始列号 (1-based)
            count: 删除列数

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始删除列: {sheet_name} 第{column_index}列")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            result = writer.delete_columns(sheet_name, column_index, count)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"删除列操作失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def format_cells(
        cls,
        file_path: str,
        sheet_name: str,
        range: str,
        formatting: Optional[Dict[str, Any]] = None,
        preset: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        @intention 设置单元格格式

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            range: 目标范围
            formatting: 自定义格式配置
            preset: 预设样式

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始格式化单元格: {range}")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            # 处理预设格式
            if preset:
                preset_formats = {
                    'title': {'font': {'name': '微软雅黑', 'size': 14, 'bold': True}, 'alignment': {'horizontal': 'center'}},
                    'header': {'font': {'name': '微软雅黑', 'size': 11, 'bold': True}, 'fill': {'color': 'D9D9D9'}},
                    'data': {'font': {'name': '微软雅黑', 'size': 10}},
                    'highlight': {'fill': {'color': 'FFFF00'}},
                    'currency': {'number_format': '¥#,##0.00'}
                }
                formatting = preset_formats.get(preset, formatting or {})

            # 构建完整的range表达式
            if '!' not in range:
                range_expression = f"{sheet_name}!{range}"
            else:
                range_expression = range

            result = writer.format_cells(range_expression, formatting or {})
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"单元格格式化失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    # --- 错误处理 ---
    @classmethod
    def _format_error_result(cls, error_message: str) -> Dict[str, Any]:
        """创建标准化的错误响应"""
        return {
            'success': False,
            'error': error_message,
            'data': None
        }

    # --- 单元格操作扩展 ---
    @classmethod
    def merge_cells(cls, file_path: str, sheet_name: str, range: str) -> Dict[str, Any]:
        """
        @intention 合并指定范围的单元格
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始合并单元格: {range}")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            result = writer.merge_cells(range, sheet_name)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"合并单元格失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def unmerge_cells(cls, file_path: str, sheet_name: str, range: str) -> Dict[str, Any]:
        """
        @intention 取消合并指定范围的单元格
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始取消合并单元格: {range}")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            result = writer.unmerge_cells(range, sheet_name)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"取消合并单元格失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def set_borders(cls, file_path: str, sheet_name: str, range: str,
                   border_style: str = "thin") -> Dict[str, Any]:
        """
        @intention 为指定范围设置边框样式
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始设置边框: {range}, 样式: {border_style}")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            result = writer.set_borders(range, border_style, sheet_name)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"设置边框失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def set_row_height(cls, file_path: str, sheet_name: str, row_index: int,
                      height: float, count: int = 1) -> Dict[str, Any]:
        """
        @intention 调整指定行的高度
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始调整行高: 行{row_index}, 高度{height}, 数量{count}")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)

            # ExcelWriter.set_row_height(row_number, height, sheet_name)
            for i in range(count):
                row_num = row_index + i
                result = writer.set_row_height(row_num, height, sheet_name)
                if not result.success:
                    break

            return format_operation_result(result)

        except Exception as e:
            error_msg = f"调整行高失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def set_column_width(cls, file_path: str, sheet_name: str, column_index: int,
                        width: float, count: int = 1) -> Dict[str, Any]:
        """
        @intention 调整指定列的宽度
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始调整列宽: 列{column_index}, 宽度{width}, 数量{count}")

        try:
            from ..core.excel_writer import ExcelWriter
            from openpyxl.utils import get_column_letter

            writer = ExcelWriter(file_path)

            # ExcelWriter.set_column_width(column, width, sheet_name)
            for i in range(count):
                col_idx = column_index + i
                column_letter = get_column_letter(col_idx)
                result = writer.set_column_width(column_letter, width, sheet_name)
                if not result.success:
                    break

            return format_operation_result(result)

        except Exception as e:
            error_msg = f"调整列宽失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def compare_sheets(cls, file1_path: str, sheet1_name: str, file2_path: str,
                      sheet2_name: str, id_column: Union[int, str] = 1,
                      header_row: int = 1) -> Dict[str, Any]:
        """
        @intention 比较两个Excel工作表，识别ID对象的新增、删除、修改
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始比较工作表: {file1_path}:{sheet1_name} vs {file2_path}:{sheet2_name}")

        try:
            from ..core.excel_compare import ExcelComparer
            from ..models.types import ComparisonOptions

            # 创建比较选项
            options = ComparisonOptions()
            comparer = ExcelComparer(options)

            # 执行比较 - 使用正确的参数顺序
            result = comparer.compare_sheets(
                file1_path, sheet1_name, file2_path, sheet2_name, options
            )
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"比较工作表失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    # --- 公式操作扩展 ---
    @classmethod
    def set_formula(cls, file_path: str, sheet_name: str, cell_range: str,
                   formula: str) -> Dict[str, Any]:
        """
        @intention 设置指定单元格或区域的公式
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始设置公式: {cell_range} = {formula}")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            result = writer.set_formula(sheet_name, cell_range, formula)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"设置公式失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def evaluate_formula(cls, formula: str, context_sheet: Optional[str] = None) -> Dict[str, Any]:
        """
        @intention 计算公式的值，不修改文件
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始计算公式: {formula}")

        try:
            from ..core.excel_writer import ExcelWriter
            writer = ExcelWriter("")  # 临时实例，不需要文件
            result = writer.evaluate_formula(formula, context_sheet)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"公式计算失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def compare_files(cls, file1_path: str, file2_path: str) -> Dict[str, Any]:
        """
        @intention 比较两个Excel文件的所有工作表
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始比较文件: {file1_path} vs {file2_path}")

        try:
            from ..models.types import ComparisonOptions
            from ..core.excel_compare import ExcelComparer

            # 标准文件比较配置
            options = ComparisonOptions(
                compare_values=True,
                compare_formulas=False,
                compare_formats=False,
                ignore_empty_cells=True,
                case_sensitive=True,
                structured_comparison=False
            )

            comparer = ExcelComparer(options)
            result = comparer.compare_files(file1_path, file2_path)
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"比较文件失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def find_last_row(
        cls,
        file_path: str,
        sheet_name: str,
        column: Optional[Union[str, int]] = None
    ) -> Dict[str, Any]:
        """
        @intention 查找表格中最后一行有数据的位置

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            sheet_name: 工作表名称
            column: 指定列来查找最后一行（可选）
                - None: 查找整个工作表的最后一行
                - 整数: 列索引 (1-based，1=A列)
                - 字符串: 列名 (A, B, C...)

        Returns:
            Dict: 包含 success、last_row、message 等信息

        Example:
            # 查找整个工作表的最后一行
            result = ExcelOperations.find_last_row("data.xlsx", "Sheet1")
            # 查找A列的最后一行有数据的位置
            result = ExcelOperations.find_last_row("data.xlsx", "Sheet1", "A")
            # 查找第3列的最后一行有数据的位置
            result = ExcelOperations.find_last_row("data.xlsx", "Sheet1", 3)
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始查找最后一行: {sheet_name}")

        try:
            from ..core.excel_reader import ExcelReader
            reader = ExcelReader(file_path)

            # 获取工作簿和工作表
            workbook = reader._get_workbook(read_only=True, data_only=True)
            sheet = reader._get_worksheet(workbook, sheet_name)

            last_row = 0
            search_info = ""

            if column is None:
                # 查找整个工作表的最后一行
                last_row = sheet.max_row
                # 从后往前查找真正有数据的最后一行
                for row_num in range(sheet.max_row, 0, -1):
                    has_data = False
                    for col_num in range(1, sheet.max_column + 1):
                        cell_value = sheet.cell(row=row_num, column=col_num).value
                        if cell_value is not None and str(cell_value).strip():
                            has_data = True
                            break
                    if has_data:
                        last_row = row_num
                        break
                else:
                    last_row = 0  # 整个工作表都没有数据
                search_info = "整个工作表"
            else:
                # 查找指定列的最后一行
                from openpyxl.utils import column_index_from_string, get_column_letter

                # 转换列参数为列索引
                if isinstance(column, str):
                    try:
                        col_index = column_index_from_string(column.upper())
                    except ValueError:
                        reader.close()
                        return cls._format_error_result(f"无效的列名: {column}")
                elif isinstance(column, int):
                    if column < 1:
                        reader.close()
                        return cls._format_error_result("列索引必须大于等于1")
                    col_index = column
                else:
                    reader.close()
                    return cls._format_error_result("列参数必须是字符串或整数")

                # 查找指定列的最后一行有数据
                for row_num in range(sheet.max_row, 0, -1):
                    cell_value = sheet.cell(row=row_num, column=col_index).value
                    if cell_value is not None and str(cell_value).strip():
                        last_row = row_num
                        break

                col_letter = get_column_letter(col_index)
                search_info = f"{col_letter}列"

            reader.close()

            return {
                'success': True,
                'data': {
                    'last_row': last_row,
                    'sheet_name': sheet_name,
                    'column': column,
                    'search_scope': search_info
                },
                'last_row': last_row,  # 兼容性字段
                'message': f"成功查找{search_info}最后一行: 第{last_row}行" if last_row > 0 else f"{search_info}没有数据"
            }

        except Exception as e:
            error_msg = f"查找最后一行失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def check_duplicate_ids(
        cls,
        file_path: str,
        sheet_name: str,
        id_column: Union[int, str] = 1,
        header_row: int = 1
    ) -> Dict[str, Any]:
        """
        检查Excel工作表中的ID重复情况

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            id_column: ID列位置 (1-based数字或列名)
            header_row: 表头行号 (1-based)

        Returns:
            Dict: 包含success、has_duplicates、duplicate_count、total_ids、unique_ids、duplicates、message
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始检查ID重复: {sheet_name}")

        try:
            from collections import Counter
            from openpyxl import load_workbook

            # 参数验证
            if not file_path or not sheet_name:
                return {
                    'success': False,
                    'message': '文件路径和工作表名不能为空',
                    'has_duplicates': False,
                    'duplicate_count': 0,
                    'total_ids': 0,
                    'unique_ids': 0,
                    'duplicates': []
                }

            # 加载工作簿
            try:
                wb = load_workbook(file_path, read_only=True)
            except FileNotFoundError:
                return {
                    'success': False,
                    'message': f'文件不存在: {file_path}',
                    'has_duplicates': False,
                    'duplicate_count': 0,
                    'total_ids': 0,
                    'unique_ids': 0,
                    'duplicates': []
                }
            except Exception as e:
                return {
                    'success': False,
                    'message': f'无法加载文件: {str(e)}',
                    'has_duplicates': False,
                    'duplicate_count': 0,
                    'total_ids': 0,
                    'unique_ids': 0,
                    'duplicates': []
                }

            # 检查工作表是否存在
            if sheet_name not in wb.sheetnames:
                return {
                    'success': False,
                    'message': f'工作表不存在: {sheet_name}',
                    'has_duplicates': False,
                    'duplicate_count': 0,
                    'total_ids': 0,
                    'unique_ids': 0,
                    'duplicates': []
                }

            ws = wb[sheet_name]

            # 处理列索引
            if isinstance(id_column, str):
                from openpyxl.utils import column_index_from_string
                try:
                    col_idx = column_index_from_string(id_column)
                except Exception:
                    return {
                        'success': False,
                        'message': f'无效的列名: {id_column}',
                        'has_duplicates': False,
                        'duplicate_count': 0,
                        'total_ids': 0,
                        'unique_ids': 0,
                        'duplicates': []
                    }
            else:
                col_idx = id_column

            # 检查表头行是否存在
            if header_row < 1 or header_row > ws.max_row:
                return {
                    'success': False,
                    'message': f'表头行不存在: {header_row}',
                    'has_duplicates': False,
                    'duplicate_count': 0,
                    'total_ids': 0,
                    'unique_ids': 0,
                    'duplicates': []
                }

            # 检查列是否存在
            if col_idx < 1 or col_idx > ws.max_column:
                return {
                    'success': False,
                    'message': f'列不存在或索引超出范围: {col_idx}',
                    'has_duplicates': False,
                    'duplicate_count': 0,
                    'total_ids': 0,
                    'unique_ids': 0,
                    'duplicates': []
                }

            # 收集ID数据
            ids_with_rows = []
            for row in range(header_row + 1, ws.max_row + 1):
                cell_value = ws.cell(row=row, column=col_idx).value
                if cell_value is not None:  # 跳过空值
                    ids_with_rows.append((cell_value, row))

            # 统计ID出现次数
            id_counter = Counter([id_val for id_val, _ in ids_with_rows])
            total_ids = len(ids_with_rows)
            unique_ids = len(id_counter)

            # 查找重复的ID
            duplicates = []
            duplicate_count = 0

            for id_value, count in id_counter.items():
                if count > 1:
                    duplicate_count += 1
                    # 找到该ID的所有行号
                    rows = [row for id_val, row in ids_with_rows if id_val == id_value]
                    # 使用绝对行号（Excel中的实际行号）
                    absolute_rows = rows

                    duplicates.append({
                        'id_value': id_value,
                        'count': count,
                        'rows': absolute_rows
                    })

            has_duplicates = duplicate_count > 0

            # 构建返回结果
            message = f"共检查{total_ids}个ID，发现{duplicate_count}个重复ID" if has_duplicates else f"共检查{total_ids}个ID，无重复ID"

            return {
                'success': True,
                'has_duplicates': has_duplicates,
                'duplicate_count': duplicate_count,
                'total_ids': total_ids,
                'unique_ids': unique_ids,
                'duplicates': duplicates,
                'message': message
            }

        except Exception as e:
            error_msg = f"检查ID重复时发生错误: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)
