# 搜索工具模块

from typing import Any

from ..api.excel_operations import ExcelOperations


def register_search_tools(mcp) -> None:
    """注册搜索相关工具"""

    @mcp.tool()
    def excel_search(
        file_path: str,
        pattern: str,
        sheet_name: str | None = None,
        case_sensitive: bool = False,
        whole_word: bool = False,
        use_regex: bool = False,
        include_values: bool = True,
        include_formulas: bool = False,
        range: str | None = None,
    ) -> dict[str, Any]:
        """文本搜索 - 返回单元格位置信息

        Args:
            file_path: Excel文件路径
            pattern: 搜索关键词
            sheet_name: 工作表名(可选)
            case_sensitive: 是否区分大小写
            whole_word: 全词匹配
            use_regex: 正则表达式搜索
            include_values: 搜索单元格值
            include_formulas: 搜索公式内容
            range: 搜索范围

        Returns:
            {success, matches, total}
        """
        return ExcelOperations.search(
            file_path,
            pattern,
            sheet_name,
            case_sensitive,
            whole_word,
            use_regex,
            include_values,
            include_formulas,
            range,
        )

    @mcp.tool()
    def excel_search_directory(
        directory_path: str,
        pattern: str,
        case_sensitive: bool = False,
        whole_word: bool = False,
        use_regex: bool = False,
        include_values: bool = True,
        include_formulas: bool = False,
        recursive: bool = True,
        file_extensions: list[str] | None = None,
        file_pattern: str | None = None,
        max_files: int = 100,
    ) -> dict[str, Any]:
        """在目录下所有Excel文件中搜索内容

        Args:
            directory_path: 目录路径
            pattern: 搜索模式
            case_sensitive: 大小写敏感
            whole_word: 全词匹配
            use_regex: 正则表达式
            include_values: 搜索值
            include_formulas: 搜索公式
            recursive: 递归搜索
            file_extensions: 文件扩展名过滤
            file_pattern: 文件名模式
            max_files: 最大文件数

        Returns:
            {success, matches, total_matches, searched_files}
        """
        return ExcelOperations.search_directory(
            directory_path,
            pattern,
            case_sensitive,
            whole_word,
            use_regex,
            include_values,
            include_formulas,
            recursive,
            file_extensions,
            file_pattern,
            max_files,
        )
