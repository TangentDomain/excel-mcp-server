"""
P11 回归测试: 表头含空单元格 / NaN / 重复列名时不应崩溃

Bug 背景:
  PK 项目 MapEvent.xlsx 的 MapEvent sheet 第二行表头有 5 个空单元格。
  原代码链路:
    1. _cell_str(float('nan')) 返回 "nan" 字符串 (没处理 NaN)
    2. 5 个空列都被命名为 "nan" → 重复列名
    3. df["nan"] 对重复列名返回 DataFrame 而非 Series
    4. .dtype 访问 DataFrame 报 AttributeError
    5. except Exception 静默吞噬 → 整个 sheet 被跳过 → SQL 查询报"表不存在"

修复:
    - header_analyzer._cell_str: 处理 NaN (float('nan') 不等于自身)
    - advanced_sql_query._load_excel_data: 按位置重建列名 + 去重 + .iloc[:, idx] 防御
    - except 改为 logger.warning 不再静默吞噬
"""

import openpyxl
import pytest

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
from excel_mcp_server_fastmcp.api.header_analyzer import _cell_str


class TestCellStrHandlesNaN:
    """_cell_str 必须把 NaN 当作空值返回 None"""

    def test_float_nan_returns_none(self):
        assert _cell_str(float("nan")) is None

    def test_numpy_nan_returns_none(self):
        np = pytest.importorskip("numpy")
        assert _cell_str(np.nan) is None

    def test_none_returns_none(self):
        assert _cell_str(None) is None

    def test_empty_string_returns_none(self):
        assert _cell_str("") is None
        assert _cell_str("   ") is None

    def test_normal_string_passthrough(self):
        assert _cell_str("EventID") == "EventID"
        assert _cell_str(123) == "123"


@pytest.fixture
def xlsx_with_empty_headers(tmp_path):
    """构造类似 MapEvent 的双行表头 xlsx: 第二行有多个空单元格"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestSheet"
    # 第一行中文描述 (全填)
    ws.append(["事件ID", "事件类型", "路径", "名称", "说明", "备注"])
    # 第二行英文字段名 — 故意让 4-6 列空 (模拟 MapEvent 的 "拉表" 列)
    ws.append(["EventID", "EventType", "", "", "", ""])
    # 数据行
    ws.append([1, "战斗", "路1", "战斗1", "说明1", ""])
    ws.append([2, "QTE", "路2", "QTE测试", "说明2", ""])
    ws.append([3, "宝箱", "路3", "宝箱1", "说明3", ""])
    path = tmp_path / "test_empty_headers.xlsx"
    wb.save(path)
    return str(path)


class TestLoadSheetWithEmptyHeaderCells:
    """表头含空单元格的 sheet 必须能被加载和查询"""

    def test_sheet_loads_without_silent_skip(self, xlsx_with_empty_headers):
        """加载后 TestSheet 必须出现在 sheets 列表里 (未被静默跳过)"""
        engine = AdvancedSQLQueryEngine()
        data = engine._load_excel_data(xlsx_with_empty_headers, "TestSheet")
        assert "TestSheet" in data, "sheet 被静默跳过了 — 回归 bug 复发"
        assert len(data["TestSheet"]) == 3, "应有 3 行数据"

    def test_no_duplicate_column_names(self, xlsx_with_empty_headers):
        """列名必须唯一 — 不允许出现重复 'nan' / ''"""
        engine = AdvancedSQLQueryEngine()
        data = engine._load_excel_data(xlsx_with_empty_headers, "TestSheet")
        cols = list(data["TestSheet"].columns)
        assert len(cols) == len(set(cols)), f"列名重复: {cols}"

    def test_empty_columns_get_chinese_fallback_name(self, xlsx_with_empty_headers):
        """空列名必须回退用第一行中文描述"""
        engine = AdvancedSQLQueryEngine()
        data = engine._load_excel_data(xlsx_with_empty_headers, "TestSheet")
        cols = list(data["TestSheet"].columns)
        assert "EventID" in cols
        assert "EventType" in cols
        col_str = " ".join(str(c) for c in cols)
        assert "路径" in col_str, f"中文描述'路径'未出现在列名: {cols}"
        assert "名称" in col_str
        assert "说明" in col_str
        assert "备注" in col_str

    def test_sql_query_works_on_sheet_with_empty_headers(self, xlsx_with_empty_headers):
        """SQL 能查询这种 sheet (端到端验证)"""
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(
            file_path=xlsx_with_empty_headers,
            sql="SELECT EventID, EventType FROM TestSheet WHERE EventID = 2",
        )
        assert result.get("success"), f"SQL 查询失败: {result.get('message')}"
        data = result.get("data", [])
        # data 格式: [[headers], [row1], ...]
        assert len(data) >= 2, f"应返回表头 + 1 行数据, 实际: {data}"
        row = data[1]
        assert row[0] == 2
        assert row[1] == "QTE"


class TestDuplicateChineseDescriptionsDedup:
    """多列中文描述相同时必须去重"""

    def test_duplicate_chinese_names_get_suffix(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dup"
        # 第一行 3 列中文描述相同
        ws.append(["同名", "同名", "同名"])
        # 第二行全空 (触发回退到中文)
        ws.append(["", "", ""])
        ws.append([1, 2, 3])
        path = str(tmp_path / "dup.xlsx")
        wb.save(path)

        engine = AdvancedSQLQueryEngine()
        data = engine._load_excel_data(path, "Dup")
        cols = list(data["Dup"].columns)
        assert len(cols) == 3
        assert len(set(cols)) == 3, f"列名未去重: {cols}"
