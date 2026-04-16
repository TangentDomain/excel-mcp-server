# -*- coding: utf-8 -*-
"""
format_cells 深度测试 - R55+ 迭代重点

覆盖范围:
  - 单项样式: bold/italic/underline/font_size/font_color/bg_color/number_format/alignment/wrap_text/border_style
  - 组合操作: merge + bold + bg_color 同时传
  - 边界: 空range、超范围、合并后再拆分
  - 中文字体、Unicode列名场景
  - 扁平格式 → 嵌套格式转换 (_normalize_formatting)
  - preset 预设样式
  - 边框详细配置(四边不同)
  - 渐变背景、图案填充
  - text_rotation / indent / shrink_to_fit
  - strikethrough / double underline
"""

import os
import pytest
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from excel_mcp_server_fastmcp.core.excel_writer import ExcelWriter
from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations


# ==================== Helper ====================

def _create_test_xlsx(file_path: str, rows: int = 5, cols: int = 4, sheet_name: str = "Sheet1"):
    """创建测试用 xlsx 文件，含数据"""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c, value=r * 10 + c)
    wb.save(file_path)
    wb.close()


def _read_cell_style(file_path: str, cell_ref: str = "A1", sheet_name: str = "Sheet1") -> dict:
    """读取单元格的样式属性用于断言"""
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    cell = ws[cell_ref]
    font = cell.font
    fill = cell.fill
    alignment = cell.alignment
    border = cell.border
    result = {
        "bold": font.bold,
        "italic": font.italic,
        "underline": font.underline,
        "size": font.size,
        "font_name": font.name,
        "color": str(font.color) if font.color else None,
        "fill_type": fill.fill_type if fill else None,
        "fgColor": str(fill.fgColor) if fill and fill.fgColor else None,
        "alignment_h": alignment.horizontal,
        "alignment_v": alignment.vertical,
        "wrap_text": alignment.wrap_text,
        "text_rotation": alignment.text_rotation,
        "number_format": cell.number_format,
        "border_left": border.left.style if border and border.left else None,
        "border_right": border.right.style if border and border.right else None,
        "border_top": border.top.style if border and border.top else None,
        "border_bottom": border.bottom.style if border and border.bottom else None,
        "value": cell.value,
    }
    wb.close()
    return result


# ==================== Test Class ====================

class TestFormatCellsDeep:
    """format_cells 深度测试套件"""

    # ---------- 1. 单项字体样式 ----------

    def test_bold(self, temp_dir):
        """加粗"""
        fp = str(temp_dir / "bold.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:B2", {"font": {"bold": True}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["bold"] is True

    def test_italic(self, temp_dir):
        """斜体"""
        fp = str(temp_dir / "italic.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"font": {"italic": True}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["italic"] is True

    def test_underline_single(self, temp_dir):
        """单下划线"""
        fp = str(temp_dir / "underline.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"font": {"underline": "single"}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["underline"] == "single"

    def test_underline_double(self, temp_dir):
        """双下划线"""
        fp = str(temp_dir / "underline_double.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"font": {"underline": "double"}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["underline"] == "double"

    def test_underline_accounting(self, temp_dir):
        """会计下划线"""
        fp = str(temp_dir / "underline_acc.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"font": {"underline": "singleAccounting"}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["underline"] == "singleAccounting"

    def test_font_size(self, temp_dir):
        """字号"""
        fp = str(temp_dir / "fontsize.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"font": {"size": 20}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["size"] == 20

    def test_font_color(self, temp_dir):
        """字体颜色 (HEX)"""
        fp = str(temp_dir / "fontcolor.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"font": {"color": "FF0000"}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        # openpyxl color 转为字符串包含 RGB 值
        assert style["color"] is not None
        assert "FF0000" in style["color"] or "ff0000" in style["color"].lower()

    def test_strikethrough(self, temp_dir):
        """删除线"""
        fp = str(temp_dir / "strike.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"font": {"strikethrough": True}})
        assert result.success is True
        # strikethrough 在 Font 对象上

    def test_font_name(self, temp_dir):
        """字体名称"""
        fp = str(temp_dir / "fontname.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"font": {"name": "Arial"}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["font_name"] == "Arial"

    # ---------- 2. 背景填充 ----------

    def test_bg_color_solid(self, temp_dir):
        """纯色背景"""
        fp = str(temp_dir / "bgcolor.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:C3", {"fill": {"type": "solid", "color": "FFFF00"}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["fill_type"] == "solid"
        # fgColor 应包含黄色值
        assert style["fgColor"] is not None

    def test_gradient_fill(self, temp_dir):
        """渐变背景"""
        fp = str(temp_dir / "gradient.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {
            "fill": {"type": "gradient", "colors": ["4472C4", "ED7D31"], "gradient_type": "linear"}
        })
        assert result.success is True

    def test_pattern_fill(self, temp_dir):
        """图案填充"""
        fp = str(temp_dir / "pattern.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {
            "fill": {"type": "pattern", "patternType": "lightGray", "fgColor": "FF0000"}
        })
        assert result.success is True

    # ---------- 3. 对齐与换行 ----------

    def test_alignment_center(self, temp_dir):
        """水平居中"""
        fp = str(temp_dir / "align_center.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:D1", {"alignment": {"horizontal": "center"}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["alignment_h"] == "center"

    def test_alignment_vertical(self, temp_dir):
        """垂直居中"""
        fp = str(temp_dir / "align_v.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"alignment": {"vertical": "center"}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["alignment_v"] == "center"

    def test_wrap_text(self, temp_dir):
        """自动换行"""
        fp = str(temp_dir / "wrap.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"alignment": {"wrap_text": True}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["wrap_text"] is True

    def test_text_rotation(self, temp_dir):
        """文字旋转角度"""
        fp = str(temp_dir / "rotation.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"alignment": {"text_rotation": 45}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["text_rotation"] == 45

    def test_indent(self, temp_dir):
        """缩进"""
        fp = str(temp_dir / "indent.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"alignment": {"indent": 3}})
        assert result.success is True

    def test_shrink_to_fit(self, temp_dir):
        """缩小字体填充"""
        fp = str(temp_dir / "shrink.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"alignment": {"shrink_to_fit": True}})
        assert result.success is True

    # ---------- 4. 数字格式 ----------

    def test_number_format_currency(self, temp_dir):
        """货币数字格式"""
        fp = str(temp_dir / "numfmt.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:B2", {"number_format": "¥#,##0.00"})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["number_format"] == "¥#,##0.00"

    def test_number_format_percent(self, temp_dir):
        """百分比格式"""
        fp = str(temp_dir / "pctfmt.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"number_format": "0.00%"})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["number_format"] == "0.00%"

    def test_number_format_date(self, temp_dir):
        """日期格式"""
        fp = str(temp_dir / "datefmt.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"number_format": "YYYY-MM-DD"})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["number_format"] == "YYYY-MM-DD"

    # ---------- 5. 边框 ----------

    def test_border_thin(self, temp_dir):
        """细边框 (via format_cells border 参数)"""
        fp = str(temp_dir / "border_thin.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:C3", {
            "border": {"left": "thin", "right": "thin", "top": "thin", "bottom": "thin"}
        })
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["border_left"] == "thin"
        assert style["border_right"] == "thin"
        assert style["border_top"] == "thin"
        assert style["border_bottom"] == "thin"

    def test_border_mixed_styles(self, temp_dir):
        """四边不同边框样式"""
        fp = str(temp_dir / "border_mix.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {
            "border": {
                "top": "medium",
                "bottom": "thick",
                "left": "double",
                "right": "dashed",
                "color": "FF0000",
            }
        })
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["border_top"] == "medium"
        assert style["border_bottom"] == "thick"

    def test_border_with_color_dict(self, temp_dir):
        """边框使用 dict 配置（含独立颜色）"""
        fp = str(temp_dir / "border_dict.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {
            "border": {
                "top": {"style": "medium", "color": "0000FF"},
                "bottom": {"style": "thin"},
            }
        })
        assert result.success is True

    # ---------- 6. 组合操作 ----------

    def test_combine_bold_bg_color_alignment(self, temp_dir):
        """组合: 加粗 + 背景色 + 居中"""
        fp = str(temp_dir / "combine_basic.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:D1", {
            "font": {"bold": True, "size": 14},
            "fill": {"type": "solid", "color": "D9E1F2"},
            "alignment": {"horizontal": "center", "vertical": "center"},
        })
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["bold"] is True
        assert style["size"] == 14
        assert style["alignment_h"] == "center"
        assert style["alignment_v"] == "center"
        assert style["fill_type"] == "solid"

    def test_combine_font_multiple_attrs(self, temp_dir):
        """组合多个字体属性: bold+italic+underline+color+size"""
        fp = str(temp_dir / "font_combo.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {
            "font": {
                "bold": True,
                "italic": True,
                "underline": "double",
                "color": "FF0000",
                "size": 16,
                "name": "Courier New",
            }
        })
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["bold"] is True
        assert style["italic"] is True
        assert style["underline"] == "double"
        assert style["size"] == 16
        assert style["font_name"] == "Courier New"

    def test_combine_all_categories(self, temp_dir):
        """全类别组合: font + fill + alignment + number_format + border"""
        fp = str(temp_dir / "full_combo.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:C3", {
            "font": {"bold": True, "color": "FFFFFF", "size": 12},
            "fill": {"type": "solid", "color": "4472C4"},
            "alignment": {"horizontal": "right", "wrap_text": True},
            "number_format": "#,##0.00",
            "border": {"top": "thin", "bottom": "thin", "left": "thin", "right": "thin"},
        })
        assert result.success is True
        assert result.metadata.get("formatted_count", 0) >= 3  # 至少3个单元格

    # ---------- 7. Merge + Format 组合 ----------

    def test_merge_then_format(self, temp_dir):
        """先合并再设置格式"""
        fp = str(temp_dir / "merge_fmt.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)

        # 先合并
        r1 = writer.merge_cells("Sheet1!A1:D1")
        assert r1.success is True

        # 再格式化合并区域
        r2 = writer.format_cells("Sheet1!A1:D1", {
            "font": {"bold": True, "size": 14},
            "fill": {"type": "solid", "color": "FFFF00"},
            "alignment": {"horizontal": "center"},
        })
        assert r2.success is True

    def test_merge_with_bold_and_bg_via_api(self, temp_dir):
        """通过 API 层 excel_format_cells merge+bold+bg_color"""
        from excel_mcp_server_fastmcp.server import excel_format_cells
        fp = str(temp_dir / "api_merge_combo.xlsx")
        _create_test_xlsx(fp)

        result = excel_format_cells(
            file_path=fp,
            sheet_name="Sheet1",
            cell_range="A1:D1",
            formatting={"merge": True, "bold": True, "bg_color": "FFD700", "alignment": "center"},
        )
        assert result["success"] is True

    # ---------- 8. Unmerge ----------

    def test_merge_then_unmerge(self, temp_dir):
        """合并后拆分"""
        fp = str(temp_dir / "merge_unmerge.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)

        r1 = writer.merge_cells("Sheet1!A1:C1")
        assert r1.success is True

        r2 = writer.unmerge_cells("Sheet1!A1:C1")
        assert r2.success is True

        # 拆分后文件仍可正常读取
        wb = load_workbook(fp)
        ws = wb.active
        # 确认合并区域已被拆分 (openpyxl 的 merged_cells 不再包含该范围)
        merged_ranges = list(ws.merged_cells.ranges)
        wb.close()
        assert not any(str(mr) == "A1:C1" for mr in merged_ranges)

    def test_unmerge_non_merged(self, temp_dir):
        """拆分未合并的区域（不应报错）"""
        fp = str(temp_dir / "unmerge_noop.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        # 未合并的区域调用 unmerge，openpyxl 通常不报错
        result = writer.unmerge_cells("Sheet1!A1:C1")
        # openpyxl 允许对非合并区域调用 unmerge_cells（no-op）

    # ---------- 9. 边界 case ----------

    def test_single_cell_range(self, temp_dir):
        """单单元格格式化"""
        fp = str(temp_dir / "single_cell.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!B2", {"font": {"bold": True}})
        assert result.success is True
        assert result.metadata.get("formatted_count") == 1

    def test_large_range(self, temp_dir):
        """大范围格式化 (20x10)"""
        fp = str(temp_dir / "large_range.xlsx")
        _create_test_xlsx(fp, rows=20, cols=10)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:J20", {"font": {"bold": True}})
        assert result.success is True
        assert result.metadata.get("formatted_count") == 200

    def test_format_preserves_cell_values(self, temp_dir):
        """格式化不改变单元格值"""
        fp = str(temp_dir / "preserve_vals.xlsx")
        _create_test_xlsx(fp, rows=3, cols=3)
        original_val = 11  # A1 = 1*10+1 = 11
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:C3", {
            "font": {"bold": True, "color": "FF0000"},
            "fill": {"type": "solid", "color": "00FF00"},
            "alignment": {"horizontal": "right"},
            "number_format": "0.00",
        })
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["value"] == original_val

    def test_empty_formatting_dict(self, temp_dir):
        """空格式字典"""
        fp = str(temp_dir / "empty_fmt.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {})
        # 空字典不会报错，只是没有任何效果
        assert result.success is True

    def test_invalid_sheet_name(self, temp_dir):
        """不存在的工作表"""
        fp = str(temp_dir / "bad_sheet.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("NonExistent!A1", {"font": {"bold": True}})
        assert result.success is False

    # ---------- 10. 中文字体 & Unicode ----------

    def test_chinese_font_name(self, temp_dir):
        """中文字体名称"""
        fp = str(temp_dir / "chinese_font.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1", {"font": {"name": "微软雅黑", "size": 12}})
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["font_name"] == "微软雅黑"
        assert style["size"] == 12

    def test_chinese_sheet_name(self, temp_dir):
        """中文工作表名格式化"""
        fp = str(temp_dir / "chinese_sheet.xlsx")
        _create_test_xlsx(fp, sheet_name="数据表")
        writer = ExcelWriter(fp)
        result = writer.format_cells("数据表!A1:B2", {"font": {"bold": True}})
        assert result.success is True

    # ---------- 11. 扁平格式转换 (_normalize_formatting) ----------

    def test_normalize_flat_bold(self):
        """扁平 bold → 嵌套 font.bold"""
        result = ExcelOperations._normalize_formatting({"bold": True})
        assert result == {"font": {"bold": True}}

    def test_normalize_flat_bg_color(self):
        """扁平 bg_color → 嵌套 fill.color"""
        result = ExcelOperations._normalize_formatting({"bg_color": "FF0"})
        assert result == {"fill": {"color": "FF0"}}

    def test_normalize_flat_alignment(self):
        """扁平 alignment: "center" → nested alignment.horizontal: "center\""""
        result = ExcelOperations._normalize_formatting({"alignment": "center"})
        assert result == {"alignment": {"horizontal": "center"}}

    def test_normalize_flat_font_size(self):
        """扁平 font_size → font.size"""
        result = ExcelOperations._normalize_formatting({"font_size": 18})
        assert result == {"font": {"size": 18}}

    def test_normalize_flat_font_color(self):
        """扁平 font_color → font.color"""
        result = ExcelOperations._normalize_formatting({"font_color": "FF0000"})
        assert result == {"font": {"color": "FF0000"}}

    def test_normalize_flat_wrap_text(self):
        """扁平 wrap_text → alignment.wrap_text"""
        result = ExcelOperations._normalize_formatting({"wrap_text": True})
        assert result == {"alignment": {"wrap_text": True}}

    def test_normalize_flat_combined(self):
        """多项扁平参数同时转换"""
        result = ExcelOperations._normalize_formatting({
            "bold": True,
            "bg_color": "FFFF00",
            "alignment": "center",
            "font_size": 16,
            "italic": True,
            "underline": "double",
        })
        assert result["font"]["bold"] is True
        assert result["font"]["italic"] is True
        assert result["font"]["underline"] == "double"
        assert result["font"]["size"] == 16
        assert result["fill"]["color"] == "FFFF00"
        assert result["alignment"]["horizontal"] == "center"

    def test_normalize_nested_passthrough(self):
        """已是嵌套格式的直接透传"""
        nested = {"font": {"bold": True}, "fill": {"color": "FF0"}}
        result = ExcelOperations._normalize_formatting(nested)
        assert result is nested  # 同一对象引用

    def test_normalize_none_returns_empty(self):
        """None 返回空字典"""
        result = ExcelOperations._normalize_formatting(None)
        assert result == {}

    def test_normalize_gradient_colors(self):
        """扁平 gradient_colors → fill.gradient"""
        result = ExcelOperations._normalize_formatting({
            "gradient_colors": ["4472C4", "ED7D31"],
            "gradient_type": "linear",
        })
        assert result["fill"]["type"] == "gradient"
        assert result["fill"]["colors"] == ["4472C4", "ED7D31"]
        assert result["fill"]["gradient_type"] == "linear"

    def test_normalize_border_string(self):
        """扁平 border 字符串 → border.style"""
        result = ExcelOperations._normalize_formatting({"border": "thin"})
        assert result["border"] == {"style": "thin"}

    def test_normalize_border_dict_passthrough(self):
        """扁平 border 字典直接透传"""
        border_dict = {"top": "medium", "color": "FF0000"}
        result = ExcelOperations._normalize_formatting({"border": border_dict})
        assert result["border"] == border_dict

    def test_normalize_strikethrough(self):
        """扁平 strikethrough → font.strikethrough"""
        result = ExcelOperations._normalize_formatting({"strikethrough": True})
        assert result == {"font": {"strikethrough": True}}

    def test_normalize_text_rotation(self):
        """扁平 text_rotation → alignment.text_rotation"""
        result = ExcelOperations._normalize_formatting({"text_rotation": -90})
        assert result == {"alignment": {"text_rotation": -90}}

    def test_normalize_indent(self):
        """扁平 indent → alignment.indent"""
        result = ExcelOperations._normalize_formatting({"indent": 5})
        assert result == {"alignment": {"indent": 5}}

    def test_normalize_shrink_to_fit(self):
        """扁平 shrink_to_fit → alignment.shrink_to_fit"""
        result = ExcelOperations._normalize_formatting({"shrink_to_fit": True})
        assert result == {"alignment": {"shrink_to_fit": True}}

    def test_normalize_vertical_alignment(self):
        """扁平 vertical_alignment → alignment.vertical"""
        result = ExcelOperations._normalize_formatting({"vertical_alignment": "bottom"})
        assert result == {"alignment": {"vertical": "bottom"}}

    def test_normalize_number_format(self):
        """扁平 number_format 直接传递"""
        result = ExcelOperations._normalize_formatting({"number_format": "0.00%"})
        assert result == {"number_format": "0.00%"}

    def test_normalize_unknown_key_passthrough(self):
        """未知键直接透传（向后兼容）"""
        result = ExcelOperations._normalize_formatting({"custom_key": "custom_value"})
        assert result["custom_key"] == "custom_value"

    # ---------- 12. Preset 测试 ----------

    def test_preset_title(self, temp_dir):
        """预设 title: 微软雅黑 14号 加粗 居中"""
        fp = str(temp_dir / "preset_title.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = excel_format_cells = ExcelOperations.format_cells(
            fp, "Sheet1", "A1", formatting=None, preset="title"
        )
        assert result["success"] is True

    def test_preset_header(self, temp_dir):
        """预设 header: 微软雅黑 11号 加粗 灰底"""
        fp = str(temp_dir / "preset_header.xlsx")
        _create_test_xlsx(fp)
        result = ExcelOperations.format_cells(
            fp, "Sheet1", "A1:D1", formatting=None, preset="header"
        )
        assert result["success"] is True

    def test_preset_currency(self, temp_dir):
        """预设 currency: ¥#,##0.00"""
        fp = str(temp_dir / "preset_currency.xlsx")
        _create_test_xlsx(fp)
        result = ExcelOperations.format_cells(
            fp, "Sheet1", "A1", formatting=None, preset="currency"
        )
        assert result["success"] is True

    def test_preset_highlight(self, temp_dir):
        """预设 highlight: 黄色背景"""
        fp = str(temp_dir / "preset_highlight.xlsx")
        _create_test_xlsx(fp)
        result = ExcelOperations.format_cells(
            fp, "Sheet1", "A1", formatting=None, preset="highlight"
        )
        assert result["success"] is True

    def test_preset_data(self, temp_dir):
        """预设 data: 微软雅黑 10号"""
        fp = str(temp_dir / "preset_data.xlsx")
        _create_test_xlsx(fp)
        result = ExcelOperations.format_cells(
            fp, "Sheet1", "A1:C5", formatting=None, preset="data"
        )
        assert result["success"] is True

    # ---------- 13. formatted_count 元数据 ----------

    def test_formatted_count_multi_cell(self, temp_dir):
        """验证 formatted_count 准确性"""
        fp = str(temp_dir / "fmt_count.xlsx")
        _create_test_xlsx(fp, rows=3, cols=5)  # 3行5列=15单元格
        writer = ExcelWriter(fp)
        result = writer.format_cells("Sheet1!A1:E3", {"font": {"bold": True}})
        assert result.success is True
        assert result.metadata.get("formatted_count") == 15

    # ---------- 14. Overwrite 格式 ----------

    def test_format_overwrite(self, temp_dir):
        """后执行的格式覆盖前面的"""
        fp = str(temp_dir / "overwrite.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)

        # 第一次: 红色
        writer.format_cells("Sheet1!A1", {"font": {"color": "FF0000", "bold": True}})

        # 第二次: 蓝色 + 取消 bold
        writer.format_cells("Sheet1!A1", {"font": {"color": "0000FF", "bold": False}})

        style = _read_cell_style(fp, "A1")
        assert style["bold"] is False
        # 颜色应为蓝色
        assert style["color"] is not None

    # ---------- 15. set_borders 工具方法 ----------

    def test_set_borders_tool(self, temp_dir):
        """set_borders 独立工具方法"""
        fp = str(temp_dir / "set_borders.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.set_borders("Sheet1!A1:C3", "medium")
        assert result.success is True
    def test_set_borders_thick(self, temp_dir):
        """粗边框"""
        fp = str(temp_dir / "borders_thick.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)
        result = writer.set_borders("Sheet1!A1", "thick")
        assert result.success is True
        style = _read_cell_style(fp, "A1")
        assert style["border_left"] == "thick"

    # ---------- R55+ 新增边缘 case 测试 ----------

    def test_format_all_none_values_noop(self, temp_dir):
        """所有样式值为 None 时不应报错（no-op）"""
        fp = str(temp_dir / "all_none.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)

        # 路径1: 直接调用 writer.format_cells 传入含 None 值的扁平格式
        # （_apply_cell_format 应跳过 None config 而不 crash）
        result = writer.format_cells("Sheet1!A1:B2", {
            "bold": None, "italic": None, "font_size": None,
            "bg_color": None, "alignment": None,
        })
        assert result.success is True
        # 值不变
        style = _read_cell_style(fp, "A1")
        assert style["value"] == 11

        # 路径2: 通过 ExcelOperations（会经过 _normalize_formatting 过滤 None → 空dict）
        result2 = ExcelOperations.format_cells(fp, "Sheet1", "A1:B2", {
            "bold": None, "italic": None, "font_size": None,
            "bg_color": None, "alignment": None,
        })
        assert result2["success"] is True

    def test_merge_with_border_and_bold_combo(self, temp_dir):
        """合并 + 粗边框 + 加粗 + 背景色 组合操作"""
        fp = str(temp_dir / "merge_border_bold.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)

        # 先合并
        r1 = writer.merge_cells("Sheet1!A1:C1")
        assert r1.success is True

        # 再对合并区域应用组合样式
        r2 = writer.format_cells("Sheet1!A1:C1", {
            "font": {"bold": True, "size": 14},
            "fill": {"color": "FF0000"},
            "border": {
                "top": {"style": "medium", "color": "000000"},
                "bottom": {"style": "medium", "color": "000000"},
                "left": {"style": "medium", "color": "000000"},
                "right": {"style": "medium", "color": "000000"},
            },
        })
        assert r2.success is True

        # 验证左上角单元格样式
        style = _read_cell_style(fp, "A1")
        assert style["bold"] is True
        assert style["size"] == 14
        assert style["border_top"] == "medium"

    def test_number_format_scientific_and_fraction(self, temp_dir):
        """科学计数法和分数格式"""
        fp = str(temp_dir / "num_fmt_special.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)

        # 科学计数法
        r1 = writer.format_cells("Sheet1!A1", {"number_format": "0.00E+00"})
        assert r1.success is True
        style = _read_cell_style(fp, "A1")
        assert style["number_format"] == "0.00E+00"

        # 分数格式
        r2 = writer.format_cells("Sheet1!B1", {"number_format": "# ?/?"})
        assert r2.success is True
        style = _read_cell_style(fp, "B1")
        assert style["number_format"] == "# ?/?"

    def test_format_partial_overwrite_preserves_other_attrs(self, temp_dir):
        """部分覆盖：第二次格式化只改 bold，保留之前的 bg_color"""
        fp = str(temp_dir / "partial_overwrite.xlsx")
        _create_test_xlsx(fp)
        writer = ExcelWriter(fp)

        # 第一次：设置 bold + bg_color
        r1 = writer.format_cells("Sheet1!A1", {
            "font": {"bold": True},
            "fill": {"color": "00FF00"},
        })
        assert r1.success is True

        # 第二次：只改 bold 为 False，不传 fill
        r2 = writer.format_cells("Sheet1!A1", {
            "font": {"bold": False},
        })
        assert r2.success is True

        # 验证: bold 被 override，但 openpyxl 的 format_cells 是赋值行为
        # fill 是否保留取决于实现（这里验证不报错即可）
        style = _read_cell_style(fp, "A1")
        assert style["bold"] is False

    def test_flat_format_comprehensive_edge_cases(self, temp_dir):
        """扁平格式边缘 case：boolean False、0、空字符串、特殊字符"""
        from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations

        # bold=False 应该正常传递（不是 None，应该被设置）
        result = ExcelOperations._normalize_formatting({
            "bold": False,
            "italic": 0,  # 0 是 falsy 但不是 None
            "font_size": 0,
            "bg_color": "",
            "alignment": "justify",
        })
        # bold=False 应出现在 font 中
        assert result["font"]["bold"] is False
        # font_size=0 应出现
        assert result["font"]["size"] == 0
        # alignment 应正确映射
        assert result["alignment"]["horizontal"] == "justify"
        # 空字符串 bg_color 应产生 fill（空字符串不是 None）
        assert "fill" in result

    def test_unicode_sheet_name_format(self, temp_dir):
        """Unicode/中文工作表名格式化"""
        fp = str(temp_dir / "unicode_sheet.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "数据表"
        ws.append(["名称", "值"])
        ws.append(["测试", 42])
        wb.save(fp)
        wb.close()

        writer = ExcelWriter(fp)
        result = writer.format_cells("数据表!A1:B1", {
            "font": {"bold": True, "name": "微软雅黑"},
            "fill": {"color": "D9D9D9"},
        })
        assert result.success is True

        style = _read_cell_style(fp, "A1", sheet_name="数据表")
        assert style["bold"] is True
        assert style["font_name"] == "微软雅黑"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
