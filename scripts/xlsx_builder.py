# -*- coding: utf-8 -*-
"""xlsx_builder —— 生成实验数据处理工作簿的通用构建器（基于 openpyxl）。

特点
----
1. 四表骨架（使用说明 / 数据输入 / 计算 / 公式与依据）
2. **自动行高**：按「文本显示宽度 ÷ 列宽」估算行数后显式设置 —— Excel 不会自动撑开
   合并单元格，不设就会把长句截断（这是最容易踩的排版坑）
3. **公式误写防护**：以 `=` 开头但含中文的"说明文字"会被拒绝写入，
   避免打开时满屏 `#NAME?`

用法
----
    from xlsx_builder import XlsxBuilder

    x = XlsxBuilder(["使用说明", "数据输入", "计算"])
    x.widths("数据输入", {"A": 10, "B": 16, "C": 14, "D": 40})
    x.title("数据输入", "A1", "吸光度记录", span="A1:D1")
    x.head_row("数据输入", 2, ["序号", "保留时间/min", "峰面积", "备注"])
    x.row("数据输入", 3, [1, 4.671, 558129, ""], fill_input=True)
    x.row("计算", 2, ["含量/%", "=C3/C4*100"])
    x.autofit_rows()
    x.save("out.xlsx")

依赖：openpyxl（`pip install openpyxl`）
"""

from __future__ import annotations

import math
import unicodedata

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.properties import PageSetupProperties

# ----------------------------------------------------------------- 配色
COLOR_TITLE = "1F3864"
COLOR_SECTION = "1F3864"
COLOR_NOTE = "595959"
COLOR_WARN = "C00000"
FILL_HEAD = "EDEDED"
FILL_INPUT = "FFF2CC"      # 手填区（浅黄）
FILL_KEY = "D9E2F3"        # 关键结果（浅蓝）

THIN = Side(style="thin", color="BFBFBF")
BORDER_BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def disp_width(s) -> float:
    """字符串的显示宽度：全角字符按 2、半角按 1 计。"""
    w = 0.0
    for ch in str(s):
        w += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return w


class XlsxBuilder:
    def __init__(self, sheets: list[str], strict: bool = True) -> None:
        """strict=True：遇到「= + 中文」的说明文字直接报错（推荐，避免产出坏文件）。
        strict=False：自动把前导 '=' 换成全角 '＝' 并给出警告。"""
        self.strict = strict
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.sheets: list[str] = []
        for name in sheets:
            ws = self.wb.create_sheet(name)
            self.sheets.append(name)
            self._page(ws)
        self.colw: dict[str, dict[str, float]] = {s: {} for s in sheets}
        self._lines: dict[tuple[str, int], tuple[float, float]] = {}   # (sheet,row)->(行数,字号)

    # ------------------------------------------------------------ 基础
    def _ws(self, sheet: str):
        if sheet not in self.sheets:
            raise KeyError(f"未定义的工作表：{sheet}（已定义：{self.sheets}）")
        return self.wb[sheet]

    @staticmethod
    def _page(ws) -> None:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

    @staticmethod
    def _looks_like_formula(v: str) -> bool:
        """区分真公式与「说明文字」。

        真公式：`=` 后紧跟函数名/引用/括号/负号（如 `=AVERAGE(数据输入!C4:C8)`）；
        说明文字：`=` 后多一个空格再接中文（如 `= 相当量25mg / 50mL`）。
        """
        body = v[1:].lstrip()
        if not body:
            return False
        if body[0].isascii() and (body[0].isalpha() or body[0] in "(+-$"):
            return True          # =A1、=SUM(...)、=-(...)、=$B$2
        if "!" in v and "(" in v:
            return True          # 中文表名引用（较少见但合法）
        return False

    def _guard(self, value):
        """拦截「以 = 开头、实为说明文字」的内容（会被当公式 → #NAME?）。"""
        if isinstance(value, str) and value.startswith("="):
            has_cjk = any(unicodedata.east_asian_width(c) in ("W", "F") for c in value)
            if has_cjk and not self._looks_like_formula(value):
                if self.strict:
                    raise ValueError(
                        f"单元格内容以 '=' 开头但不像公式（应为说明文字）：{value!r}\n"
                        f"请去掉前导 '='；或建 XlsxBuilder(..., strict=False) 让它自动改成全角 '＝'。"
                    )
                import warnings
                warnings.warn(f"说明文字以 '=' 开头，已改为全角：{value!r}", stacklevel=3)
                return "＝" + value[1:]
        return value

    def _track(self, sheet: str, row: int, text, width_chars: float, font_pt: float) -> None:
        """记录某行某格需要的显示行数（取最大值）。"""
        if not width_chars:
            return
        lines = max(1.0, math.ceil(disp_width(text) / max(width_chars - 1, 1)))
        old = self._lines.get((sheet, row))
        if old is None or lines > old[0]:
            self._lines[(sheet, row)] = (lines, font_pt)

    def _span_width(self, sheet: str, span: str) -> float:
        min_col, _, max_col, _ = range_boundaries(span)
        total = 0.0
        for c in range(min_col, max_col + 1):
            total += self.colw[sheet].get(get_column_letter(c), 8.43)
        return total

    # ------------------------------------------------------------ 列宽
    def widths(self, sheet: str, spec: dict[str, float]) -> None:
        ws = self._ws(sheet)
        for col, w in spec.items():
            ws.column_dimensions[col].width = w
            self.colw[sheet][col] = float(w)

    # ------------------------------------------------------------ 写入
    def _write(self, sheet: str, addr: str, value, *, size: float = 11, bold: bool = False,
               color: str | None = None, fill: str | None = None, wrap: bool = True,
               align: str = "left", valign: str = "center", border: bool = False,
               fmt: str | None = None) -> None:
        ws = self._ws(sheet)
        cell = ws[addr]
        cell.value = self._guard(value)
        cell.font = Font(name="宋体", size=size, bold=bold,
                         color=color if color else "000000")
        cell.alignment = Alignment(horizontal=align, vertical=valign, wrap_text=wrap)
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)
        if border:
            cell.border = BORDER_BOX
        if fmt:
            cell.number_format = fmt

    def merge(self, sheet: str, span: str) -> None:
        self._ws(sheet).merge_cells(span)

    def title(self, sheet: str, addr: str, text: str, span: str | None = None,
              size: float = 14) -> None:
        if span:
            self.merge(sheet, span)
        self._write(sheet, addr, text, size=size, bold=True, color=COLOR_TITLE,
                    align="center")
        row = int("".join(ch for ch in addr.split(":")[0] if ch.isdigit()))
        w = self._span_width(sheet, span) if span else 8.43
        self._track(sheet, row, text, w, size)

    def section(self, sheet: str, addr: str, text: str, span: str | None = None,
                size: float = 11) -> None:
        if span:
            self.merge(sheet, span)
        self._write(sheet, addr, text, size=size, bold=True, color=COLOR_SECTION)
        row = int("".join(ch for ch in addr.split(":")[0] if ch.isdigit()))
        w = self._span_width(sheet, span) if span else 8.43
        self._track(sheet, row, text, w, size)

    def note(self, sheet: str, addr: str, text: str, span: str | None = None,
             color: str = COLOR_NOTE) -> None:
        if span:
            self.merge(sheet, span)
        self._write(sheet, addr, text, size=9, color=color)
        row = int("".join(ch for ch in addr.split(":")[0] if ch.isdigit()))
        w = self._span_width(sheet, span) if span else 8.43
        self._track(sheet, row, text, w, 9)

    def head_row(self, sheet: str, row: int, values: list[str], start_col: int = 1) -> None:
        for i, v in enumerate(values):
            col = get_column_letter(start_col + i)
            self._write(sheet, f"{col}{row}", v, size=10, bold=True, fill=FILL_HEAD,
                        align="center", border=True)
            self._track(sheet, row, v, self.colw[sheet].get(col, 8.43), 10)

    def row(self, sheet: str, row: int, values: list, *, start_col: int = 1,
            fill_input: bool = False, fill_key: bool = False, bold: bool = False,
            num_fmts: dict[int, str] | None = None, align: str = "center") -> None:
        """写一行数据。num_fmts 的键是列序号（从 1 起）。"""
        for i, v in enumerate(values):
            col = get_column_letter(start_col + i)
            fill = FILL_INPUT if fill_input else (FILL_KEY if fill_key else None)
            fmt = (num_fmts or {}).get(i + 1)
            # 数值不加边框、文本加边框，视觉上更像表格
            is_num = isinstance(v, (int, float)) and not isinstance(v, bool)
            self._write(sheet, f"{col}{row}", v, size=10, bold=bold, fill=fill,
                        align=align, border=not is_num, fmt=fmt)
            self._track(sheet, row, v, self.colw[sheet].get(col, 8.43), 10)

    def blank(self, sheet: str, row: int, n_cols: int = 4) -> None:
        self._ws(sheet).row_dimensions[row].height = 6

    # ------------------------------------------------------------ 行高
    def autofit_rows(self, min_h: float = 16.0) -> None:
        """按估算行数设置行高（★ 必调，否则长文本被截断）。"""
        for (sheet, row), (lines, pt) in self._lines.items():
            h = max(min_h, lines * (pt * 1.42) + 4)
            self._ws(sheet).row_dimensions[row].height = h

    # ------------------------------------------------------------ 保存
    def save(self, path: str) -> str:
        self.autofit_rows()
        self.wb.save(path)
        return path
