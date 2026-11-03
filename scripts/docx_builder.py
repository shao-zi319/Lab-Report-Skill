# -*- coding: utf-8 -*-
"""docx_builder —— 生成中文实验报告的通用构建器（基于 python-docx）。

只做一件事：把「标题 / 正文 / 三线表 / 注 / 图」按固定版式写进 docx。
数值一律由调用方从 _data.py 传入，本模块不涉及任何业务计算。

用法
----
    from docx_builder import DocxBuilder

    b = DocxBuilder()                       # 默认紧凑型
    b.cover("××片的含量测定", "实验报告")
    b.h1("一、实验目的")
    b.p("1. 掌握……")
    b.h2("（一）原始数据记录")
    b.caption("表 1　吸光度记录")
    b.table(["溶液", "A(λ₂)", "A(λ₁)", "ΔA"],
            [["对照品", "0.648", "0.002", "0.646"]],
            widths=[4, 2, 2, 2])
    b.note("① 划改处取改后值。")
    b.save("out.docx")

版式
----
    DocxBuilder(layout="compact")   紧凑型：10.5 pt / 行距 1.28 / 边距 2.0×1.5 cm
    DocxBuilder(layout="loose")     宽松型：12 pt / 行距 1.5 / 边距 2.54×3.17 cm

依赖：python-docx（`pip install python-docx`）
"""

from __future__ import annotations

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ----------------------------------------------------------------- 版式预设
LAYOUTS = {
    "compact": {
        "body_pt": 10.5, "line": 1.28, "indent_pt": 21.0,
        "margin_lr_cm": 1.5, "margin_tb_cm": 2.0,
        "h1_pt": 12.5, "h2_pt": 11.0, "h3_pt": 10.5,
        "cap_pt": 9.5, "note_pt": 9.5, "table_pt": 9.0,
        "h1_color": "1F3864", "note_color": "444444",
        "title_pt": 18.0, "subtitle_pt": 14.0,
    },
    "loose": {
        "body_pt": 12.0, "line": 1.5, "indent_pt": 24.0,
        "margin_lr_cm": 3.17, "margin_tb_cm": 2.54,
        "h1_pt": 14.0, "h2_pt": 12.0, "h3_pt": 12.0,
        "cap_pt": 10.5, "note_pt": 9.5, "table_pt": 10.5,
        "h1_color": "000000", "note_color": "444444",
        "title_pt": 16.0, "subtitle_pt": 14.0,
    },
}

SONG, HEI, KAI, LATIN = "宋体", "黑体", "楷体", "Times New Roman"


class DocxBuilder:
    def __init__(self, layout: str = "compact") -> None:
        self.cfg = LAYOUTS[layout]
        self.doc = Document()
        self._setup_page()
        self._setup_default_style()
        self._table_no = 0
        self._fig_no = 0

    # ------------------------------------------------------------ 基础设施
    def _setup_page(self) -> None:
        s = self.doc.sections[0]
        s.page_width, s.page_height = Cm(21.0), Cm(29.7)
        s.left_margin = s.right_margin = Cm(self.cfg["margin_lr_cm"])
        s.top_margin = s.bottom_margin = Cm(self.cfg["margin_tb_cm"])
        self.avail_cm = 21.0 - 2 * self.cfg["margin_lr_cm"]

    def _setup_default_style(self) -> None:
        # Normal 样式（所有段落的基底）
        st = self.doc.styles["Normal"]
        st.font.name = LATIN
        st.font.size = Pt(self.cfg["body_pt"])
        st.element.rPr.rFonts.set(qn("w:eastAsia"), SONG)
        pf = st.paragraph_format
        pf.line_spacing = self.cfg["line"]
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        # docDefaults（未显式指定字号时的兜底，校验会查这里）
        try:
            rpr = self.doc.styles.element.find(qn("w:docDefaults")).find(
                qn("w:rPrDefault")).find(qn("w:rPr"))
            sz = rpr.find(qn("w:sz"))
            if sz is None:
                sz = OxmlElement("w:sz")
                rpr.append(sz)
            sz.set(qn("w:val"), str(int(round(self.cfg["body_pt"] * 2))))
            szcs = rpr.find(qn("w:szCs"))
            if szcs is not None:
                szcs.set(qn("w:val"), str(int(round(self.cfg["body_pt"] * 2))))
            fonts = rpr.find(qn("w:rFonts"))
            if fonts is not None:
                fonts.set(qn("w:eastAsia"), SONG)
                fonts.set(qn("w:ascii"), LATIN)
                fonts.set(qn("w:hAnsi"), LATIN)
        except AttributeError:
            pass          # 该模板没有 docDefaults 时忽略

    # ------------------------------------------------------------ 段落
    def _para(self, text: str, *, size: float | None = None, bold: bool = False,
              ea: str = SONG, color: str | None = None, indent: bool = True,
              align: str = "justify", before: float = 0, after: float = 0,
              line: float | None = None, keep: bool = False) -> None:
        p = self.doc.add_paragraph()
        fmt = p.paragraph_format
        fmt.alignment = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
                         "left": WD_ALIGN_PARAGRAPH.LEFT,
                         "center": WD_ALIGN_PARAGRAPH.CENTER}[align]
        fmt.space_before = Pt(before)
        fmt.space_after = Pt(after)
        fmt.line_spacing = line if line is not None else self.cfg["line"]
        fmt.keep_with_next = keep
        if indent:
            fmt.first_line_indent = Pt(self.cfg["indent_pt"])
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size if size is not None else self.cfg["body_pt"])
        run.font.name = LATIN
        run._element.rPr.rFonts.set(qn("w:eastAsia"), ea)
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
        return

    def p(self, text: str) -> None:
        """正文段落（首行缩进、两端对齐）。"""
        self._para(text)

    def pc(self, text: str, size: float | None = None, bold: bool = False) -> None:
        """居中段落（公式、反应式、简短说明）。"""
        self._para(text, size=size, bold=bold, indent=False, align="center", before=2, after=2)

    def h1(self, text: str) -> None:
        self._para(text, size=self.cfg["h1_pt"], bold=True, ea=HEI,
                   color=self.cfg["h1_color"], indent=False, align="left",
                   before=12, after=6, keep=True)

    def h2(self, text: str) -> None:
        self._para(text, size=self.cfg["h2_pt"], bold=True, ea=HEI,
                   color=self.cfg["h1_color"], indent=False, align="left",
                   before=8, after=4, keep=True)

    def h3(self, text: str) -> None:
        self._para(text, size=self.cfg["h3_pt"], bold=True, ea=KAI,
                   indent=False, align="left", before=5, after=3, keep=True)

    def caption(self, text: str) -> None:
        """表题 / 图题：居中加粗，与下一段（表格/图）同页。"""
        self._para(text, size=self.cfg["cap_pt"], bold=True, ea=SONG,
                   indent=False, align="center", before=8, after=3, keep=True)

    def note(self, text: str) -> None:
        """表注：楷体灰字，左缩进。"""
        p = self.doc.add_paragraph()
        fmt = p.paragraph_format
        fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        fmt.left_indent = Pt(12)
        fmt.space_before = Pt(2)
        fmt.space_after = Pt(6)
        fmt.line_spacing = 1.15
        run = p.add_run(text)
        run.font.size = Pt(self.cfg["note_pt"])
        run.font.name = LATIN
        run._element.rPr.rFonts.set(qn("w:eastAsia"), KAI)
        run.font.color.rgb = RGBColor.from_string(self.cfg["note_color"])

    def page_break(self) -> None:
        self.doc.add_page_break()

    # ------------------------------------------------------------ 封面
    def cover(self, title: str, subtitle: str = "实验报告",
              lines: list[str] | None = None) -> None:
        self._para(title, size=self.cfg["title_pt"], bold=True, ea=HEI,
                   color=self.cfg["h1_color"], indent=False, align="center", after=4)
        self._para(subtitle, size=self.cfg["subtitle_pt"], bold=True, ea=HEI,
                   color=self.cfg["h1_color"], indent=False, align="center", after=12)
        for ln in (lines or [
            "组号：第 ____ 组　　同组人：____________　　日期：______ 年 ___ 月 ___ 日",
            "温度：______ ℃　　湿度：______ %　　仪器编号：____________",
        ]):
            self._para(ln, size=self.cfg["body_pt"] - 0.5, indent=False,
                       align="center", after=2)

    # ------------------------------------------------------------ 三线表
    def table(self, headers: list[str], rows: list[list], widths: list[float] | None = None):
        """三线表：顶线 1.5 pt、表头下 0.75 pt、底线 1.5 pt，无竖线无内部横线。

        widths 为相对列宽（如 [4, 2, 2, 2]），按可用版心宽度等比换算。
        """
        self._table_no += 1
        n_rows, n_cols = len(rows) + 1, len(headers)
        t = self.doc.add_table(rows=n_rows, cols=n_cols)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        t.allow_autofit = False

        widths = widths or [1] * n_cols
        total = sum(widths)
        col_cm = [self.avail_cm * w / total for w in widths]

        self._tbl_borders(t)
        self._tbl_layout_fixed(t, col_cm)

        data = [headers] + [[str(c) for c in row] for row in rows]
        for ri, row in enumerate(data):
            # 表头行重复（跨页时）
            if ri == 0:
                trPr = t.rows[0]._tr.get_or_add_trPr()
                el = OxmlElement("w:tblHeader")
                el.set(qn("w:val"), "true")
                trPr.append(el)
            for ci, val in enumerate(row):
                cell = t.cell(ri, ci)
                cell.width = Cm(col_cm[ci])
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                pg = cell.paragraphs[0]
                pg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pg.paragraph_format.line_spacing = 1.0
                pg.paragraph_format.space_before = Pt(1.5)
                pg.paragraph_format.space_after = Pt(1.5)
                run = pg.add_run(val)
                run.bold = ri == 0
                run.font.size = Pt(self.cfg["table_pt"])
                run.font.name = LATIN
                run._element.rPr.rFonts.set(qn("w:eastAsia"), SONG)
                if ri == 0:
                    self._cell_bottom_border(cell, sz=6)
        return t

    @staticmethod
    def _tbl_borders(t) -> None:
        tblPr = t._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge, val, sz in (("top", "single", 12), ("left", "none", 0),
                              ("bottom", "single", 12), ("right", "none", 0),
                              ("insideH", "none", 0), ("insideV", "none", 0)):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), val)
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
            borders.append(el)
        tblPr.append(borders)

    @staticmethod
    def _tbl_layout_fixed(t, col_cm: list[float]) -> None:
        tblPr = t._tbl.tblPr
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tblPr.append(layout)
        grid = t._tbl.tblGrid
        for gc, w in zip(grid.findall(qn("w:gridCol")), col_cm):
            gc.set(qn("w:w"), str(int(w * 567)))     # 1 cm ≈ 567 twips

    @staticmethod
    def _cell_bottom_border(cell, sz: int) -> None:
        tcPr = cell._tc.get_or_add_tcPr()
        borders = OxmlElement("w:tcBorders")
        el = OxmlElement("w:bottom")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
        tcPr.append(borders)

    # ------------------------------------------------------------ 插图
    def picture(self, path: str, width_cm: float = 12.0, caption: str | None = None) -> None:
        """插入图片；caption 为空时自动编号为「图 N」。"""
        self._fig_no += 1
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.add_run().add_picture(path, width=Cm(width_cm))
        self._para(caption or f"图 {self._fig_no}", size=self.cfg["cap_pt"], bold=True,
                   indent=False, align="center", before=3, after=8)

    # ------------------------------------------------------------ 保存
    def save(self, path: str) -> str:
        self.doc.save(path)
        return path

    @property
    def table_count(self) -> int:
        return self._table_no
