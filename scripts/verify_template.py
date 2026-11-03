# -*- coding: utf-8 -*-
"""verify.py 模板 —— **只读**校验（不修改任何文件）。

检查 docx / xlsx 的结构、版式、标题序列、表格、**正文数值与 _data.py 复算是否一致**。

用法
----
    python verify.py report.docx                      # 只查报告
    python verify.py report.docx data.xlsx            # 报告 + 工作簿
    python verify.py report.docx data.xlsx --save     # 同时写入 verify_output.txt

约定
----
    · 与本文件同目录放一份 `_data.py`，第 ⑥ 节会自动做数值复算；没有则跳过
    · 第 ④ 节按「一、…八、」识别一级标题，改了章节名要同步 EXP_H1
    · 某节没有对应内容就整节删掉，不要留空检查
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

PASS = FAIL = 0
LINES: list[str] = []

# 按你的报告骨架填写（用于"编号连续无跳号"与"内容未丢失"两项检查）
EXP_H1 = ["一、实验目的", "二、实验原理", "三、实验仪器与试剂",
          "四、实验方法", "五、数据记录、计算与判定", "六、讨论与思考题"]


def value_checks(D) -> list[tuple[str, str]]:
    """★ 第 ⑥ 节要核对的数值：返回 [(名称, 期望字符串), ...]。

    每一项都会在正文里查找该字符串，找不到即 FAIL。按项目补充，例如：

        return [
            ("含量",  f"{D.content()['mean']:.2f}%"),
            ("线性r", f"{D.methodology()['linear']['r']:.4f}"),
            ("限度",  f"{D.CONTENT_LOWER:.1f}%~{D.CONTENT_UPPER:.1f}%"),
        ]
    """
    return []


def verdict_checks(D, text: str) -> list[tuple[str, bool]]:
    """★ 第 ⑦ 节判定一致性：返回 [(名称, 条件), ...]，text 是报告全文。例如：

        return [
            ("判定文字与 pass 一致",
             ("符合规定" in text) == bool(D.content()["pass"])),
            ("限度区间已写明",
             f"{D.CONTENT_LOWER:.1f}%~{D.CONTENT_UPPER:.1f}%" in text),
        ]
    """
    return []


def check(name: str, cond: bool, detail: str = "") -> None:
    global PASS, FAIL
    line = f"  {'PASS' if cond else 'FAIL'}  {name}" + (f"  → {detail}" if detail and not cond else "")
    LINES.append(line)
    print(line)
    if cond:
        PASS += 1
    else:
        FAIL += 1


def out(s: str = "") -> None:
    LINES.append(s)
    print(s)


# ================================================================== 读 docx
def read_docx(path: Path):
    z = zipfile.ZipFile(path)
    for part in z.namelist():                       # ① 结构：全部 XML 可解析
        if part.endswith(".xml"):
            ET.fromstring(z.read(part).decode("utf-8"))
    root = ET.fromstring(z.read("word/document.xml").decode("utf-8"))
    paras = root.findall(f".//{W}p")
    texts = ["".join(n.text or "" for n in p.iter(f"{W}t")) for p in paras]
    body = root.find(f".//{W}body")
    return z, root, paras, texts, body


def rpr_of(p):
    """段落标记或首个 run 的 rPr。"""
    for r in p.findall(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            return rpr
    ppr = p.find(f"{W}pPr")
    return ppr.find(f"{W}rPr") if ppr is not None else None


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    save = "--save" in sys.argv
    if not args:
        print("用法：python verify.py <report.docx> [data.xlsx] [--save]")
        return 2

    docx = Path(args[0])
    xlsx = Path(args[1]) if len(args) > 1 else None
    if not docx.exists():
        print(f"文件不存在：{docx}")
        return 2

    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import _data as D          # noqa: F401  —— 有就做复算，没有就跳过
    except Exception:              # noqa: BLE001
        D = None

    out("=" * 78)
    out(f"只读校验：{docx.name}")
    out("=" * 78)

    z, root, paras, texts, body = read_docx(docx)
    joined = "\n".join(texts)

    # ① 结构 ------------------------------------------------------
    out("\n[1] 结构与可解析性")
    check("docx 内全部 XML 可解析", True)
    check("正文段落数 > 0", len(paras) > 0, f"实际 {len(paras)}")

    # ② 页面 ------------------------------------------------------
    out("\n[2] 页面设置")
    sect = root.find(f".//{W}sectPr")
    pg = sect.find(f"{W}pgSz") if sect is not None else None
    check("纸型为 A4", pg is not None and pg.get(f"{W}w") in ("11906",), 
          f"w={pg.get(f'{W}w') if pg is not None else 'None'}")

    # ③ 默认版式 --------------------------------------------------
    out("\n[3] 默认字体与行距")
    try:
        styles = ET.fromstring(z.read("word/styles.xml").decode("utf-8"))

        def style_sz(style_id: str):
            for s in styles.findall(f"{W}style"):
                if s.get(f"{W}styleId") == style_id:
                    el = s.find(f"{W}rPr/{W}sz")
                    return el.get(f"{W}val") if el is not None else None
            return None

        normal = style_sz("Normal")
        dd = styles.find(f"{W}docDefaults")
        default = None
        if dd is not None:
            el = dd.find(f"{W}rPrDefault/{W}rPr/{W}sz")
            default = el.get(f"{W}val") if el is not None else None
        check("Normal 或 docDefaults 已设置字号", (normal or default) is not None,
              f"Normal={normal}, docDefaults={default}")
        out(f"       Normal 字号 = {normal}（21 = 10.5 pt，24 = 12 pt），"
            f"docDefaults = {default}")
    except Exception as e:                          # noqa: BLE001
        check("默认字号可读", False, str(e))

    # ④ 标题序列 --------------------------------------------------
    out("\n[4] 一级标题序列")
    h1 = [t.strip() for t in texts if t.strip()[:2] in
          ("一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、")]
    check("一级标题数量", len(h1) == len(EXP_H1), f"实际 {len(h1)} / 期望 {len(EXP_H1)}")
    check("一级标题与预期全等（编号连续无跳号）", h1 == EXP_H1, f"{h1}")

    # ⑤ 表格 ------------------------------------------------------
    out("\n[5] 表格")
    tables = root.findall(f".//{W}tbl")
    check("检测到表格" if tables else "无表格（与设计一致）", True, f"共 {len(tables)} 张")
    for i, tbl in enumerate(tables, 1):
        borders = tbl.find(f"{W}tblPr/{W}tblBorders")
        if borders is None:
            check(f"表 {i} 有三线表边框", False, "缺少 tblBorders")
            continue
        vals = {b.tag.replace(W, ""): (b.get(f"{W}val"), b.get(f"{W}sz"))
                for b in borders}
        ok = (vals.get("top", ("", ""))[0] == "single"
              and vals.get("bottom", ("", ""))[0] == "single"
              and vals.get("insideV", ("", ""))[0] == "none"
              and vals.get("left", ("", ""))[0] == "none")
        check(f"表 {i} 为三线表（无竖线、上下有实线）", ok, str(vals))

    # ⑥ 数值复算 ★ ------------------------------------------------
    out("\n[6] 正文数值与 _data.py 复算")
    if D is None:
        out("  （未找到 _data.py，跳过本节）")
    else:
        checks = value_checks(D)
        if not checks:
            out("  （尚未登记核对项：在 value_checks() 里加几行即生效）")
        for name, expect in checks:
            check(f"{name} 与 _data.py 一致：{expect}", expect in joined)

    # ⑦ 判定一致性 ------------------------------------------------
    out("\n[7] 判定与限度")
    if D is None:
        out("  （跳过）")
    else:
        vc = verdict_checks(D, joined)
        if not vc:
            out("  （尚未登记判定检查：在 verdict_checks() 里加几行即生效）")
        for name, cond in vc:
            check(name, bool(cond))

    # ⑧ 内容纪律 --------------------------------------------------
    out("\n[8] 内容纪律")
    check("无「待填」残留（当 DATA_OBS_FILLED=True）",
          not (D is not None and D.DATA_OBS_FILLED and "（待填）" in joined))
    check("正文未出现「None」「nan」等异常字样",
          "None" not in joined and "nan" not in joined.lower())

    # ⑨ 工作簿 ----------------------------------------------------
    out("\n[9] 工作簿")
    if xlsx is None:
        out("  （未提供 xlsx，跳过）")
    elif not xlsx.exists():
        check("xlsx 存在", False, str(xlsx))
    else:
        from openpyxl import load_workbook
        wb_f = load_workbook(xlsx, data_only=False)      # 公式
        wb_v = load_workbook(xlsx, data_only=True)       # 缓存值
        n_formula = n_err = 0
        for ws in wb_f.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.data_type == "f":
                        n_formula += 1
        for ws in wb_v.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("#"):
                        n_err += 1
        check("工作簿可打开", True, f"{len(wb_f.worksheets)} 张表")
        check("公式数 > 0", n_formula > 0, f"实际 {n_formula}")
        check("错误单元格数为 0", n_err == 0, f"实际 {n_err}")

    # ⑩ 收尾 ------------------------------------------------------
    out("\n[10] 文件完整性")
    try:
        from docx import Document
        Document(str(docx))
        check("docx 可被 python-docx 正常打开", True)
    except Exception as e:                          # noqa: BLE001
        check("docx 可被 python-docx 正常打开", False, str(e))

    out()
    out("=" * 78)
    out(f"汇总：PASS {PASS} / FAIL {FAIL}")
    out("=" * 78)

    if save:
        target = Path("verify_output.txt")
        target.write_text("\n".join(LINES), encoding="utf-8")
        print(f"已写入 {target.resolve()}")

    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
