# -*- coding: utf-8 -*-
"""
把 PROJECT_OVERVIEW 的 Markdown 转换成排版好的 Word 文档。
只做结构化渲染：标题 / 表格 / 列表 / 引用提示 / 行内标记。

用法：
  python build_overview_docx.py
  python build_overview_docx.py --src a.md --dst b.docx --font 9.5 --spacing 1.05 \
      --margin 2.0 --footer "PROJECT_OVERVIEW 精简版"
"""
import argparse
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

_ap = argparse.ArgumentParser()
_ap.add_argument("--src", default=r"D:\dshworkarea\丹寨苗族蜡染\PROJECT_OVERVIEW.md")
_ap.add_argument("--dst", default=r"D:\dshworkarea\丹寨苗族蜡染\PROJECT_OVERVIEW_v0.2.docx")
_ap.add_argument("--font", type=float, default=10.5, help="正文字号 pt")
_ap.add_argument("--spacing", type=float, default=1.25, help="行距倍数")
_ap.add_argument("--margin", type=float, default=2.4, help="左右边距 cm")
_ap.add_argument("--footer", default="PROJECT_OVERVIEW v0.2")
_args = _ap.parse_args()

SRC = _args.src
DST = _args.dst
BODY_PT = _args.font
LINE_SPACING = _args.spacing
SIDE_MARGIN = _args.margin
FOOTER_TEXT = _args.footer

EA_BODY = "宋体"
EA_HEAD = "黑体"
EA_MONO = "Consolas"
LATIN = "Times New Roman"

TEXT_WIDTH_CM = 21.0 - 2 * SIDE_MARGIN

COMPACT = BODY_PT < 10.5
TITLE_PT = 15 if COMPACT else 20
H2_PT = 11.5 if COMPACT else 13.5
H3_PT = 10.0 if COMPACT else 11.5
H_SPACE_BEFORE = 5 if COMPACT else 10
H_SPACE_AFTER = 2 if COMPACT else 4
QUOTE_PT = 8.5 if COMPACT else 9.5  # A4 宽 - 左右边距


# ---------- 底层工具 ----------
def set_run(run, ea=EA_BODY, latin=LATIN, size=None, bold=False,
            italic=False, color=None):
    run.font.name = latin
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), ea)


def add_bottom_border(paragraph, sz=6, color="BFBFBF"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def shade_cell(cell, hex_fill="EDEDED"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)


def add_page_field(paragraph, size=9):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)
    set_run(run, size=size)


# ---------- 行内渲染 ----------
INLINE_RE = re.compile(r"(\*\*.+?\*\*|`[^`]+`|【[^】]*】|\[Verified\]|\[Preliminary\]|\[TBD\])")

COLOR_TBD = RGBColor(0x9E, 0x9E, 0x9E)
COLOR_VERIFIED = RGBColor(0x1F, 0x6F, 0x3C)
COLOR_PRELIM = RGBColor(0xB5, 0x6C, 0x00)


def fill_runs(paragraph, text, size=None, ea=EA_BODY, bold_all=False):
    """把一段 markdown 行内文本写进段落，返回段落。"""
    text = text.replace("`", "")
    for part in INLINE_RE.split(text):
        if not part:
            continue
        bold = bold_all
        color = None
        ea_font = ea
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            part = part[2:-2]
            bold = True
        elif part.startswith("【") and part.endswith("】"):
            color = COLOR_TBD
            ea_font = "楷体"
        elif part == "[Verified]":
            color = COLOR_VERIFIED
            bold = True
        elif part == "[Preliminary]":
            color = COLOR_PRELIM
            bold = True
        elif part == "[TBD]":
            color = COLOR_TBD
            bold = True
        run = paragraph.add_run(part)
        set_run(run, ea=ea_font, size=(BODY_PT if size is None else size), bold=bold, color=color)
    return paragraph


# ---------- 表格 ----------
def cjk_len(s):
    return sum(2 if "\u3000" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef" else 1 for ch in s)


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def set_grid(table, widths_cm):
    """按给定宽度重写 tblGrid，保证固定布局下列宽真的生效。"""
    grid = table._tbl.tblGrid
    for gc in list(grid.findall(W_NS + "gridCol")):
        grid.remove(gc)
    for w in widths_cm:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(Cm(w).twips)))
        grid.append(gc)


def add_table(doc, rows):
    header, body = rows[0], rows[1:]
    ncol = max(len(r) for r in rows)
    table = doc.add_table(rows=0, cols=ncol)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False  # python-docx 会写入 w:tblLayout fixed，勿手工追加

    # 依据内容长度分配列宽（长文本列给更多空间；列多时抬高下限，避免窄列一字一行）
    floor = 10 if ncol >= 6 else 8
    weights = []
    for c in range(ncol):
        longest = max((cjk_len(r[c]) if c < len(r) else 0) for r in rows)
        weights.append(min(max(longest, floor), 32))
    total = sum(weights)
    widths_cm = [TEXT_WIDTH_CM * w / total for w in weights]
    set_grid(table, widths_cm)
    widths = [Cm(w) for w in widths_cm]

    # 列数多时缩小字号，保证不挤成一字一行
    cell_pt = 7.5 if ncol >= 6 else (8.0 if ncol == 5 else 8.5)

    def write_row(cells, data, bold=False, fill=None):
        for i, cell in enumerate(cells):
            txt = data[i] if i < len(data) else ""
            cell.width = widths[i]
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.line_spacing = 1.0
            fill_runs(p, txt, size=cell_pt, bold_all=bold)
            if fill:
                shade_cell(cell, fill)

    hcells = table.add_row().cells
    for i, cell in enumerate(hcells):
        cell.width = widths[i]
    write_row(hcells, header, bold=True, fill="E8EEF4")
    repeat_header(table.rows[0])

    for r in body:
        cells = table.add_row().cells
        for i, cell in enumerate(cells):
            cell.width = widths[i]
        write_row(cells, r)

    # 表后留白
    gap = doc.add_paragraph()
    gap.paragraph_format.space_after = Pt(4)
    gap.paragraph_format.space_before = Pt(0)
    for run in gap.runs:
        run.font.size = Pt(4)
    return table


# ---------- 主流程 ----------
def main():
    with open(SRC, encoding="utf-8") as f:
        lines = f.read().split("\n")

    doc = Document()

    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(SIDE_MARGIN)
    sec.right_margin = Cm(SIDE_MARGIN)
    sec.top_margin = Cm(2.0 if SIDE_MARGIN < 2.4 else 2.2)
    sec.bottom_margin = Cm(2.0 if SIDE_MARGIN < 2.4 else 2.2)

    normal = doc.styles["Normal"]
    normal.font.name = LATIN
    normal.font.size = Pt(BODY_PT)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), EA_BODY)
    normal.paragraph_format.line_spacing = LINE_SPACING
    normal.paragraph_format.space_after = Pt(3 if BODY_PT >= 10.5 else 2)

    for lvl, size in ((1, TITLE_PT), (2, H2_PT), (3, H3_PT)):
        st = doc.styles["Heading %d" % lvl]
        st.font.name = LATIN
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0x1F, 0x30, 0x4E)
        if st.element.rPr is None:
            st.element.get_or_add_rPr()
        rf = st.element.rPr.rFonts
        if rf is None:
            rf = OxmlElement("w:rFonts")
            st.element.rPr.insert(0, rf)
        rf.set(qn("w:eastAsia"), EA_HEAD)

    # 页脚：文件名 + 页码
    footer_p = sec.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    grey = RGBColor(0x88, 0x88, 0x88)
    set_run(footer_p.add_run(FOOTER_TEXT + "　·　第 "), size=9, color=grey)
    add_page_field(footer_p, size=9)
    set_run(footer_p.add_run(" 页"), size=9, color=grey)

    i = 0
    n = len(lines)
    while i < n:
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # 分隔线
        if re.fullmatch(r"-{3,}", stripped):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(6)
            add_bottom_border(p)
            i += 1
            continue

        # 表格块
        if stripped.startswith("|"):
            block = []
            while i < n and lines[i].strip().startswith("|"):
                block.append(lines[i].strip())
                i += 1
            rows = []
            for k, row in enumerate(block):
                if k == 1 and re.fullmatch(r"\|[\s\-:|\u3000]+\|", row):
                    continue
                cells = [c.strip() for c in row.strip("|").split("|")]
                rows.append(cells)
            if rows:
                add_table(doc, rows)
            continue

        # 标题
        m = re.match(r"^(#{1,4})\s+(.*)$", stripped)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            if level == 1:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(4 if COMPACT else 10)
                fill_runs(p, text, size=TITLE_PT, ea=EA_HEAD, bold_all=True)
            else:
                h = doc.add_heading(level=min(level - 1, 3))
                h.paragraph_format.space_before = Pt(H_SPACE_BEFORE if level == 2 else (H_SPACE_BEFORE - 2))
                h.paragraph_format.space_after = Pt(H_SPACE_AFTER)
                fill_runs(h, text, size={2: H2_PT, 3: H3_PT}.get(level, H3_PT),
                          ea=EA_HEAD, bold_all=True)
            i += 1
            continue

        # 引用提示
        if stripped.startswith(">"):
            text = stripped.lstrip(">").strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.6)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)
            for run in fill_runs(p, text, size=QUOTE_PT, ea="楷体").runs:
                run.italic = True
                run.font.color.rgb = RGBColor(0x55, 0x5F, 0x6D)
            i += 1
            continue

        # 复选框
        m = re.match(r"^-\s+\[( |x)\]\s+(.*)$", stripped)
        if m:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.left_indent = Cm(0.9)
            p.paragraph_format.first_line_indent = Cm(-0.4)
            box = "☑ " if m.group(1) == "x" else "☐ "
            set_run(p.add_run(box), size=11)
            fill_runs(p, m.group(2))
            i += 1
            continue

        # 列表项（含缩进层级）
        m = re.match(r"^(\s*)-\s+(.*)$", raw)
        if m:
            indent = len(m.group(1))
            style = "List Bullet 2" if indent >= 2 else "List Bullet"
            p = doc.add_paragraph(style=style)
            p.paragraph_format.left_indent = Cm(0.75 + (0.6 if indent >= 2 else 0))
            p.paragraph_format.first_line_indent = Cm(-0.35)
            p.paragraph_format.space_after = Pt(2)
            fill_runs(p, m.group(2))
            i += 1
            continue

        # 有序列表：保留字面编号，避免 Word 自动编号串号
        m = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if m:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.85)
            p.paragraph_format.first_line_indent = Cm(-0.85)
            p.paragraph_format.space_after = Pt(2)
            set_run(p.add_run(m.group(1) + ". "), size=BODY_PT)
            fill_runs(p, m.group(2))
            i += 1
            continue

        # 普通段落
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        fill_runs(p, stripped)
        i += 1

    doc.save(DST)
    print("saved:", DST)


if __name__ == "__main__":
    main()
