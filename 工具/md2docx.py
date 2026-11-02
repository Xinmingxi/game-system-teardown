
# -*- coding: utf-8 -*-
"""md2docx.py - 把 Markdown 草稿转成排好版的 Word 文档。

用法:
    python 工具/md2docx.py 输入.md 输出.docx
"""
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BT = chr(96)
FENCE = BT * 3
CN_FONT = "微软雅黑"
EN_FONT = "Calibri"
MONO_FONT = "Consolas"
BODY = 10.5
ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
GREY = RGBColor(0x55, 0x55, 0x55)


def _fonts(el, name):
    rPr = el.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), name)


def style_font(style, size, bold=False, color=None):
    style.font.name = EN_FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
    _fonts(style.element, CN_FONT)


def run_font(run, name, size=None, bold=None, color=None):
    run.font.name = name
    _fonts(run._element, name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


INLINE = re.compile(r"(\*\*[^*]+\*\*|" + BT + r"[^" + BT + r"]+" + BT + r")")


def add_inline(p, text, size=BODY, color=None):
    for tok in INLINE.split(text):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            r = p.add_run(tok[2:-2])
            run_font(r, EN_FONT, size=size, bold=True, color=color)
        elif tok.startswith(BT) and tok.endswith(BT) and len(tok) > 2:
            r = p.add_run(tok[1:-1])
            run_font(r, MONO_FONT, size=size - 0.5, color=RGBColor(0xB0, 0x30, 0x30))
        else:
            r = p.add_run(tok)
            run_font(r, EN_FONT, size=size, color=color)


def build_styles(doc):
    st = doc.styles["Normal"]
    style_font(st, BODY)
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.35
    for name, size in (("Heading 1", 20), ("Heading 2", 15), ("Heading 3", 12.5), ("Heading 4", 11)):
        s = doc.styles[name]
        style_font(s, size, bold=True, color=ACCENT)
        s.paragraph_format.space_before = Pt(14 if name == "Heading 1" else 10)
        s.paragraph_format.space_after = Pt(6)
    for name in ("List Bullet", "List Bullet 2", "List Bullet 3", "List Number"):
        try:
            style_font(doc.styles[name], BODY)
        except KeyError:
            pass


def setup_page(doc, title):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.2)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run(title)
    run_font(r, EN_FONT, size=8, color=RGBColor(0x88, 0x88, 0x88))


def cells_of(line):
    return [c.strip() for c in line.strip().strip("|").split("|")]


def is_divider(line):
    s = line.strip()
    return "-" in s and bool(re.fullmatch(r"\|?[\s:\-|]+\|?", s))


def add_table(doc, block):
    rows = [cells_of(b) for b in block]
    rows = [r for i, r in enumerate(rows) if not (i == 1 and is_divider(block[1]))]
    ncol = max(len(r) for r in rows)
    table = doc.add_table(rows=0, cols=ncol)
    table.style = "Table Grid"
    for i, row in enumerate(rows):
        cells = table.add_row().cells
        for j in range(ncol):
            text = row[j] if j < len(row) else ""
            p = cells[j].paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            add_inline(p, text, size=9)
            if i == 0:
                for r in p.runs:
                    r.bold = True
    doc.add_paragraph()


def convert(src, dst):
    with open(src, encoding="utf-8") as f:
        raw = f.read()
    title = ""
    for line in raw.splitlines():
        if line.startswith("# "):
            title = line[2:].strip()
            break
    doc = Document()
    doc.core_properties.title = title
    build_styles(doc)
    setup_page(doc, title)

    lines = raw.splitlines()
    i = 0
    in_code = False
    while i < len(lines):
        line = lines[i]
        s = line.strip()

        if s.startswith(FENCE):
            in_code = not in_code
            i += 1
            continue
        if in_code:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.left_indent = Cm(0.6)
            r = p.add_run(line.rstrip())
            run_font(r, MONO_FONT, size=8.5, color=RGBColor(0x33, 0x33, 0x33))
            i += 1
            continue
        if not s:
            i += 1
            continue

        if s.startswith("|"):
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i])
                i += 1
            if len(block) >= 2:
                add_table(doc, block)
            continue

        if re.fullmatch(r"-{3,}|\*{3,}|_{3,}", s):
            i += 1
            continue

        m = re.match(r"^(#{1,4})\s+(.*)$", s)
        if m:
            lv = len(m.group(1))
            p = doc.add_heading(level=lv)
            add_inline(p, m.group(2), size={1: 20, 2: 15, 3: 12.5, 4: 11}[lv], color=ACCENT)
            i += 1
            continue

        if s.startswith(">"):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.6)
            add_inline(p, s.lstrip("> ").strip(), color=GREY)
            for r in p.runs:
                r.italic = True
            i += 1
            continue

        m = re.match(r"^(\s*)[-*+]\s+(.*)$", line)
        if m:
            indent, text = m.group(1), m.group(2)
            mark = ""
            if text.startswith("[ ] "):
                mark, text = "☐ ", text[4:]
            elif re.match(r"^\[[xX]\] ", text):
                mark, text = "☑ ", text[4:]
            lv = min(len(indent) // 2 + 1, 3)
            p = doc.add_paragraph(style=("List Bullet" if lv == 1 else "List Bullet " + str(lv)))
            add_inline(p, mark + text)
            i += 1
            continue

        m = re.match(r"^(\s*)(\d+)[.)]\s+(.*)$", line)
        if m:
            indent, num, text = m.group(1), m.group(2), m.group(3)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.75 + 0.6 * (len(indent) // 2))
            add_inline(p, num + ". " + text)
            i += 1
            continue

        p = doc.add_paragraph()
        add_inline(p, s)
        i += 1

    doc.save(dst)
    return dst


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: python md2docx.py <input.md> <output.docx>")
        sys.exit(2)
    print("OK ->", convert(sys.argv[1], sys.argv[2]))
