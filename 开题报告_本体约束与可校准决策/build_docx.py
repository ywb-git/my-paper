#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按范本版式生成开题报告 docx。

范本版式（从 开题报告-唐宁俊V3.docx 实测）：
  封面：北 京 科 技 大 学 楷体22pt粗体居中；硕士学位研究生 / 选题报告及文献综述 楷体36pt粗体居中
        论文题目 15pt；指导教师等 15pt 左对齐，首行缩进 1333500 EMU
  目录：黑体15pt粗体居中
  Heading1：中文黑体15pt，编号 Times New Roman 15pt
  Heading2：黑体14pt
  Heading3：黑体
  正文：首行缩进 304800 EMU（约2字符）
页面：A4（7560310 x 10692130 EMU），左右边距 1143000，上下边距 914400
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

MD = Path("开题报告_正文.md")
OUT = Path("开题报告_V1.docx")

TITLE = "正例稀缺条件下材料合成路线可行性的概率校准策略研究"


def set_run(run, cn_font=None, en_font="Times New Roman", size=None, bold=None):
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if en_font:
        run.font.name = en_font
        run._element.rPr.rFonts.set(qn("w:eastAsia"), cn_font or en_font)
    elif cn_font:
        run.font.name = cn_font
        run._element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)


def add_para(doc, text, *, cn="宋体", en="Times New Roman", size=12, bold=False,
             align=None, indent=None, space_before=0, space_after=0, line=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    if indent is not None:
        pf.first_line_indent = Emu(indent)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if line is not None:
        pf.line_spacing = line
    r = p.add_run(text)
    set_run(r, cn_font=cn, en_font=en, size=size, bold=bold)
    return p


def add_rich_para(doc, text):
    """解析 **粗体** 与 [n] 上标引用，生成段落。"""
    import re as _re
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Emu(304800)
    pf.space_before = Pt(3)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.5
    token = _re.compile(r"\*\*(.+?)\*\*|\[([0-9,\-]+)\]")
    pos = 0
    for m in token.finditer(text):
        if m.start() > pos:
            r = p.add_run(text[pos:m.start()])
            set_run(r, cn_font="宋体", size=12)
        if m.group(1) is not None:
            r = p.add_run(m.group(1))
            set_run(r, cn_font="宋体", size=12, bold=True)
        else:
            r = p.add_run("[" + m.group(2) + "]")
            set_run(r, cn_font="宋体", size=12)
            r.font.superscript = True
        pos = m.end()
    if pos < len(text):
        r = p.add_run(text[pos:])
        set_run(r, cn_font="宋体", size=12)
    return p


def build():
    doc = Document()

    # ---- 页面设置 ----
    sec = doc.sections[0]
    sec.page_width = Emu(7560310)
    sec.page_height = Emu(10692130)
    sec.left_margin = Emu(1143000)
    sec.right_margin = Emu(1143000)
    sec.top_margin = Emu(914400)
    sec.bottom_margin = Emu(914400)

    # ---- 默认样式 ----
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    from docx.shared import RGBColor
    for name, sz in (("Heading 1", 15), ("Heading 2", 14), ("Heading 3", 12)):
        s = doc.styles[name]
        s.font.name = "黑体"
        s.font.size = Pt(sz)
        s.font.bold = True
        s.font.color.rgb = RGBColor(0, 0, 0)
        s.element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")

    # ---- 封面 ----
    for _ in range(2):
        doc.add_paragraph()
    add_para(doc, "北 京 科 技 大 学", cn="楷体", en="楷体", size=22, bold=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    add_para(doc, "硕士学位研究生", cn="楷体", en="楷体", size=36, bold=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_para(doc, "选题报告及文献综述", cn="楷体", en="楷体", size=36, bold=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)

    for _ in range(2):
        doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("论文题目：" + TITLE)
    set_run(r, cn_font="宋体", size=13)

    for _ in range(3):
        doc.add_paragraph()

    info = [
        "指导教师：       艾冬梅",
        "单    位：       数理学院",
        "姓    名：       唐宁俊",
        "学    号：       M202410824",
        "专业名称：       统计学",
        "入学时间：       2024年9月5日",
    ]
    for line in info:
        add_para(doc, line, size=15, align=WD_ALIGN_PARAGRAPH.LEFT,
                 indent=1333500, space_after=8)

    for _ in range(3):
        doc.add_paragraph()
    add_para(doc, "2025年12月", size=15, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=18)

    doc.add_page_break()

    # ---- 目录 ----
    add_para(doc, "目    录", cn="黑体", en="黑体", size=15, bold=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)
    toc = [
        (0, "1 选题背景及选题意义"),
        (1, "1.1 选题背景"),
        (1, "1.2 选题意义"),
        (0, "2 文献综述"),
        (1, "2.1 材料合成路线推荐方法综述"),
        (1, "2.2 大模型不确定性校准方法综述"),
        (1, "2.3 现有研究的局限性"),
        (0, "3 研究内容、研究方案、时间安排"),
        (1, "3.1 研究内容"),
        (2, "3.1.0 预测目标的界定与标签口径"),
        (2, "3.1.1 基于本体约束的合成可行性决策空间构造"),
        (2, "3.1.2 决策模型的领域知识注入与自动训练"),
        (2, "3.1.3 四类概率校准策略的比较与证据可靠性分层"),
        (1, "3.2 研究方案"),
        (1, "3.3 时间安排"),
        (0, "4 研究的局限性"),
        (0, "参考文献"),
    ]
    for lvl, text in toc:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Emu(304800 * lvl)
        pf.space_after = Pt(4)
        r = p.add_run(text)
        set_run(r, cn_font="宋体", size=12)

    doc.add_page_break()

    # ---- 正文：解析 md ----
    lines = MD.read_text(encoding="utf-8").splitlines()
    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue
        # "# 题目" 已放在封面，正文跳过，避免重复
        if line.startswith("# ") and not line.startswith("## "):
            continue
        # ## -> H1, ### -> H2, #### -> H3
        if line.startswith("#### "):
            doc.add_heading(line[5:].strip(), level=3)
        elif line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=2)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=1)
        elif line.strip() == "{{FIG31}}":
            ip = doc.add_paragraph()
            ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ip.paragraph_format.space_before = Pt(6)
            run = ip.add_run()
            run.add_picture("图3-1_技术路线图.png", width=Cm(15.2))
            cap = doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.space_after = Pt(8)
            cr = cap.add_run("图 3-1　技术路线图")
            set_run(cr, cn_font="宋体", size=10.5)
        elif line.startswith("[") and "]" in line and len(line) > 30:
            # 参考文献条目：悬挂缩进
            pp = doc.add_paragraph()
            pf = pp.paragraph_format
            pf.left_indent = Emu(304800 * 2)
            pf.first_line_indent = Emu(-304800 * 2)
            pf.space_before = Pt(2)
            pf.space_after = Pt(2)
            pf.line_spacing = 1.25
            r = pp.add_run(line.strip())
            set_run(r, cn_font="宋体", size=10.5)
        else:
            add_rich_para(doc, line.strip())

    doc.save(OUT)
    print(f"[ok] {OUT.resolve()}")


if __name__ == "__main__":
    build()
