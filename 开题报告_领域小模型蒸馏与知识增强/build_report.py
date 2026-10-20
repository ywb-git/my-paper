#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 Markdown 生成符合北京科技大学开题报告格式的 docx。

格式基准取自同目录范本：开题报告-唐宁俊V3.docx / 开题报告-李壮壮V1.3.docx。
- 页面 A4，页边距 上下 2.54cm、左右 3.17cm
- 封面：学校 22pt 粗体居中；「硕士学位研究生」「选题报告及文献综述」36pt 粗体；
  信息栏 15pt；日期 15pt 居中
- 目录：自动生成域（打开 Word 后按 F9 或提示更新）
- Heading 1 = 15pt，Heading 2 = 14pt，Heading 3 = 14pt（均加粗、黑体）
- 正文 Normal = 12pt，首行缩进 0.847cm，行距 1.3，中文宋体
- 参考文献：List Paragraph + 自动编号，无首行缩进

用法：
    python build_report.py                          # TOC 域版 -> 开题报告_格式版.docx
    python build_report.py --static-toc             # 静态目录版 -> 开题报告_格式版_静态目录.docx
    python build_report.py in.md out.docx           # 指定输入输出
    python build_report.py --force                  # 覆盖已存在的输出文件

两种模式的默认输出文件名不同，不会互相覆盖。若目标文件已存在且未加 --force，
脚本会提示并退出（退出码 2），避免误覆盖。

需要 Python >= 3.10；依赖 python-docx >= 1.0。
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ---------------------------------------------------------------- 常量

# 页面
PAGE_W, PAGE_H = Cm(21.0), Cm(29.7)
MARGIN_TB, MARGIN_LR = Cm(2.54), Cm(3.17)

# 字号
SZ_SCHOOL = Pt(22)
SZ_TITLE_BIG = Pt(36)
SZ_COVER_INFO = Pt(15)
SZ_H1 = Pt(15)
SZ_H2 = Pt(14)
SZ_H3 = Pt(14)
SZ_BODY = Pt(12)

# 字体
FONT_CN_BODY = "宋体"
FONT_CN_HEAD = "黑体"
FONT_EN = "Times New Roman"

# 正文段落格式
BODY_INDENT = Cm(0.847)
BODY_LINE_SPACING = 1.3

# 封面信息栏：标签 -> 占位值（留空为下划线）
COVER_FIELDS = [
    ("指导教师：", ""),
    ("单    位：", ""),
    ("姓    名：", ""),
    ("学    号：", ""),
    ("专业名称：", ""),
    ("入学时间：", ""),
]
COVER_DATE = ""  # 留空由用户填写；也可写 "2026年11月"

# 默认输出文件名（两种模式分开，避免互相覆盖）
DEFAULT_OUT_TOC = "开题报告_格式版.docx"
DEFAULT_OUT_STATIC = "开题报告_格式版_静态目录.docx"

# 图题默认编号（可用 --fig-caption 覆盖）
DEFAULT_FIG_CAPTION = "图 4-1　技术路线图"

# 命令行用法（--help 直接打印此常量，不再从 __doc__ 切割）
USAGE = """用法：
  python build_report.py [选项] [输入.md] [输出.docx]

选项：
  --static-toc          生成静态目录版（默认输出 开题报告_格式版_静态目录.docx）
  --fig-caption TEXT    图题文字（默认 图 4-1　技术路线图）
  --force               覆盖已存在的输出文件
  -h, --help            显示本帮助

两种模式的默认输出文件名不同，不会互相覆盖。若目标文件已存在且未加 --force，
脚本提示后以退出码 2 退出。输出目录必须已存在。

需要 Python >= 3.10；依赖 python-docx >= 1.0。"""


# ---------------------------------------------------------------- 字体工具


def set_run_font(run, *, size=None, bold=None, cn=FONT_CN_BODY, en=FONT_EN,
                 color=None):
    """设置 run 的中英文字体、字号、加粗。"""
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color

    run.font.name = en
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)
    rfonts.set(qn("w:eastAsia"), cn)


def _clear_paragraph(paragraph):
    """清空段落里已有的 run（避免 cell.text="" 留下空 run）。"""
    for r in list(paragraph.runs):
        r._element.getparent().remove(r._element)


def style_set_font(style, *, size=None, bold=None, cn=FONT_CN_BODY, en=FONT_EN):
    """设置样式本身的中英文字体。"""
    if size is not None:
        style.font.size = size
    if bold is not None:
        style.font.bold = bold
    style.font.name = en
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)
    rfonts.set(qn("w:eastAsia"), cn)


# ---------------------------------------------------------------- 域代码


def add_field(paragraph, instr: str, *, placeholder: str = ""):
    """向段落插入一个 Word 域（用于目录、页码）。"""
    run = paragraph.add_run()

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run._r.append(fld_begin)
    run._r.append(instr_el)
    run._r.append(fld_sep)

    if placeholder:
        ph = paragraph.add_run(placeholder)
        set_run_font(ph, size=SZ_BODY, cn=FONT_CN_BODY)

    end_run = paragraph.add_run()
    end_run._r.append(fld_end)
    return paragraph


def add_toc(paragraph, levels: str = "1-3"):
    """插入目录域。"""
    return add_field(paragraph, f' TOC \\o "{levels}" \\h \\z \\u ',
                     placeholder="（请在此处右键选择「更新域」生成目录）")


def add_page_number_footer(section):
    """页脚居中页码。"""
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(p, " PAGE ", placeholder="1")


# ---------------------------------------------------------------- 文档初始化


def init_document() -> Document:
    doc = Document()

    # 页面
    for s in doc.sections:
        s.page_width, s.page_height = PAGE_W, PAGE_H
        s.top_margin = s.bottom_margin = MARGIN_TB
        s.left_margin = s.right_margin = MARGIN_LR

    # Normal
    normal = doc.styles["Normal"]
    style_set_font(normal, size=SZ_BODY, cn=FONT_CN_BODY, en=FONT_EN)
    pf = normal.paragraph_format
    pf.line_spacing = BODY_LINE_SPACING
    pf.first_line_indent = BODY_INDENT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    # 标题样式
    for name, size in (("Heading 1", SZ_H1), ("Heading 2", SZ_H2), ("Heading 3", SZ_H3)):
        st = doc.styles[name]
        style_set_font(st, size=size, bold=True, cn=FONT_CN_HEAD, en=FONT_EN)
        st.font.color.rgb = RGBColor(0, 0, 0)  # 去蓝色
        st.paragraph_format.first_line_indent = Cm(0)
        st.paragraph_format.line_spacing = BODY_LINE_SPACING
        st.paragraph_format.space_before = Pt(12 if name == "Heading 1" else 6)
        st.paragraph_format.space_after = Pt(6)

    # List Paragraph（参考文献用）
    if "List Paragraph" in [s.name for s in doc.styles]:
        lp = doc.styles["List Paragraph"]
        style_set_font(lp, size=SZ_BODY, cn=FONT_CN_BODY, en=FONT_EN)
        lp.paragraph_format.first_line_indent = Cm(0)
        lp.paragraph_format.line_spacing = BODY_LINE_SPACING

    return doc


# ---------------------------------------------------------------- 封面


def build_cover(doc: Document, title: str):
    """按范本生成封面。"""

    def blank(n=1):
        for _ in range(n):
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing = 1.0
            set_run_font(p.add_run(""), size=Pt(12))

    def center_line(text, size, bold=False, space_after=0):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(space_after)
        set_run_font(p.add_run(text), size=size, bold=bold, cn=FONT_CN_HEAD)
        return p

    blank(2)
    center_line("北 京 科 技 大 学", SZ_SCHOOL, bold=True, space_after=18)
    center_line("硕士学位研究生", SZ_TITLE_BIG, bold=True, space_after=6)
    center_line("选题报告及文献综述", SZ_TITLE_BIG, bold=True, space_after=30)
    blank(1)

    # 论文题目：超长时自动缩字号，避免折行
    t_size = SZ_COVER_INFO
    if len(title) > 26:
        t_size = Pt(13)
    if len(title) > 34:
        t_size = Pt(12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(30)
    set_run_font(p.add_run(f"论文题目： {title}"), size=t_size, cn=FONT_CN_BODY)

    blank(1)
    for label, value in COVER_FIELDS:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(4)
        filled = value if value else " " * 12 + "＿" * 8
        set_run_font(p.add_run(f"{label}{filled}"), size=SZ_COVER_INFO, cn=FONT_CN_BODY)

    blank(3)
    center_line(COVER_DATE if COVER_DATE else " " * 6 + "＿" * 8, SZ_COVER_INFO)

    # 分页
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ---------------------------------------------------------------- 目录


def build_toc(doc: Document, static_entries=None):
    """目录。

    static_entries 为 None 时插入 TOC 域（Word 打开后更新域即得真实目录）；
    否则直接写入静态目录条目（用于 PDF 转换，域不会被 LibreOffice 计算）。
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_after = Pt(12)
    set_run_font(p.add_run("目    录"), size=SZ_COVER_INFO, bold=True, cn=FONT_CN_HEAD)

    if static_entries is None:
        tp = doc.add_paragraph()
        tp.paragraph_format.first_line_indent = Cm(0)
        add_toc(tp)
    else:
        for level, text in static_entries:
            tp = doc.add_paragraph()
            tp.paragraph_format.first_line_indent = Cm(0)
            tp.paragraph_format.line_spacing = 1.3
            tp.paragraph_format.left_indent = Cm(0.0)
            tp.paragraph_format.space_after = Pt(2)
            label = f"　" * (level - 1) * 2 + text
            set_run_font(tp.add_run(label), size=Pt(12), cn=FONT_CN_BODY)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ---------------------------------------------------------------- Markdown 解析


RE_HEADING = re.compile(r"^(#{1,4})\s+(.*)$")
RE_ORDERED = re.compile(r"^(\d+)[.、]\s+(.*)$")
RE_UNORDERED = re.compile(r"^[-*+]\s+(.*)$")
RE_TABLE_SEP = re.compile(r"^\|[\s:|-]+\|$")


def parse_markdown(md_text: str):
    """把 Markdown 拆成块列表。

    返回 (blocks, warnings)。warnings 是本函数内产生的解析警告列表，
    不依赖任何模块级状态，可安全地在同一进程内重复调用。
    """
    blocks = []
    warnings: list[str] = []

    def warn(msg: str):
        warnings.append(msg)

    lines = md_text.splitlines()
    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # 标题
        m = RE_HEADING.match(stripped)
        if m:
            level = len(m.group(1))
            blocks.append(("heading", (level, m.group(2).strip())))
            i += 1
            continue

        # 表格
        if stripped.startswith("|") and i + 1 < len(lines) and RE_TABLE_SEP.match(lines[i + 1].strip()):
            header = [c.strip() for c in stripped.strip("|").split("|")]
            rows = []
            j = i + 2
            while j < len(lines) and lines[j].strip().startswith("|"):
                cells = [c.strip() for c in lines[j].strip().strip("|").split("|")]
                if len(cells) == len(header):
                    rows.append(cells)
                else:
                    # 补齐或截断到表头列数，并警告（不再静默丢行）
                    warn(
                        f"表格第 {j + 1} 行列数 {len(cells)} 与表头 {len(header)} 不一致，"
                        f"已自动{'补齐' if len(cells) < len(header) else '截断'}："
                        f"{lines[j].strip()[:60]}"
                    )
                    if len(cells) < len(header):
                        cells = cells + [""] * (len(header) - len(cells))
                    else:
                        cells = cells[: len(header)]
                    rows.append(cells)
                j += 1
            blocks.append(("table", (header, rows)))
            i = j
            continue

        # 图片占位
        if stripped in ("{{FIG31}}", "{{FIG41}}", "{{FIG}}"):
            blocks.append(("figure", stripped))
            i += 1
            continue

        # 未识别的占位符（整行或行内）
        unknown_ph = re.findall(r"\{\{[^}]+\}\}", stripped)
        if unknown_ph:
            for ph in unknown_ph:
                warn(f"第 {i + 1} 行有未识别的占位符 {ph}，将按普通文本输出")

        # 有序列表
        m = RE_ORDERED.match(stripped)
        if m:
            blocks.append(("ordered", m.group(2).strip()))
            i += 1
            continue

        # 无序列表
        m = RE_UNORDERED.match(stripped)
        if m:
            blocks.append(("bullet", m.group(1).strip()))
            i += 1
            continue

        # 普通段落
        blocks.append(("para", stripped))
        i += 1

    return blocks, warnings


# ---------------------------------------------------------------- 富文本写入


RE_INLINE = re.compile(r"(\*\*.+?\*\*|`[^`]+`|\^\{[^}]*\}|\$\$?[^$]+\$\$?)")


def add_rich_text(paragraph, text: str, *, size=SZ_BODY, cn=FONT_CN_BODY, base_bold=False,
                  superscript_refs=True):
    """写入带 **粗体**、`代码`、上标引用 [n] 的文本。

    Markdown 里的 [n] 渲染为上标。
    """
    # 先把 [n] 转成占位，避免和粗体正则冲突
    parts = RE_INLINE.split(text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            inner = part[2:-2]
            _write_plain_with_refs(paragraph, inner, size=size, cn=cn,
                                   bold=True, superscript_refs=superscript_refs)
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            r = paragraph.add_run(part[1:-1])
            set_run_font(r, size=size, bold=base_bold, cn="Consolas")
        elif part.startswith("^"):
            r = paragraph.add_run(part[2:-1] if part.startswith("^{") else part[1:])
            r.font.superscript = True
            set_run_font(r, size=size, bold=base_bold, cn=cn)
        elif part.startswith("$"):
            # 数学公式原样输出（不解析）
            r = paragraph.add_run(part)
            set_run_font(r, size=size, bold=base_bold, cn=cn)
        else:
            _write_plain_with_refs(paragraph, part, size=size, cn=cn,
                                   bold=base_bold, superscript_refs=superscript_refs)


RE_REF = re.compile(r"(\[\d+\](?:\[\d+\])*)")


def _write_plain_with_refs(paragraph, text, *, size, cn, bold, superscript_refs):
    if not superscript_refs:
        r = paragraph.add_run(text)
        set_run_font(r, size=size, bold=bold, cn=cn)
        return
    for seg in RE_REF.split(text):
        if not seg:
            continue
        if RE_REF.fullmatch(seg):
            r = paragraph.add_run(seg)
            r.font.superscript = True
            set_run_font(r, size=size, bold=bold, cn=cn)
        else:
            r = paragraph.add_run(seg)
            set_run_font(r, size=size, bold=bold, cn=cn)


# ---------------------------------------------------------------- 正文渲染


def render_blocks(doc: Document, blocks, *, fig_path: Path | None = None,
                  fig_caption: str = DEFAULT_FIG_CAPTION):
    for kind, payload in blocks:
        if kind == "heading":
            level, text = payload
            # md 里 # = 文档题目（已在封面处理），## = Heading 1，### = Heading 2，#### = Heading 3
            if level == 1:
                continue
            style = {2: "Heading 1", 3: "Heading 2", 4: "Heading 3"}.get(level)
            if style is None:
                continue
            p = doc.add_paragraph(style=style)
            set_run_font(p.add_run(text), size={2: SZ_H1, 3: SZ_H2, 4: SZ_H3}[level],
                         bold=True, cn=FONT_CN_HEAD)

        elif kind == "para":
            p = doc.add_paragraph()
            add_rich_text(p, payload, size=SZ_BODY, cn=FONT_CN_BODY)

        elif kind == "ordered":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.847)
            add_rich_text(p, payload, size=SZ_BODY, cn=FONT_CN_BODY)

        elif kind == "bullet":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.847)
            r = p.add_run("— ")
            set_run_font(r, size=SZ_BODY, cn=FONT_CN_BODY)
            add_rich_text(p, payload, size=SZ_BODY, cn=FONT_CN_BODY)

        elif kind == "table":
            header, rows = payload
            t = doc.add_table(rows=1, cols=len(header))
            t.style = "Table Grid"
            for idx, h in enumerate(header):
                cp = t.rows[0].cells[idx].paragraphs[0]
                _clear_paragraph(cp)
                cp.paragraph_format.first_line_indent = Cm(0)
                cp.paragraph_format.line_spacing = 1.0
                set_run_font(cp.add_run(h.replace("**", "")), size=Pt(10.5),
                             bold=True, cn=FONT_CN_HEAD)
            for row in rows:
                cells = t.add_row().cells
                for idx, val in enumerate(row):
                    cp = cells[idx].paragraphs[0]
                    _clear_paragraph(cp)
                    cp.paragraph_format.first_line_indent = Cm(0)
                    cp.paragraph_format.line_spacing = 1.0
                    set_run_font(cp.add_run(val.replace("**", "")), size=Pt(10.5),
                                 cn=FONT_CN_BODY)
            # 表后空一行
            sp = doc.add_paragraph()
            sp.paragraph_format.first_line_indent = Cm(0)
            sp.paragraph_format.line_spacing = 1.0
            set_run_font(sp.add_run(""), size=Pt(6))

        elif kind == "figure":
            if fig_path and fig_path.exists():
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                p.add_run().add_picture(str(fig_path), width=Cm(13.5))

                cap = doc.add_paragraph()
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cap.paragraph_format.first_line_indent = Cm(0)
                cap.paragraph_format.space_after = Pt(10)
                set_run_font(cap.add_run(fig_caption), size=Pt(10.5), cn=FONT_CN_BODY)


# ---------------------------------------------------------------- 参考文献


def render_references(doc: Document, ref_lines: list[str]):
    """参考文献：自动编号列表。"""
    if not ref_lines:
        return
    for idx, line in enumerate(ref_lines, 1):
        text = re.sub(r"^\[\d+\]\s*", "", line).strip()
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(0.847)
        p.paragraph_format.hanging_indent = Cm(0.847)
        p.paragraph_format.line_spacing = BODY_LINE_SPACING
        set_run_font(p.add_run(f"[{idx}] {text}"), size=Pt(10.5), cn=FONT_CN_BODY)


# ---------------------------------------------------------------- 主流程


def split_markdown(md_text: str):
    """拆分正文与参考文献。返回 (title, body_md, ref_lines)。"""
    lines = md_text.splitlines()

    title = "（未设置题目）"
    for ln in lines:
        m = re.match(r"^#\s+(.*)$", ln.strip())
        if m:
            title = m.group(1).strip()
            break

    # 参考文献起点
    ref_start = None
    for i, ln in enumerate(lines):
        if re.match(r"^#{1,3}\s*参考文献\s*$", ln.strip()):
            ref_start = i
            break

    if ref_start is None:
        body_lines = lines
        ref_lines = []
    else:
        body_lines = lines[:ref_start]
        tail = lines[ref_start + 1:]
        ref_lines = [l.strip() for l in tail if re.match(r"^\[\d+\]", l.strip())]

    return title, "\n".join(body_lines), ref_lines


def build(md_path: Path, out_path: Path, fig_path: Path | None = None,
          static_toc: bool = False, fig_caption: str = DEFAULT_FIG_CAPTION,
          md_text: str | None = None):
    """从 Markdown 生成 docx，写入 out_path。

    参数：
        md_path      输入 Markdown 路径（须为 UTF-8）
        out_path     输出 docx 路径（父目录须已存在）
        fig_path     技术路线图路径；None 或不存在时不插图
        static_toc   True 写静态目录条目；False 写 TOC 域（Word 更新后生成）
        fig_caption  图题文字
        md_text      已读取的 Markdown 文本；给出时不再重复读文件

    返回 (n_refs, n_heads, warnings)：
        n_refs    写入的参考文献条数
        n_heads   渲染的标题个数
        warnings  解析期警告列表；本函数不依赖模块级状态，
                  可在同一进程内重复调用而不互相污染
    """
    if md_text is None:
        md_text = md_path.read_text(encoding="utf-8")
    title, body_md, ref_lines = split_markdown(md_text)

    blocks, warnings = parse_markdown(body_md)

    # 静态目录条目（heading level 2/3/4 -> 目录 level 1/2/3）
    entries = None
    if static_toc:
        entries = []
        for kind, payload in blocks:
            if kind == "heading":
                lv, txt = payload
                if lv in (2, 3, 4):
                    entries.append((lv - 1, txt))
        entries.append((1, "参考文献"))

    doc = init_document()
    build_cover(doc, title)
    build_toc(doc, static_entries=entries)
    render_blocks(doc, blocks, fig_path=fig_path, fig_caption=fig_caption)

    # 参考文献标题
    hp = doc.add_paragraph(style="Heading 1")
    set_run_font(hp.add_run("参考文献"), size=SZ_H1, bold=True, cn=FONT_CN_HEAD)
    render_references(doc, ref_lines)

    add_page_number_footer(doc.sections[0])

    # 清理末尾空段落，避免生成空白尾页
    body = doc.element.body
    for p_el in reversed(list(body.findall(qn("w:p")))):
        texts = p_el.findall(".//" + qn("w:t"))
        has_text = any((t.text or "").strip() for t in texts)
        has_break = p_el.findall(".//" + qn("w:br"))
        has_pic = p_el.findall(".//" + qn("w:drawing"))
        if has_text or has_break or has_pic:
            break
        body.remove(p_el)

    # 让 Word 打开时提示更新域（目录、页码）
    settings = doc.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")

    # 文档属性（避免 author 默认为 python-docx）
    cp = doc.core_properties
    cp.author = ""
    cp.last_modified_by = ""
    cp.title = title
    cp.comments = "由 build_report.py 从 Markdown 生成"

    try:
        doc.save(str(out_path))
    except PermissionError:
        raise PermissionError(f"无法写入输出文件（权限不足）：{out_path}") from None
    n_heads = len([b for b in blocks if b[0] == "heading"])
    return len(ref_lines), n_heads, warnings


def main(argv):
    here = Path(__file__).resolve().parent

    # ---- 参数解析（支持 --key value 形式的 --fig-caption）----
    positional = []
    flags = set()
    fig_caption = DEFAULT_FIG_CAPTION
    args = argv[1:]
    j = 0
    while j < len(args):
        a = args[j]
        if a == "--fig-caption":
            if j + 1 >= len(args):
                print("--fig-caption 缺少取值", file=sys.stderr)
                sys.exit(2)
            val = args[j + 1]
            # K1: 取值不得是另一个选项，否则会把 --force 等吞掉
            if val.startswith("-"):
                print(f"--fig-caption 的取值不能以 - 开头：{val!r}", file=sys.stderr)
                print("如需以 - 开头，请把图题写在配置常量 DEFAULT_FIG_CAPTION 里。",
                      file=sys.stderr)
                sys.exit(2)
            # K3: 不接受空串或全空白
            if not val.strip():
                print("--fig-caption 的取值不能为空", file=sys.stderr)
                sys.exit(2)
            # K2: 重复传入时告警（后者生效）
            if fig_caption != DEFAULT_FIG_CAPTION:
                print(f"[警告] --fig-caption 重复传入，后者覆盖前者：{fig_caption!r} -> {val!r}",
                      file=sys.stderr)
            fig_caption = val
            j += 2
            continue
        if a.startswith("--") or a in ("-h",):
            flags.add(a)
        else:
            positional.append(a)
        j += 1

    known = {"--static-toc", "--force", "--help", "-h"}
    unknown = flags - known
    if unknown or flags & {"--help", "-h"}:
        if unknown:
            print(f"未知参数：{', '.join(sorted(unknown))}", file=sys.stderr)
            sys.exit(2)
        print(USAGE)
        sys.exit(0)

    static = "--static-toc" in flags
    force = "--force" in flags

    if len(positional) > 2:
        print(f"参数过多：只接受「输入 md」与「输出 docx」两个位置参数，"
              f"实际收到 {len(positional)} 个", file=sys.stderr)
        sys.exit(2)

    md = Path(positional[0]) if len(positional) >= 1 else here / "开题报告_正文.md"

    # 两种模式的默认输出名分开，避免互相覆盖
    if len(positional) >= 2:
        out = Path(positional[1])
    else:
        out = here / (DEFAULT_OUT_STATIC if static else DEFAULT_OUT_TOC)

    # ---- 输入检查 ----
    if not md.exists():
        print(f"找不到输入文件：{md}", file=sys.stderr)
        sys.exit(1)

    # L2: 输入必须是文件而非目录
    if md.is_dir():
        print(f"输入路径是目录，不是文件：{md}", file=sys.stderr)
        sys.exit(1)

    # L1+O1: 全文读一次（不再只读 4096 字节），编码/权限问题在此暴露；
    # 读到的内容随后直接传给 build，避免重复读取与 TOCTOU。
    # 不用 os.access 预检：它与真正的 open 之间存在时间窗，
    # 且在 Windows 上只检查只读属性、不检查 ACL。
    try:
        md_text = md.read_text(encoding="utf-8")
    except UnicodeDecodeError as e:
        print(f"输入文件不是 UTF-8 编码：{md}", file=sys.stderr)
        print(f"  位置 {e.start}：{e.reason}", file=sys.stderr)
        print("请转换为 UTF-8 后重试（如 iconv -f GBK -t UTF-8 文件.md）。", file=sys.stderr)
        sys.exit(1)
    except PermissionError:
        print(f"输入文件不可读（权限不足）：{md}", file=sys.stderr)
        sys.exit(1)
    except OSError as e:
        print(f"读取输入文件失败：{md}", file=sys.stderr)
        print(f"  {e}", file=sys.stderr)
        sys.exit(1)

    # ---- G2: 输入输出同路径会销毁输入，必须拒绝 ----
    try:
        if out.exists() and md.exists() and out.resolve() == md.resolve():
            print(f"输入与输出是同一个文件：{md}", file=sys.stderr)
            print("这会销毁输入 md，已中止。请指定不同的输出路径。", file=sys.stderr)
            sys.exit(2)
    except OSError:
        pass

    # ---- G1: 输出目录必须存在（不自动创建，避免拼错路径静默建目录）----
    out_parent = out.parent if out.parent != Path("") else Path(".")
    if not out_parent.exists():
        print(f"输出目录不存在：{out_parent}", file=sys.stderr)
        print("请先创建该目录，或改用一个已存在的输出路径。", file=sys.stderr)
        sys.exit(1)
    if not out_parent.is_dir():
        print(f"输出路径的父级不是目录：{out_parent}", file=sys.stderr)
        sys.exit(1)

    # ---- P1: 输出不能是已存在的目录 ----
    if out.is_dir():
        print(f"输出路径是目录，不是文件：{out}", file=sys.stderr)
        print("请指定一个 .docx 文件路径。", file=sys.stderr)
        sys.exit(1)

    # ---- 覆盖前确认 ----
    if out.exists() and not force:
        print(f"输出文件已存在：{out}", file=sys.stderr)
        print("如确认覆盖，请加 --force 参数重跑。", file=sys.stderr)
        sys.exit(2)

    # ---- P2/P3: 输出可写性 ----
    # 目标已存在则测目标本身，否则测父目录；用实际 open 而非 os.access，
    # 以避开 TOCTOU 与 Windows 上 R_OK 只查只读属性的差异。
    probe_target = out if out.exists() else out_parent / f".__probe_{os.getpid()}"
    try:
        with open(probe_target, "ab"):
            pass
    except PermissionError:
        print(f"输出位置不可写（权限不足）：{probe_target}", file=sys.stderr)
        sys.exit(1)
    except OSError as e:
        print(f"输出位置不可写：{probe_target}", file=sys.stderr)
        print(f"  {e}", file=sys.stderr)
        sys.exit(1)
    else:
        if probe_target != out:
            try:
                probe_target.unlink()
            except OSError:
                pass

    # ---- 找图 ----
    fig = None
    for cand in ("技术路线.png", "图4-1_技术路线图.png", "图3-1_技术路线图.png"):
        cand_path = here / cand
        if cand_path.exists():
            fig = cand_path
            break

    # ---- 构建 ----
    n_refs, n_heads, warnings = build(
        md, out, fig_path=fig, static_toc=static, fig_caption=fig_caption,
        md_text=md_text,
    )

    # ---- 结果汇总 ----
    print(f"[ok] {out}")
    print(f"     模式：{'静态目录' if static else 'TOC 域'}")
    print(f"     标题数 {n_heads}，参考文献 {n_refs} 条")
    if fig:
        print(f"     已嵌入图：{fig.name}")
        print(f"     图题：{fig_caption}")
    else:
        print("     ！未找到技术路线图，图片位置留空")

    if warnings:
        print(f"\n[警告] 共 {len(warnings)} 条：")
        for w in warnings:
            print(f"   - {w}")
    else:
        print("     无解析警告")


if __name__ == "__main__":
    main(sys.argv)
