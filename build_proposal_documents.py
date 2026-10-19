"""把开题报告 Markdown 转换为排版好的 Word 文档。

用法：python build_proposal_documents.py
不覆盖已存在的 .docx；若已存在会跳过并提示。
"""
import os
import re
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

BODY_CN = "宋体"
HEAD_CN = "黑体"
LATIN = "Times New Roman"
INK = RGBColor(0x1A, 0x1A, 0x1A)
ACCENT = RGBColor(0x1B, 0x3A, 0x52)


def set_run_font(run, size=10.5, bold=False, color=INK, latin=LATIN, cn=BODY_CN):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = latin
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cn)


def add_mixed_runs(paragraph, text, size=10.5, base_bold=False, color=INK):
    """处理 **粗体**、`代码`、http(s) 链接。"""
    token = re.compile(r"(\*\*.+?\*\*|`[^`]+`|https?://\S+)")
    for piece in token.split(text):
        if not piece:
            continue
        if piece.startswith("**") and piece.endswith("**") and len(piece) > 4:
            run = paragraph.add_run(piece[2:-2])
            set_run_font(run, size=size, bold=True, color=color)
        elif piece.startswith("`") and piece.endswith("`") and len(piece) > 2:
            run = paragraph.add_run(piece[1:-1])
            set_run_font(run, size=size - 0.5, bold=base_bold, color=color, latin="Consolas", cn=BODY_CN)
        elif piece.startswith("http"):
            run = paragraph.add_run(piece)
            set_run_font(run, size=size - 1, bold=False, color=RGBColor(0x1F, 0x4E, 0x79))
        else:
            run = paragraph.add_run(piece)
            set_run_font(run, size=size, bold=base_bold, color=color)


def _extract_date(md_path: str) -> str:
    """从 md 的「报告日期：」行提取日期，取不到则回退。"""
    import re
    try:
        with open(md_path, encoding="utf-8") as f:
            for line in f:
                m = re.match(r"\s*报告日期[：:]\s*(.+)", line)
                if m:
                    return m.group(1).strip()
    except OSError:
        pass
    return "2026 年 10 月 18 日"


def build(md_path: str, docx_path: str, force: bool = False):
    if os.path.exists(docx_path) and not force:
        print(f"[skip] 已存在，不覆盖：{docx_path}")
        return

    with open(md_path, "r", encoding="utf-8") as fh:
        lines = fh.read().splitlines()

    doc = Document()

    # 页面设置：A4
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.8)

    normal = doc.styles["Normal"]
    normal.font.name = LATIN
    normal.font.size = Pt(10.5)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_CN)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    normal.paragraph_format.line_spacing = 1.5

    # 页眉页脚
    header_p = section.header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = header_p.add_run("硕士学位论文开题报告")
    set_run_font(hr, size=8, color=RGBColor(0x80, 0x80, 0x80), cn=HEAD_CN)

    footer_p = section.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer_p.add_run(_extract_date(md_path))
    set_run_font(fr, size=8, color=RGBColor(0x80, 0x80, 0x80), cn=HEAD_CN)

    i = 0
    total = len(lines)
    first_h1 = True
    while i < total:
        raw = lines[i]
        line = raw.strip()

        if not line:
            i += 1
            continue

        # 表格
        if line.startswith("|"):
            block = []
            while i < total and lines[i].strip().startswith("|"):
                block.append(lines[i].strip())
                i += 1
            rows = [
                [c.strip() for c in r.strip("|").split("|")]
                for r in block
                if not re.fullmatch(r"\|[\s:\-|]+\|", r)
            ]
            if rows:
                cols = len(rows[0])
                table = doc.add_table(rows=len(rows), cols=cols)
                table.style = "Table Grid"
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                for ri, row in enumerate(rows):
                    for ci in range(cols):
                        cell = table.cell(ri, ci)
                        cell.text = ""
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_before = Pt(2)
                        p.paragraph_format.space_after = Pt(2)
                        p.paragraph_format.line_spacing = 1.15
                        text = row[ci] if ci < len(row) else ""
                        add_mixed_runs(p, text, size=9, base_bold=(ri == 0))
            doc.add_paragraph()
            continue

        if line == "---":
            i += 1
            continue

        if line.startswith("# "):
            doc.add_paragraph()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(line[2:].strip())
            set_run_font(r, size=20, bold=True, color=ACCENT, cn=HEAD_CN)
            first_h1 = False
            i += 1
            continue

        if line.startswith("## "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(16)
            p.paragraph_format.space_after = Pt(6)
            r = p.add_run(line[3:].strip())
            size = 14 if first_h1 else 15
            set_run_font(r, size=size, bold=True, color=ACCENT, cn=HEAD_CN)
            i += 1
            continue

        if line.startswith("### "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(line[4:].strip())
            set_run_font(r, size=11.5, bold=True, color=ACCENT, cn=HEAD_CN)
            i += 1
            continue

        if line.startswith("- ") or line.startswith("* "):
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.line_spacing = 1.4
            p.paragraph_format.space_after = Pt(2)
            add_mixed_runs(p, line[2:].strip())
            i += 1
            continue

        # 元信息行（学科专业 / 研究方向 / 报告日期 / 工作语言）
        META_PREFIXES = ("学科专业", "研究方向", "报告日期", "工作语言")
        if any(line.startswith(p) for p in META_PREFIXES) and "：" in line:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_mixed_runs(p, line, size=10.5)
            i += 1
            continue

        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0.74)
        p.paragraph_format.space_after = Pt(4)
        add_mixed_runs(p, line)
        i += 1

    doc.save(docx_path)
    print(f"[ok] {docx_path}")


if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    # 顺序与优先级一致（01 首选 → 06），便于按序生成与核对
    targets = [
        ("开题报告01_材料问答中的选择性拒答_2026-10-18.md",
         "开题报告01_材料问答中的选择性拒答_2026-10-18.docx"),
        ("开题报告02_材料知识冲突的识别与消解_2026-10-18.md",
         "开题报告02_材料知识冲突的识别与消解_2026-10-18.docx"),
        ("开题报告03_材料设计的交互式多轮修正_2026-10-18.md",
         "开题报告03_材料设计的交互式多轮修正_2026-10-18.docx"),
        ("开题报告04_材料逆向设计的运行间稳定性_2026-10-18.md",
         "开题报告04_材料逆向设计的运行间稳定性_2026-10-18.docx"),
        ("开题报告05_闭环材料发现中的失败样本利用_2026-10-17.md",
         "开题报告05_闭环材料发现中的失败样本利用_2026-10-17.docx"),
        ("开题报告06_可合成性引导的材料生成_2026-10-18.md",
         "开题报告06_可合成性引导的材料生成_2026-10-18.docx"),
    ]
    for md, dx in targets:
        md_abs = os.path.join(base_dir, md)
        if not os.path.exists(md_abs):
            print(f"[missing] {md}")
            sys.exit(1)
        build(md_abs, os.path.join(base_dir, dx), force=True)
