#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""两遍构建交付件：把真实页码写进目录域结果。

为什么需要两遍：
    Word 的目录是「域」，未更新时显示的是缓存结果。若不填缓存，交付件在
    不会自动更新域的阅览器里只显示一句提示文字，与范本（带页码的目录）不一致。
    页码只能由排版结果得到，故先用占位目录构建 → 转 PDF → 定位各标题所在页
    → 再构建一次，把（标题, 页码）写进 TOC 域的缓存结果。域本身保留，
    在 Word/WPS 里仍可按 F9 重算。

用法：
    python build_deliverable.py                     # 默认 开题报告_正文_V4.md -> 开题报告_V4.docx
    python build_deliverable.py in.md out.docx

依赖：soffice（LibreOffice）、pdfinfo、pdftotext（poppler-utils）。
"""

from __future__ import annotations

import re
import subprocess
import sys
import tempfile
from pathlib import Path

import build_report as br

HERE = Path(__file__).resolve().parent
FIG_CANDIDATES = ("技术路线.png", "图4-1_技术路线图.png", "图3-1_技术路线图.png")


def norm(s: str) -> str:
    """去掉所有空白后比较：PDF 里长标题会被折行，直接比对会漏。"""
    return re.sub(r"\s+", "", s or "")


def find_fig():
    for cand in FIG_CANDIDATES:
        p = HERE / cand
        if p.exists():
            return p
    return None


def to_pdf(docx: Path) -> Path:
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(docx.parent), str(docx)],
        check=True, capture_output=True, timeout=300,
    )
    pdf = docx.with_suffix(".pdf")
    if not pdf.exists():
        raise SystemExit(f"PDF 未生成：{pdf}")
    return pdf


def page_texts(pdf: Path) -> list[str]:
    info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout
    m = re.search(r"Pages:\s+(\d+)", info)
    if not m:
        raise SystemExit(f"pdfinfo 解析失败：{info}")
    n = int(m.group(1))
    out = []
    for i in range(1, n + 1):
        txt = subprocess.run(
            ["pdftotext", "-f", str(i), "-l", str(i), str(pdf), "-"],
            capture_output=True, text=True,
        ).stdout
        out.append(norm(txt))
    return out


def collect_heads(md: Path):
    """返回 [(目录层级, 标题文本)]，与 build() 内的口径一致。"""
    _title, body, _refs = br.split_markdown(md.read_text(encoding="utf-8"))
    blocks, _ = br.parse_markdown(body)
    heads = []
    first_para = None
    for kind, payload in blocks:
        if kind == "heading" and payload[0] in (2, 3, 4):
            heads.append((payload[0] - 1, payload[1]))
        elif kind == "para" and first_para is None and heads:
            first_para = payload
    heads.append((1, "参考文献"))
    return heads, first_para


def locate(pages: list[str], heads, anchor: str):
    """定位正文首页与各标题页；返回 (offset, {标题: 显示页})。"""
    key_anchor = norm(anchor)[:24]
    body_start = None
    for i, t in enumerate(pages, 1):
        if key_anchor and key_anchor in t:
            body_start = i
            break
    if body_start is None:
        raise SystemExit(f"未能在 PDF 中定位正文首段（锚点：{key_anchor!r}）")

    # 页码只在正文节内查找，避免命中目录页上的同名条目
    found = {}
    for _lvl, txt in heads:
        key = norm(txt)
        for i in range(body_start, len(pages) + 1):
            if key and key in pages[i - 1]:
                found[txt] = i - (body_start - 1)
                break
    return body_start - 1, found


def main(argv):
    md = Path(argv[1]) if len(argv) > 1 else HERE / "开题报告_正文_V4.md"
    out = Path(argv[2]) if len(argv) > 2 else HERE / "开题报告_V4.docx"
    if not md.exists():
        raise SystemExit(f"缺少输入：{md}")

    fig = find_fig()
    heads, first_para = collect_heads(md)
    if first_para is None:
        raise SystemExit("未找到正文首段，无法定位正文起始页")

    tmpdir = Path(tempfile.mkdtemp())
    pages: dict = {}
    offset = None
    total = len(heads)

    for attempt in (1, 2, 3):
        probe = tmpdir / f"pass{attempt}.docx"
        br.build(md, probe, fig_path=fig, toc_pages=pages or None)
        pdf = to_pdf(probe)
        texts = page_texts(pdf)
        new_offset, new_pages = locate(texts, heads, first_para)

        missing = [t for _l, t in heads if t not in new_pages]
        if missing:
            print(f"[警告] 第 {attempt} 遍未定位到 {len(missing)} 条页码：{missing[:3]}")
        print(f"[pass{attempt}] 总页数 {len(texts)}，正文起始 PDF 页 {new_offset + 1}，"
              f"定位到 {len(new_pages)}/{total} 条")

        if new_pages == pages and new_offset == offset:
            print("         页码已稳定，停止迭代")
            pages, offset = new_pages, new_offset
            break
        pages, offset = new_pages, new_offset

    br.build(md, out, fig_path=fig, toc_pages=pages)
    print(f"[ok] {out}")
    print(f"     目录：TOC 域 + {len(pages)}/{total} 条带页码的域结果（正文从第 1 页起算）")
    print(f"     图：{fig.name if fig else '未找到'}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
