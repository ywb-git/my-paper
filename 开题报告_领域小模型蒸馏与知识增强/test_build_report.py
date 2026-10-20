#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_report.py 的自检用例。

只覆盖解析与格式的关键不变量，不做完整渲染比对（docx 二进制含时间戳，不稳定）。

运行：
    ../.venv-docx/bin/python -m unittest -v test_build_report.py
"""

import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import build_report as br  # noqa: E402

from docx import Document  # noqa: E402
from docx.shared import Emu  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402


class TestParser(unittest.TestCase):
    """解析器不变量。"""

    def test_heading_levels(self):
        md = "# 题\n\n## 章\n\n### 节\n\n#### 小节\n"
        blocks, _ = br.parse_markdown(md)
        heads = [p for k, p in blocks if k == "heading"]
        self.assertEqual([h[0] for h in heads], [1, 2, 3, 4])
        self.assertEqual(heads[1][1], "章")

    def test_table_ragged_row_is_padded_not_dropped(self):
        """列数不齐的行必须保留（补齐），且发出警告。"""
        md = "| A | B | C |\n|---|---|---|\n| 1 | 2 | 3 |\n| 1 | 2 |\n| 4 | 5 | 6 |\n"
        blocks, warnings = br.parse_markdown(md)
        tables = [p for k, p in blocks if k == "table"]
        self.assertEqual(len(tables), 1)
        header, rows = tables[0]
        self.assertEqual(len(rows), 3, "三行都应保留")
        self.assertEqual(rows[1], ["1", "2", ""], "缺列应补空")
        self.assertTrue(warnings, "应产生警告")
        self.assertIn("列数", warnings[0])

    def test_table_extra_columns_truncated_with_warning(self):
        md = "| A | B |\n|---|---|\n| 1 | 2 | 3 |\n"
        blocks, warnings = br.parse_markdown(md)
        _, rows = [p for k, p in blocks if k == "table"][0]
        self.assertEqual(rows[0], ["1", "2"])
        self.assertTrue(warnings)

    def test_unknown_placeholder_warns(self):
        md = "正文\n\n{{FIG99}}\n"
        blocks, warnings = br.parse_markdown(md)
        self.assertTrue(any("未识别" in w for w in warnings))
        # 仍作为段落输出，不丢内容
        self.assertTrue(any(k == "para" and "{{FIG99}}" in p for k, p in blocks))

    def test_known_placeholder_no_warning(self):
        md = "{{FIG41}}\n"
        blocks, warnings = br.parse_markdown(md)
        self.assertEqual([k for k, _ in blocks], ["figure"])
        self.assertFalse(warnings)

    def test_table_separator_only_forms_table(self):
        """没有分隔行不应识别为表格。"""
        md = "| A | B |\n| 1 | 2 |\n"
        blocks, _ = br.parse_markdown(md)
        self.assertFalse(any(k == "table" for k, _ in blocks))


class TestSplit(unittest.TestCase):
    """正文与参考文献拆分。"""

    def test_title_and_refs(self):
        md = "# 题目X\n\n## 章\n\n正文\n\n## 参考文献\n\n[1] A.\n[2] B.\n"
        title, body, refs = br.split_markdown(md)
        self.assertEqual(title, "题目X")
        self.assertEqual(len(refs), 2)
        self.assertNotIn("参考文献", body)

    def test_no_refs_section(self):
        md = "# 题目\n\n## 章\n\n正文\n"
        title, body, refs = br.split_markdown(md)
        self.assertEqual(title, "题目")
        self.assertEqual(refs, [])

    def test_missing_title_uses_placeholder(self):
        title, _, _ = br.split_markdown("## 章\n正文\n")
        self.assertIn("未设置题目", title)


class TestBuildOutput(unittest.TestCase):
    """渲染产物的格式不变量。"""

    @classmethod
    def setUpClass(cls):
        import tempfile
        cls.tmp = tempfile.TemporaryDirectory()
        cls.out = Path(cls.tmp.name) / "t.docx"
        md = HERE / "开题报告_正文.md"
        if not md.exists():
            raise unittest.SkipTest("缺少 开题报告_正文.md")
        cls.n_refs, cls.n_heads, cls.warnings = br.build(md, cls.out)
        cls.doc = Document(str(cls.out))

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_page_setup(self):
        s = self.doc.sections[0]
        self.assertAlmostEqual(Emu(s.page_width).cm, 21.0, places=1)
        self.assertAlmostEqual(Emu(s.page_height).cm, 29.7, places=1)
        self.assertAlmostEqual(Emu(s.top_margin).cm, 2.54, places=2)
        self.assertAlmostEqual(Emu(s.left_margin).cm, 3.17, places=2)

    def test_heading_sizes(self):
        h1 = [p for p in self.doc.paragraphs if p.style.name == "Heading 1"]
        h2 = [p for p in self.doc.paragraphs if p.style.name == "Heading 2"]
        self.assertTrue(h1 and h2)
        self.assertEqual(Emu(h1[0].runs[0].font.size).pt, 15.0)
        self.assertEqual(Emu(h2[0].runs[0].font.size).pt, 14.0)

    def test_heading_color_not_blue(self):
        for p in self.doc.paragraphs:
            if not p.style.name.startswith("Heading") or not p.runs:
                continue
            col = p.runs[0].font.color
            if col is not None and col.rgb is not None:
                self.assertIn(str(col.rgb), ("000000", "FFFFFF"),
                              f"标题出现非黑色 {col.rgb}")

    def test_body_style(self):
        """样式级只管字号与行距；首行缩进已下放到段落级（含 firstLineChars）。"""
        n = self.doc.styles["Normal"]
        self.assertEqual(Emu(n.font.size).pt, 12.0)
        self.assertEqual(n.paragraph_format.line_spacing, 1.3)
        # 首行缩进不再放在样式级：正文段落由 apply_body_format 逐个设置，
        # 样式级重复设置会被段落级覆盖，留在样式里只会误导。
        self.assertIsNone(n.paragraph_format.first_line_indent,
                          "Normal 样式不应再设 first_line_indent")

    def test_reference_numbering_continuous(self):
        import re
        nums = []
        for p in self.doc.paragraphs:
            m = re.match(r"^\[(\d+)\]", p.text.strip())
            if m:
                nums.append(int(m.group(1)))
        self.assertEqual(nums, list(range(1, len(nums) + 1)), "参考文献编号须连续")
        self.assertEqual(len(nums), self.n_refs)

    def test_citations_superscript(self):
        import re
        found = False
        for p in self.doc.paragraphs:
            for r in p.runs:
                if re.fullmatch(r"\[\d+\](?:\[\d+\])*", r.text) and r.font.superscript:
                    found = True
                    break
            if found:
                break
        self.assertTrue(found, "正文引用应渲染为上标")

    def test_table_has_no_empty_leading_run(self):
        """表格单元格不应因 cell.text=\"\" 留下空 run。"""
        self.assertTrue(self.doc.tables)
        cell = self.doc.tables[0].rows[0].cells[0]
        para = cell.paragraphs[0]
        self.assertEqual(len(para.runs), 1, f"应只有 1 个 run，实际 {len(para.runs)}")
        self.assertTrue(para.text.strip())

    def test_cover_info_blank(self):
        """封面个人信息必须留空（下划线占位）。"""
        head = [p.text for p in self.doc.paragraphs[:20] if p.text.strip()]
        joined = "\n".join(head)
        self.assertIn("北京科技大学", joined.replace(" ", ""))
        self.assertIn("选题报告及文献综述", joined)
        self.assertIn("＿", joined, "信息栏应有下划线占位")

    def test_document_ends_with_content_not_empty_para(self):
        body = self.doc.element.body
        kids = list(body)
        # 最后一个非 sectPr 元素不应是纯空段落
        tail = [k for k in kids if k.tag != qn("w:sectPr")]
        self.assertTrue(tail)
        last = tail[-1]
        if last.tag == qn("w:p"):
            txt = "".join(t.text or "" for t in last.findall(".//" + qn("w:t")))
            self.assertTrue(txt.strip(), "末尾不应是空段落")


class TestCLIContract(unittest.TestCase):
    """命令行契约：默认输出名按模式区分。"""

    def test_default_names_differ(self):
        self.assertNotEqual(br.DEFAULT_OUT_TOC, br.DEFAULT_OUT_STATIC)


class TestGlobalState(unittest.TestCase):
    """parse_markdown 之间不得共享警告状态（无模块级全局）。"""

    def test_warnings_do_not_accumulate(self):
        md = "| A |\n|---|\n| 1 | 2 |\n"
        _, w1 = br.parse_markdown(md)
        _, w2 = br.parse_markdown(md)
        self.assertEqual(len(w1), 1)
        self.assertEqual(len(w2), 1, "第二次调用不应累积上一次的警告")

    def test_no_module_level_warnings_attr(self):
        self.assertFalse(hasattr(br, "WARNINGS"),
                         "不应再有模块级 WARNINGS 全局")

    def test_inline_unknown_placeholder_warns(self):
        """行内占位符也要告警（原实现只识别整行）。"""
        md = "见图 {{FIG99}} 所示。\n"
        _, warnings = br.parse_markdown(md)
        self.assertTrue(any("未识别" in w for w in warnings),
                        "行内占位符应告警")


class TestCLI(unittest.TestCase):
    """main() 的命令行契约（此前完全没有测试）。"""

    def setUp(self):
        import tempfile
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.md = self.dir / "in.md"
        self.md.write_text("# 题\n\n## 章\n\n正文\n\n## 参考文献\n\n[1] A.\n",
                           encoding="utf-8")

    def tearDown(self):
        self.tmp.cleanup()

    def _run(self, *args):
        """调用 main，返回 (退出码, stdout, stderr)。"""
        import contextlib
        import io
        out, err = io.StringIO(), io.StringIO()
        code = 0
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                br.main(["build_report.py", *args])
            except SystemExit as e:
                code = e.code if isinstance(e.code, int) else 1
        return code, out.getvalue(), err.getvalue()

    def test_help_exits_zero(self):
        code, out, _ = self._run("--help")
        self.assertEqual(code, 0)
        self.assertIn("用法", out)
        self.assertIn("--fig-caption", out)

    def test_unknown_flag_exits_two(self):
        code, _, err = self._run("--bogus")
        self.assertEqual(code, 2)
        self.assertIn("未知参数", err)

    def test_too_many_positionals_exits_two(self):
        code, _, err = self._run("a.md", "b.docx", "c.docx")
        self.assertEqual(code, 2)
        self.assertIn("参数过多", err)

    def test_missing_input_exits_one(self):
        code, _, err = self._run(str(self.dir / "nope.md"), str(self.dir / "o.docx"))
        self.assertEqual(code, 1)
        self.assertIn("找不到输入文件", err)

    def test_refuses_overwrite_without_force(self):
        out = self.dir / "o.docx"
        code1, _, _ = self._run(str(self.md), str(out))
        self.assertEqual(code1, 0)
        self.assertTrue(out.exists())
        code2, _, err = self._run(str(self.md), str(out))
        self.assertEqual(code2, 2, "已存在且无 --force 应拒绝")
        self.assertIn("已存在", err)

    def test_force_overwrites(self):
        out = self.dir / "o.docx"
        self._run(str(self.md), str(out))
        code, _, _ = self._run(str(self.md), str(out), "--force")
        self.assertEqual(code, 0)

    def test_refuses_same_input_output_path(self):
        """输入输出同路径会销毁输入 md，必须拒绝。"""
        before = self.md.read_text(encoding="utf-8")
        code, _, err = self._run(str(self.md), str(self.md), "--force")
        self.assertEqual(code, 2)
        self.assertIn("同一个文件", err)
        self.assertEqual(self.md.read_text(encoding="utf-8"), before,
                         "输入文件内容不得被改动")

    def test_missing_output_dir_exits_one(self):
        code, _, err = self._run(str(self.md), str(self.dir / "no" / "such" / "o.docx"))
        self.assertEqual(code, 1)
        self.assertIn("输出目录不存在", err)

    def test_default_output_names_differ_by_mode(self):
        """真正调用 main()，验证不带位置参数时落到两个不同的默认文件名。

        用临时目录替换 build_report 模块的 __file__ 所在目录，
        使 main 内的 here 指向临时目录，从而在临时目录里真实落盘，
        不污染仓库。
        """
        self.assertNotEqual(br.DEFAULT_OUT_TOC, br.DEFAULT_OUT_STATIC)

        import shutil
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            # 准备 main 需要的默认输入与图
            (d / "开题报告_正文.md").write_text(
                "# 题\n\n## 章\n\n正文\n", encoding="utf-8")
            figsrc = None
            for cand in ("技术路线.png", "图4-1_技术路线图.png", "图3-1_技术路线图.png"):
                if (HERE / cand).exists():
                    figsrc = HERE / cand
                    break
            if figsrc:
                shutil.copy(figsrc, d / figsrc.name)

            # 用 patch 把 br.__file__ 指到临时目录，使 here 解析为临时目录
            from unittest import mock
            fake_file = str(d / "build_report.py")
            with mock.patch.object(br, "__file__", fake_file):
                code1, _, err1 = self._run()                 # 无位置参数 -> TOC 域版
                code2, _, err2 = self._run("--static-toc")   # 无位置参数 -> 静态目录版

            self.assertEqual(code1, 0, f"TOC 模式失败: {err1}")
            self.assertEqual(code2, 0, f"静态模式失败: {err2}")
            toc_path = d / br.DEFAULT_OUT_TOC
            static_path = d / br.DEFAULT_OUT_STATIC
            self.assertTrue(toc_path.exists(), f"未生成 {toc_path.name}")
            self.assertTrue(static_path.exists(), f"未生成 {static_path.name}")
            self.assertNotEqual(toc_path.name, static_path.name,
                                "两种模式的默认名不能相同")

    def test_fig_caption_option_applied(self):
        """--fig-caption 的取值必须真的出现在生成文档的图题里。"""
        # 构造带 {{FIG}} 占位与图的输入，才能产生图题
        md_with_fig = self.dir / "withfig.md"
        md_with_fig.write_text("# 题\n\n## 章\n\n{{FIG}}\n", encoding="utf-8")
        fig = None
        for cand in ("技术路线.png", "图4-1_技术路线图.png", "图3-1_技术路线图.png"):
            if (HERE / cand).exists():
                fig = HERE / cand
                break
        if fig is None:
            raise unittest.SkipTest("缺少技术路线图")

        out = self.dir / "o.docx"
        code, _, _ = self._run(str(md_with_fig), str(out),
                               "--fig-caption", "图 9-9　自定义")
        self.assertEqual(code, 0)
        doc = Document(str(out))
        caps = [p.text.strip() for p in doc.paragraphs if p.text.strip().startswith("图 ")]
        self.assertIn("图 9-9　自定义", caps, f"图题未写入，实际 {caps}")

    def test_fig_caption_rejects_option_value(self):
        """--fig-caption 的取值若以 - 开头应报错（防止吞掉 --force）。"""
        code, _, err = self._run(str(self.md), "--fig-caption", "--force")
        self.assertEqual(code, 2)
        self.assertIn("不能以 - 开头", err)

    def test_fig_caption_rejects_empty(self):
        code, _, err = self._run(str(self.md), "--fig-caption", "   ")
        self.assertEqual(code, 2)
        self.assertIn("不能为空", err)

    def test_input_is_directory_rejected(self):
        code, _, err = self._run(str(self.dir), str(self.dir / "o.docx"))
        self.assertEqual(code, 1)
        self.assertIn("是目录", err)

    def test_non_utf8_input_rejected(self):
        bad = self.dir / "bad.md"
        bad.write_bytes(b"\xff\xfe test")
        code, _, err = self._run(str(bad), str(self.dir / "o.docx"))
        self.assertEqual(code, 1)
        self.assertIn("UTF-8", err)

    def test_tail_non_utf8_rejected(self):
        """非法字节在文件后部时也必须检出（原实现只读前 4096 字节）。"""
        bad = self.dir / "tail_bad.md"
        head = ("正常内容" * 3000).encode("utf-8")
        bad.write_bytes(head + b"\xff\xfe")
        code, _, err = self._run(str(bad), str(self.dir / "o.docx"))
        self.assertEqual(code, 1, "文件后部的非法字节应被检出")
        self.assertIn("UTF-8", err)

    def test_output_is_directory_rejected(self):
        tdir = self.dir / "outdir"
        tdir.mkdir()
        code, _, err = self._run(str(self.md), str(tdir), "--force")
        self.assertEqual(code, 1)
        self.assertIn("输出路径是目录", err)

    def test_build_accepts_prefetched_text(self):
        """build(md_text=...) 应跳过重复读文件。"""
        out = self.dir / "prefetched.docx"
        text = "# 题\n\n## 章\n\n正文\n"
        n_refs, n_heads, warnings = br.build(self.md, out, md_text=text)
        self.assertTrue(out.exists())
        # n_heads 统计 md 里的标题数（# 题 与 ## 章 各计 1）
        self.assertEqual(n_heads, 2)

    def test_fig_caption_missing_value_exits_two(self):
        code, _, err = self._run(str(self.md), "--fig-caption")
        self.assertEqual(code, 2)
        self.assertIn("缺少取值", err)


class TestFormatCoverage(unittest.TestCase):
    """此前未覆盖的格式断言：目录域、页码域、图题、表格样式。"""

    @classmethod
    def setUpClass(cls):
        import tempfile
        cls.tmp = tempfile.TemporaryDirectory()
        cls.out = Path(cls.tmp.name) / "t.docx"
        md = HERE / "开题报告_正文.md"
        if not md.exists():
            raise unittest.SkipTest("缺少 开题报告_正文.md")
        fig = None
        for cand in ("技术路线.png", "图4-1_技术路线图.png", "图3-1_技术路线图.png"):
            if (HERE / cand).exists():
                fig = HERE / cand
                break
        cls.n_refs, cls.n_heads, cls.warnings = br.build(md, cls.out, fig_path=fig)
        cls.doc = Document(str(cls.out))

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_toc_field_present_in_toc_mode(self):
        joined = "\n".join(p._p.xml for p in self.doc.paragraphs)
        self.assertIn("TOC", joined, "TOC 域应存在")

    def test_page_number_field_in_footer(self):
        footer = self.doc.sections[0].footer
        xml = "\n".join(p._p.xml for p in footer.paragraphs)
        self.assertIn("PAGE", xml, "页脚应有 PAGE 域")

    def test_figure_caption_present(self):
        caps = [p.text.strip() for p in self.doc.paragraphs
                if p.text.strip().startswith("图 ")]
        self.assertTrue(caps, "应有图题")

    def test_table_style_is_grid(self):
        self.assertTrue(self.doc.tables)
        self.assertEqual(self.doc.tables[0].style.name, "Table Grid")


class TestCrossref(unittest.TestCase):
    """参考文献交叉引用：书签 + REF 域。"""

    @classmethod
    def setUpClass(cls):
        import tempfile
        cls.tmp = tempfile.TemporaryDirectory()
        cls.out = Path(cls.tmp.name) / "x.docx"
        md = HERE / "开题报告_正文.md"
        if not md.exists():
            raise unittest.SkipTest("缺少 开题报告_正文.md")
        cls.n_refs, cls.n_heads, cls.warnings = br.build(md, cls.out)
        cls.doc = Document(str(cls.out))
        cls.xml = "".join(p._p.xml for p in cls.doc.paragraphs)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_ref_fields_exist(self):
        self.assertGreater(self.xml.count("REF _Ref_ref"), 0,
                           "正文引用应为 REF 域而非纯文本")

    def test_bookmarks_match_reference_count(self):
        from docx.oxml.ns import qn
        names = [e.get(qn("w:name"))
                 for e in self.doc.element.body.iter(qn("w:bookmarkStart"))]
        ref_bm = [n for n in names if n and n.startswith("_Ref_ref")]
        self.assertEqual(len(ref_bm), self.n_refs,
                         "每条参考文献应有一个书签")

    def test_ref_field_has_cached_display(self):
        """每个 REF 域的 separate 与 end 之间必须有形如 [n] 的缓存显示值。

        注意：不能用 "[1]" in xml 之类兜底——参考文献条目本身就含 [1] 字面量，
        那样断言恒真，无法捕获缓存值缺失（正是这条曾让书签为空的问题逃过检测）。
        """
        import re
        self.assertGreater(self.xml.count("REF _Ref_ref"), 0, "没有 REF 域")
        # 每个 REF 域都应在其 fldChar separate 之后紧跟 <w:t>[n]</w:t>
        pat = re.compile(
            r"REF (_Ref_ref\d+)[^<]*</w:instrText>.*?fldCharType=\"separate\".*?"
            r"<w:t[^>]*>\[(\d+)\]</w:t>", re.S)
        found = pat.findall(self.xml)
        n_ref_fields = self.xml.count("REF _Ref_ref")
        self.assertEqual(len(found), n_ref_fields,
                         f"{n_ref_fields} 个 REF 域中只有 {len(found)} 个带 [n] 缓存值")

    def test_bookmark_contains_only_index(self):
        """书签范围必须只包住 [n] 编号，不能包整条文献。

        否则 REF 域会显示成整条文献全文；若书签为空，Word 更新域后引用会消失。
        """
        from docx.oxml.ns import qn
        body = self.doc.element.body
        starts = {}
        for el in body.iter():
            if el.tag == qn("w:bookmarkStart"):
                nm = el.get(qn("w:name"))
                if nm and nm.startswith("_Ref_ref"):
                    starts[nm] = el
        self.assertEqual(len(starts), self.n_refs, "书签数应等于文献数")

        empty = []
        for nm, st in starts.items():
            bid = st.get(qn("w:id"))
            seq = list(body.iter())
            i = seq.index(st)
            texts = []
            for el2 in seq[i + 1:]:
                tag2 = el2.tag.split("}")[-1]
                if tag2 == "bookmarkEnd" and el2.get(qn("w:id")) == bid:
                    break
                if tag2 == "t" and (el2.text or "").strip():
                    texts.append(el2.text.strip())
            joined = "".join(texts)
            if not joined:
                empty.append((nm, "空书签"))
            elif not (joined.startswith("[") and joined.endswith("]")):
                empty.append((nm, f"含多余文本: {joined[:30]}"))
        self.assertEqual(empty, [], f"书签范围不合规: {empty[:3]}")

    def test_no_bare_text_citation(self):
        """正文引用不应是纯文本；参考文献条目的编号 run 不在此列。"""
        import re
        from docx.oxml.ns import qn
        # 先收集参考文献段落的编号 run（它们本就是纯文本 [n]，且被书签包裹）
        ref_num_elems = set()
        for p in self.doc.paragraphs:
            if p.text.strip().startswith("[") and re.match(r"^\[\d+\]", p.text.strip()):
                # 该段是参考文献条目：第一个 run 就是编号
                if p.runs and re.fullmatch(r"\[\d+\]", p.runs[0].text.strip()):
                    ref_num_elems.add(id(p.runs[0]._element))
        bare = 0
        for p in self.doc.paragraphs:
            if p.text.strip().startswith("[") and re.match(r"^\[\d+\]", p.text.strip()):
                continue  # 跳过参考文献段落本身
            for r in p.runs:
                if re.fullmatch(r"\[\d+\](?:\[\d+\])*", r.text.strip()):
                    if "fldChar" not in r._element.xml:
                        bare += 1
        self.assertEqual(bare, 0, f"正文仍有 {bare} 处纯文本引用")

    def test_bookmark_names_are_unique(self):
        from docx.oxml.ns import qn
        names = [e.get(qn("w:name"))
                 for e in self.doc.element.body.iter(qn("w:bookmarkStart"))]
        ref_bm = [n for n in names if n and n.startswith("_Ref_ref")]
        self.assertEqual(len(ref_bm), len(set(ref_bm)), "书签名不得重复")

    def test_reference_hanging_indent_real(self):
        """参考文献必须是真悬挂缩进：left 与 hanging 同时存在，且无 firstLine。"""
        from docx.oxml.ns import qn
        target = None
        for p in self.doc.paragraphs:
            if p.text.strip().startswith("[1]"):
                target = p
                break
        self.assertIsNotNone(target, "未找到第一条参考文献")
        ppr = target._p.find(qn("w:pPr"))
        ind = ppr.find(qn("w:ind"))
        self.assertIsNotNone(ind, "参考文献缺少 w:ind")
        self.assertEqual(ind.get(qn("w:left")), "480")
        self.assertEqual(ind.get(qn("w:hanging")), "480",
                         "hanging 必须为正数才能形成悬挂缩进")
        self.assertIsNone(ind.get(qn("w:firstLine")),
                          "firstLine 与 hanging 互斥，不应同时存在")

    def test_body_paragraph_indent_chars(self):
        """正文段落首行缩进须带 firstLineChars=200（中文 2 字符）。"""
        from docx.oxml.ns import qn
        for p in self.doc.paragraphs:
            if p.style.name == "Normal" and len(p.text.strip()) > 60 \
                    and not p.text.strip().startswith("["):
                ppr = p._p.find(qn("w:pPr"))
                ind = ppr.find(qn("w:ind")) if ppr is not None else None
                self.assertIsNotNone(ind)
                self.assertEqual(ind.get(qn("w:firstLineChars")), "200")
                self.assertEqual(ind.get(qn("w:firstLine")), "480")
                return
        self.fail("未找到正文段落")

    def test_line_spacing_encoded_as_multiple(self):
        """行距须用 line=312 + lineRule=auto（1.3 倍），不能是固定 twips。"""
        from docx.oxml.ns import qn
        for p in self.doc.paragraphs:
            if p.style.name == "Normal" and len(p.text.strip()) > 60:
                ppr = p._p.find(qn("w:pPr"))
                sp = ppr.find(qn("w:spacing")) if ppr is not None else None
                self.assertIsNotNone(sp)
                self.assertEqual(sp.get(qn("w:line")), "312")
                self.assertEqual(sp.get(qn("w:lineRule")), "auto")
                return
        self.fail("未找到正文段落")


if __name__ == "__main__":
    unittest.main(verbosity=2)
