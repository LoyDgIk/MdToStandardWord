# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import re
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from md2std import cli, docx_builder, md_parser, model


def _build_docx_xml(markdown: str) -> str:
    sdoc = md_parser.parse(markdown)
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        docx_builder.build(sdoc, cli._DEFAULT_TEMPLATE, path)
        with zipfile.ZipFile(path) as zf:
            return zf.read("word/document.xml").decode("utf-8", errors="ignore")
    finally:
        if os.path.exists(path):
            os.remove(path)


def _build_template_docx_xml(markdown: str, kind: str = "group") -> str:
    sdoc = md_parser.parse(markdown)
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        docx_builder.build(sdoc, cli._resolve_default_template(kind), path, kind=kind)
        with zipfile.ZipFile(path) as zf:
            return zf.read("word/document.xml").decode("utf-8", errors="ignore")
    finally:
        if os.path.exists(path):
            os.remove(path)


def _build_cover_docx_xml(markdown: str, kind: str = "group") -> str:
    sdoc = md_parser.parse(markdown)
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        docx_builder.build_cover(sdoc, path, kind=kind)
        with zipfile.ZipFile(path) as zf:
            return zf.read("word/document.xml").decode("utf-8", errors="ignore")
    finally:
        if os.path.exists(path):
            os.remove(path)


def _build_cover_docx_parts(markdown: str, kind: str = "group") -> dict:
    sdoc = md_parser.parse(markdown)
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        docx_builder.build_cover(sdoc, path, kind=kind)
        with zipfile.ZipFile(path) as zf:
            return {
                "document": zf.read("word/document.xml"),
                "rels": zf.read("word/_rels/document.xml.rels"),
                "media": {
                    name: zf.read(name)
                    for name in zf.namelist()
                    if name.startswith("word/media/") and not name.endswith("/")
                },
                "styles": zf.read("word/styles.xml"),
                "numbering": zf.read("word/numbering.xml"),
                "settings": zf.read("word/settings.xml"),
            }
    finally:
        if os.path.exists(path):
            os.remove(path)


class CrossReferenceParserTest(unittest.TestCase):
    def test_new_reference_syntax_is_typed_and_forward_refs_work(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "见{{tbl:classify:label}}。\n\n"
            "{表：#tbl:classify} 温泉利用分类\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| 医疗保健 |\n"
        )

        paragraph = next(b for b in doc.body if isinstance(b, model.Paragraph))
        ref = next(s for s in paragraph.spans if isinstance(s, model.RefSpan))
        table = next(b for b in doc.body if isinstance(b, model.TableModel))

        self.assertEqual(ref.ref_type, "tbl")
        self.assertEqual(ref.target, "classify")
        self.assertEqual(ref.mode, "label")
        self.assertEqual(table.anchor_id, "classify")

    def test_legacy_reference_syntax_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "旧交叉引用语法"):
            md_parser.parse("# 范围\n\n按式（{@eq-depth:a}）计算。")

    def test_caption_handwritten_number_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "不要手写编号"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:bad} 表1　温泉利用分类\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| 医疗保健 |\n"
            )

    def test_unknown_reference_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "未知交叉引用"):
            md_parser.parse("# 范围\n\n见{{tbl:missing:label}}。")

    def test_formula_tag_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "不要使用 LaTeX"):
            md_parser.parse("# 范围\n\n$$x=1\\tag{1}$${#eq:x}")

    def test_html_table_with_colspan_is_parsed(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:speed} 测试车速\n\n"
            "<table><tr><td>最高设计车速（<eq>v_{\\text{max}}</eq>）</td>"
            "<td>测试车速</td></tr>"
            "<tr><td><eq>v_{\\text{max}}</eq>≤45</td><td>80%</td></tr>"
            "<tr><td colspan=\"2\">注：按临近分度线取值。</td></tr></table>"
        )

        table = next(b for b in doc.body if isinstance(b, model.TableModel))
        self.assertEqual(table.header, ["最高设计车速（vmax）", "测试车速"])
        self.assertEqual(table.header_parts[0][1].kind, "formula")
        self.assertEqual(table.header_parts[0][1].text, "v_{\\text{max}}")
        self.assertEqual(table.rows[-1], ["注：按临近分度线取值。"])
        self.assertEqual(table.row_colspans[-1], [2])

    def test_nested_ordered_list_is_kept_as_second_level(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "1. 车辆装备：\n"
            "   1. 为整备质量再加上驾驶员。\n"
            "   2. 所配轮胎气压为正常行驶用气压。\n"
        )

        lists = [b for b in doc.body if isinstance(b, model.ListBlock)]
        self.assertEqual(len(lists), 2)
        self.assertEqual(lists[0].level, 1)
        self.assertEqual(lists[1].level, 2)
        self.assertEqual([item.spans[0].text for item in lists[1].items], [
            "为整备质量再加上驾驶员。",
            "所配轮胎气压为正常行驶用气压。",
        ])

    def test_foreword_extra_notes_support_dash_list_group(self):
        doc = md_parser.parse(
            "---\n"
            "foreword:\n"
            "  extra_notes:\n"
            "    - 本文件及其所代替文件的历次版本发布情况为：\n"
            "    - - 1994年首次发布为GB 15082—1994。\n"
            "      - 本次为第三次修订。\n"
            "---\n\n"
            "# 范围\n\n正文。\n"
        )

        self.assertEqual(doc.meta.foreword.extra_notes[0], "本文件及其所代替文件的历次版本发布情况为：")
        self.assertEqual(doc.meta.foreword.extra_notes[1], [
            "1994年首次发布为GB 15082—1994。",
            "本次为第三次修订。",
        ])


class CrossReferenceDocxTest(unittest.TestCase):
    _W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    def test_docx_uses_visible_builtin_seq_and_resets_appendix_tables(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "见{{tbl:main:label}}，完整题名为{{tbl:main:full}}，按{{eq:rate:label}}计算。\n\n"
            "{表：#tbl:main} 主表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| 正文 |\n\n"
            "![主图{#fig:flow}](missing.png)\n\n"
            "$$x=1$${#eq:rate}\n\n"
            "# 附录 规范性 第一附录\n\n"
            "{表：#tbl:appA} 附录A表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| A |\n\n"
            "# 附录 资料性 第二附录\n\n"
            "{表：#tbl:appB} 附录B表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| B |\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n"
        )

        self.assertIn(" SEQ 表 \\* ARABIC \\r 1 ", xml)
        self.assertIn(" SEQ 图 \\* ARABIC \\r 1 ", xml)
        self.assertIn(" SEQ 公式 \\* ARABIC \\r 1 ", xml)
        self.assertIn("<w:t>A.</w:t>", xml)
        self.assertIn("<w:t>B.</w:t>", xml)
        self.assertNotIn("SEQ 表A", xml)
        self.assertNotIn("SEQ 表B", xml)
        self.assertNotIn("SEQ 图表", xml)
        self.assertNotIn("SEQ 公式A", xml)
        self.assertNotIn("SEQ 公式B", xml)
        self.assertGreaterEqual(xml.count('<w:numId w:val="0"/>'), 3)
        self.assertIn(" REF _Ref", xml)
        self.assertIn('w:name="_Ref', xml)
        self.assertIn("<w:vanish", xml)
        self.assertIn("<w:hyperlink", xml)

    def test_formula_caption_anchor_runs_are_hidden(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "按{{eq:rate:label}}计算。\n\n"
            "$$H=(T_r-T_0)/G+h$${#eq:rate}\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n"
        )
        root = ET.fromstring(xml)
        formula_anchor = None
        for paragraph in root.findall(".//w:p", self._W_NS):
            if any(
                instr.text and "SEQ 公式" in instr.text
                for instr in paragraph.findall(".//w:instrText", self._W_NS)
            ):
                formula_anchor = paragraph
                break

        self.assertIsNotNone(formula_anchor)
        runs = formula_anchor.findall("w:r", self._W_NS)
        self.assertGreater(len(runs), 0)
        for run in runs:
            self.assertIsNotNone(
                run.find("w:rPr/w:vanish", self._W_NS),
                ET.tostring(run, encoding="unicode"),
            )

    def test_declared_standard_styles_are_used_for_examples_sources_and_appendix_title(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "示例：\n\n"
            "这是示例内容。\n\n"
            "[来源：GB/T 1.1—2020，3.1]\n\n"
            "# 附录 规范性 样式验证附录\n\n"
            "## 附录条\n\n"
            "附录正文。\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n"
        )

        self.assertIn('w:pStyle w:val="182"', xml)  # 标准文件_示例内容
        self.assertIn('w:rStyle w:val="191"', xml)  # 标准文件_来源
        appendix_head_start = xml.index('w:pStyle w:val="76"')
        appendix_head_end = xml.index("</w:p>", appendix_head_start)
        appendix_head = xml[appendix_head_start:appendix_head_end]
        self.assertGreaterEqual(appendix_head.count("<w:br/>"), 2)
        self.assertIn("（规范性）", appendix_head)
        self.assertIn("样式验证附录", appendix_head)

    def test_appendix_starts_on_new_page(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "正文。\n\n"
            "# 附录 规范性 分页附录\n\n"
            "附录正文。\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n"
        )
        root = ET.fromstring(xml)
        appendix_head = self._et_paragraph_containing(root, "分页附录")

        self.assertIsNotNone(appendix_head)
        self.assertIsNotNone(appendix_head.find("w:pPr/w:pageBreakBefore", self._W_NS))

    def test_document_end_line_follows_references_before_section_break(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "正文。\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n"
        )
        root = ET.fromstring(xml)
        body = root.find("w:body", self._W_NS)
        paragraphs = [p for p in body if p.tag == self._w_tag("p")]

        ref_title_index = next(
            i for i, p in enumerate(paragraphs)
            if self._et_text(p) == "参考文献"
        )
        ref_item_index = next(
            i for i, p in enumerate(paragraphs[ref_title_index + 1:], ref_title_index + 1)
            if "GB/T 1.1—2020" in self._et_text(p)
        )
        end_line_index = next(
            i for i, p in enumerate(paragraphs[ref_item_index + 1:], ref_item_index + 1)
            if p.findall(".//w:drawing", self._W_NS)
            and not self._et_text(p).strip()
            and self._jc_value(p) == "center"
        )
        self.assertLess(ref_item_index, end_line_index)
        self.assertEqual(end_line_index, len(paragraphs) - 1)
        self.assertFalse(any(
            p.find("w:pPr/w:sectPr", self._W_NS) is not None
            for p in paragraphs[end_line_index + 1:]
        ))

    def test_optional_appendix_reference_and_index_sections_can_be_absent(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "正文。\n"
        )
        root = ET.fromstring(xml)
        body = root.find("w:body", self._W_NS)
        paragraphs = [p for p in body if p.tag == self._w_tag("p")]
        texts = [self._et_text(p) for p in paragraphs]

        self.assertNotIn("参考文献", texts)
        self.assertNotIn("索引", texts)
        self.assertFalse(any(
            p.find('w:pPr/w:pStyle[@w:val="76"]', self._W_NS) is not None
            for p in paragraphs
        ))

        body_index = next(i for i, text in enumerate(texts) if text == "正文。")
        end_line_index = next(
            i for i, p in enumerate(paragraphs[body_index + 1:], body_index + 1)
            if p.findall(".//w:drawing", self._W_NS)
            and not self._et_text(p).strip()
            and self._jc_value(p) == "center"
        )
        self.assertLess(body_index, end_line_index)
        self.assertEqual(end_line_index, len(paragraphs) - 1)

    def test_index_section_renders_groups_items_and_document_end_line(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "正文。\n\n"
            "# 索引\n\n"
            "## B\n\n"
            "- 必备要素：3.2.5，6.2.2.1\n"
            "- 标准：3.1.2，4.1\n\n"
            "## C\n\n"
            "- 参考文献：表 3，8.2\n"
        )
        root = ET.fromstring(xml)
        body = root.find("w:body", self._W_NS)
        paragraphs = [p for p in body if p.tag == self._w_tag("p")]
        texts = [self._et_text(p) for p in paragraphs]

        self.assertIn("索引", texts)
        self.assertIn("B", texts)
        self.assertIn("C", texts)
        self.assertIn("必备要素3.2.5，6.2.2.1", texts)
        self.assertIn("参考文献表 3，8.2", texts)

        index_item = next(p for p in paragraphs if self._et_text(p).startswith("必备要素"))
        self.assertIsNotNone(index_item.find('w:pPr/w:pStyle[@w:val="210"]', self._W_NS))
        self.assertIsNotNone(index_item.find(".//w:tab", self._W_NS))
        end_line = paragraphs[-1]
        self.assertTrue(end_line.findall(".//w:drawing", self._W_NS))
        self.assertEqual(self._jc_value(end_line), "center")

    def test_index_items_require_group_and_colon(self):
        with self.assertRaisesRegex(ValueError, "索引项必须放在"):
            md_parser.parse("# 范围\n\n正文。\n\n# 索引\n\n- 标准：3.1\n")
        with self.assertRaisesRegex(ValueError, "索引项必须写成"):
            md_parser.parse("# 范围\n\n正文。\n\n# 索引\n\n## B\n\n- 标准 3.1\n")

    def test_caption_label_bookmarks_include_complete_seq_field(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "{表：#tbl:main} 正文表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| 正文 |\n\n"
            "# 附录 规范性 第一附录\n\n"
            "{表：#tbl:appA} 附录A表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| A |\n\n"
            "# 附录 资料性 第二附录\n\n"
            "{表：#tbl:appB} 附录B表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| B |\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n"
        )

        for anchor_id, title, prefix in (("main", "正文表", None), ("appB", "附录B表", "B.")):
            para = self._paragraph_containing(xml, title)
            label_name = docx_builder._native_ref_name("tbl", anchor_id, "label")
            num_name = docx_builder._native_ref_name("tbl", anchor_id, "num")
            label_id = self._bookmark_id(para, label_name)
            num_id = self._bookmark_id(para, num_name)

            field_end = para.index('<w:fldChar w:fldCharType="end"/>')
            self.assertLess(para.index("<w:t>表</w:t>"), field_end)
            if prefix is not None:
                self.assertLess(para.index(f"<w:t>{prefix}</w:t>"), field_end)
            self.assertGreater(para.index(f'<w:bookmarkEnd w:id="{label_id}"/>'), field_end)
            self.assertGreater(para.index(f'<w:bookmarkEnd w:id="{num_id}"/>'), field_end)

    def _paragraph_containing(self, xml: str, text: str) -> str:
        text_pos = xml.index(text)
        para_start = xml.rfind("<w:p", 0, text_pos)
        para_end = xml.index("</w:p>", text_pos) + len("</w:p>")
        return xml[para_start:para_end]

    def _bookmark_id(self, xml: str, name: str) -> str:
        match = re.search(
            r'<w:bookmarkStart w:id="([^"]+)" w:name="' + re.escape(name) + r'"/>',
            xml,
        )
        self.assertIsNotNone(match, name)
        return match.group(1)

    def _et_paragraph_containing(self, root, text: str):
        for paragraph in root.findall(".//w:p", self._W_NS):
            if text in self._et_text(paragraph):
                return paragraph
        return None

    def _et_text(self, element) -> str:
        return "".join(t.text or "" for t in element.findall(".//w:t", self._W_NS))

    def _jc_value(self, paragraph) -> str:
        jc = paragraph.find("w:pPr/w:jc", self._W_NS)
        return jc.get(self._w_tag("val")) if jc is not None else ""

    def _w_tag(self, name: str) -> str:
        return "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + name


class TemplateBackendDocxTest(unittest.TestCase):
    _W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    def test_national_template_backend_uses_national_template_and_cleans_placeholders(self):
        xml = _build_template_docx_xml(
            "---\n"
            "standard_type: 国家标准\n"
            "number: GB/T 99999-2026\n"
            "title: 国家模板测试标准\n"
            "title_en: National template test standard\n"
            "ics: \"27.010\"\n"
            "ccs: F 10\n"
            "publish_date: 2026-06-01\n"
            "implement_date: 2026-07-01\n"
            "publisher: 国家市场监督管理总局 国家标准化管理委员会\n"
            "---\n"
            "# 范围\n\n"
            "正文。\n",
            kind="national",
        )

        self.assertIn("中华人民共和国国家标准", xml)
        self.assertIn("GB/T 99999-2026", xml)
        self.assertIn("国家模板测试标准", xml)
        self.assertIn("National template test standard", xml)
        self.assertIn("27.010", xml)
        self.assertIn("F 10", xml)
        self.assertIn("2026-06-01发布", xml)
        self.assertIn("2026-07-01实施", xml)
        self.assertNotIn("国家市场监督管理总局", xml)
        self.assertNotIn("国家标准化管理委员会", xml)
        self.assertIn("国标发布单位", xml)
        self.assertNotIn("点击此处添加", xml)
        self.assertNotIn("本草案完成时间", xml)
        self.assertNotIn("章标题", xml)
        self.assertNotIn("条标题", xml)

    def test_template_backend_inserts_body_standard_title_before_scope(self):
        xml = _build_template_docx_xml(
            "---\n"
            "title: 模板正文标题测试标准\n"
            "---\n"
            "# 范围\n\n"
            "正文。\n",
            kind="group",
        )
        paragraphs = self._paragraphs(xml)
        texts = [self._et_text(p) for p in paragraphs]
        scope_index = next(i for i, text in enumerate(texts) if text == "范围")

        self.assertEqual(texts[scope_index - 1], "模板正文标题测试标准")
        self.assertEqual(self._jc_value(paragraphs[scope_index - 1]), "center")

    def _paragraphs(self, xml: str):
        root = ET.fromstring(xml)
        body = root.find("w:body", self._W_NS)
        return [p for p in body if p.tag == self._w_tag("p")]

    def _et_text(self, element) -> str:
        return "".join(t.text or "" for t in element.findall(".//w:t", self._W_NS))

    def _jc_value(self, paragraph) -> str:
        jc = paragraph.find("w:pPr/w:jc", self._W_NS)
        return jc.get(self._w_tag("val")) if jc is not None else ""

    def _w_tag(self, name: str) -> str:
        return "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + name


class CoverBackendDocxTest(unittest.TestCase):
    _W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    def test_cover_backend_omits_optional_sections_and_template_body_placeholders(self):
        xml = _build_cover_docx_xml(
            "---\n"
            "standard_type: 团体标准\n"
            "number: T/TEST 001-2026\n"
            "title: 样式后端测试标准\n"
            "title_en: Style backend test standard\n"
            "ics: \"27.010\"\n"
            "ccs: F 10\n"
            "publish_date: 2026-06-01\n"
            "implement_date: 2026-07-01\n"
            "publisher: 测试协会\n"
            "---\n"
            "# 范围\n\n"
            "正文。\n"
        )
        root = ET.fromstring(xml)
        body = root.find("w:body", self._W_NS)
        paragraphs = [p for p in body if p.tag == self._w_tag("p")]
        texts = [self._et_text(p) for p in paragraphs]

        self.assertNotIn("参考文献", texts)
        self.assertNotIn("索引", texts)
        self.assertNotIn("点击此处添加标准名称", xml)
        self.assertNotIn("章标题", xml)
        self.assertNotIn("条标题", xml)
        self.assertEqual(texts[-1], "")
        self.assertEqual(self._jc_value(paragraphs[-1]), "center")
        self.assertTrue(paragraphs[-1].findall(".//w:drawing", self._W_NS))

    def test_cover_backend_sections_and_body_standard_title_before_scope(self):
        xml = _build_cover_docx_xml(
            "---\n"
            "title: 封面正文标题测试标准\n"
            "introduction: |\n"
            "  引言内容。\n"
            "---\n"
            "# 范围\n\n"
            "正文。\n"
        )
        root = ET.fromstring(xml)
        body = root.find("w:body", self._W_NS)
        paragraphs = [p for p in body if p.tag == self._w_tag("p")]
        texts = [self._et_text(p) for p in paragraphs]
        scope_index = next(i for i, text in enumerate(texts) if text == "范围")
        title_index = scope_index - 1

        self.assertEqual(texts[title_index], "封面正文标题测试标准")
        self.assertEqual(self._jc_value(paragraphs[title_index]), "center")
        self.assertTrue(any(
            p.find("w:pPr/w:sectPr", self._W_NS) is not None
            for p in paragraphs[:title_index]
        ))
        self.assertGreaterEqual(sum(
            1 for p in paragraphs
            if p.find("w:pPr/w:sectPr", self._W_NS) is not None
        ), 4)

    def test_cover_backend_sections_before_body_appendix_references_and_index(self):
        xml = _build_cover_docx_xml(
            "---\n"
            "title: 分节测试标准\n"
            "---\n"
            "# 范围\n\n"
            "正文。\n\n"
            "# 附录 规范性 测试附录\n\n"
            "附录正文。\n\n"
            "# 附录 资料性 第二附录\n\n"
            "第二附录正文。\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n\n"
            "# 索引\n\n"
            "## B\n\n"
            "- 标准：1\n"
        )
        root = ET.fromstring(xml)
        body = root.find("w:body", self._W_NS)
        paragraphs = [p for p in body if p.tag == self._w_tag("p")]
        texts = [self._et_text(p) for p in paragraphs]

        marker_indices = {
            "正文": next(i for i, text in enumerate(texts) if text == "范围") - 1,
            "附录A": next(i for i, text in enumerate(texts) if "测试附录" in text),
            "附录B": next(i for i, text in enumerate(texts) if "第二附录" in text),
            "参考文献": next(i for i, text in enumerate(texts) if text == "参考文献"),
            "索引": next(i for i, text in enumerate(texts) if text == "索引"),
        }
        self.assertEqual(texts[marker_indices["正文"]], "分节测试标准")

        for marker, idx in marker_indices.items():
            with self.subTest(marker=marker):
                self.assertTrue(
                    any(
                        p.find("w:pPr/w:sectPr", self._W_NS) is not None
                        for p in paragraphs[:idx]
                    ),
                    marker,
                )
                prev_section = max(
                    i for i, p in enumerate(paragraphs[:idx])
                    if p.find("w:pPr/w:sectPr", self._W_NS) is not None
                )
                blocking_texts = [t for t in texts[prev_section + 1:idx] if t.strip()]
                self.assertEqual(blocking_texts, [], marker)

    def test_cover_backend_group_and_national_cover_metadata(self):
        cases = [
            (
                "group",
                "团体标准",
                "T/TEST 002-2026",
                "团体封面测试标准",
                "Group cover test standard",
                "测试协会",
            ),
            (
                "national",
                "国家标准",
                "GB/T 99999-2026",
                "国家封面测试标准",
                "National cover test standard",
                "国家市场监督管理总局 国家标准化管理委员会",
            ),
        ]
        for kind, standard_type, number, title, title_en, publisher in cases:
            with self.subTest(kind=kind):
                xml = _build_cover_docx_xml(
                    "---\n"
                    f"standard_type: {standard_type}\n"
                    f"number: {number}\n"
                    f"title: {title}\n"
                    f"title_en: {title_en}\n"
                    "ics: \"27.010\"\n"
                    "ccs: F 10\n"
                    "publish_date: 2026-06-01\n"
                    "implement_date: 2026-07-01\n"
                    f"publisher: {publisher}\n"
                    "---\n"
                    "# 范围\n\n"
                    "正文。\n",
                    kind=kind,
                )

                self.assertIn(number, xml)
                self.assertIn(title, xml)
                self.assertIn(title_en, xml)
                self.assertIn("27.010", xml)
                self.assertIn("F 10", xml)
                self.assertIn("2026-06-01发布", xml)
                self.assertIn("2026-07-01实施", xml)
                if kind == "national":
                    self.assertNotIn("国家市场监督管理总局", xml)
                    self.assertNotIn("国家标准化管理委员会", xml)
                    self.assertIn("国标发布单位", xml)
                else:
                    self.assertIn(publisher, xml)
                self.assertNotIn("点击此处添加标准名称", xml)

    def test_cover_backend_keeps_seq_groups_and_appendix_scope(self):
        xml = _build_cover_docx_xml(
            "# 范围\n\n"
            "见{{tbl:main:label}}，按{{eq:rate:label}}计算。\n\n"
            "{表：#tbl:main} 主表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| 正文 |\n\n"
            "$$x=1$${#eq:rate}\n\n"
            "# 附录 规范性 第一附录\n\n"
            "{表：#tbl:appA} 附录A表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| A |\n\n"
            "# 附录 资料性 第二附录\n\n"
            "{表：#tbl:appB} 附录B表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| B |\n"
        )

        self.assertIn(" SEQ 表 \\* ARABIC \\r 1 ", xml)
        self.assertIn(" SEQ 公式 \\* ARABIC \\r 1 ", xml)
        self.assertIn("<w:t>A.</w:t>", xml)
        self.assertIn("<w:t>B.</w:t>", xml)
        self.assertNotIn("SEQ TableA", xml)
        self.assertNotIn("SEQ TableB", xml)
        self.assertNotIn("SEQ Equation", xml)

    def test_cover_backend_uses_cover_blueprint_parts(self):
        parts = _build_cover_docx_parts("# 范围\n\n正文。\n", kind="group")
        with zipfile.ZipFile(os.path.join("templates", "cover_group.docx")) as zf:
            self.assertEqual(
                self._canonical_xml(parts["numbering"]),
                self._canonical_xml(zf.read("word/numbering.xml")),
            )
            self.assertEqual(self._style_names(parts["styles"]), self._style_names(zf.read("word/styles.xml")))
            self.assertNotEqual(parts["document"], zf.read("word/document.xml"))

    def test_cover_backend_end_line_uses_packaged_cover_image(self):
        cases = [
            ("group", os.path.join("templates", "cover_group.docx"), "word/media/image1.jpeg"),
            ("national", os.path.join("templates", "cover_national.docx"), "word/media/image3.jpg"),
        ]
        for kind, cover_path, expected_image in cases:
            with self.subTest(kind=kind):
                parts = _build_cover_docx_parts("# 范围\n\n正文。\n", kind=kind)
                embedded_image = self._last_paragraph_image_bytes(parts)
                with zipfile.ZipFile(cover_path) as zf:
                    self.assertEqual(embedded_image, zf.read(expected_image))

    def test_national_cover_does_not_overlay_publisher_text_on_publisher_image(self):
        xml = _build_cover_docx_xml(
            "---\n"
            "standard_type: 国家标准\n"
            "number: GB 15082—2026\n"
            "title: 汽车、摩托车用车速表\n"
            "title_en: Speedometer for motor vehicle and motorcycle\n"
            "publish_date: 2026-04-30\n"
            "implement_date: 2027-07-01\n"
            "publisher: |-\n"
            "  国家市场监督管理总局\n"
            "  国家标准化管理委员会\n"
            "---\n\n"
            "# 范围\n\n正文。\n",
            kind="national",
        )

        self.assertNotIn("国家市场监督管理总局", xml)
        self.assertNotIn("国家标准化管理委员会", xml)

    def test_docx_html_table_colspan_emits_gridspan(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "{表：#tbl:speed} 测试车速\n\n"
            "<table><tr><td>最高设计车速（<eq>v_{\\text{max}}</eq>）</td><td>测试车速</td></tr>"
            "<tr><td><eq>v_{\\text{max}}</eq>≤45</td><td>80%<eq>v_{\\text{max}}</eq></td></tr>"
            "<tr><td colspan=\"2\">注：按临近分度线取值。</td></tr></table>"
        )

        self.assertIn("<w:gridSpan w:val=\"2\"/>", xml)
        self.assertIn("<m:oMath", xml)
        self.assertIn("注：按临近分度线取值。", xml)
        self.assertIn("<w:vAlign w:val=\"center\"/>", xml)
        self.assertIn("<w:jc w:val=\"center\"/>", xml)
        self.assertIn("<w:jc w:val=\"left\"/>", xml)

    def test_docx_long_table_is_split_as_continued_table(self):
        rows = "".join("| %d | 值%d |\n" % (i, i) for i in range(1, 8))
        xml = _build_docx_xml(
            "# 范围\n\n"
            "{表：#tbl:long} 长表\n\n"
            "| 项 | 值 |\n"
            "| --- | --- |\n" +
            rows
        )

        self.assertEqual(xml.count(" SEQ 表 "), 1)
        self.assertIn("（续）", xml)
        self.assertIn(" REF _Ref", xml)
        self.assertGreaterEqual(xml.count("<w:tblHeader w:val=\"true\"/>"), 2)

    def test_cover_blueprints_keep_complete_cover_section(self):
        pairs = [
            (self._first_existing_path(
                os.path.join("templates", "团体标准模板.docx"),
                "2 团体标准——模板.docx",
            ), os.path.join("templates", "cover_group.docx")),
            (self._first_existing_path(
                os.path.join("templates", "国家标准模板.docx"),
                "国家标准.docx",
            ), os.path.join("templates", "cover_national.docx")),
        ]
        for source_path, cover_path in pairs:
            with self.subTest(cover_path=cover_path):
                self.assertEqual(
                    self._first_section_texts(source_path),
                    self._first_section_texts(cover_path),
                )
                self.assertNotIn("章标题", "".join(self._doc_texts(cover_path)))
                self.assertNotIn("条标题", "".join(self._doc_texts(cover_path)))

    def _first_existing_path(self, *paths: str) -> str:
        for path in paths:
            if os.path.exists(path):
                return path
        raise AssertionError("找不到测试资产：%s" % "；".join(paths))

    def _et_text(self, element) -> str:
        return "".join(t.text or "" for t in element.findall(".//w:t", self._W_NS))

    def _jc_value(self, paragraph) -> str:
        jc = paragraph.find("w:pPr/w:jc", self._W_NS)
        return jc.get(self._w_tag("val")) if jc is not None else ""

    def _w_tag(self, name: str) -> str:
        return "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + name

    def _style_names(self, styles_xml: bytes) -> set:
        root = ET.fromstring(styles_xml)
        return {
            name.get(self._w_tag("val"))
            for name in root.findall(".//w:style/w:name", self._W_NS)
            if name.get(self._w_tag("val"))
        }

    def _canonical_xml(self, xml: bytes) -> bytes:
        return ET.tostring(ET.fromstring(xml), encoding="utf-8")

    def _last_paragraph_image_bytes(self, parts: dict) -> bytes:
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        document = ET.fromstring(parts["document"])
        paragraphs = [p for p in document.find("w:body", self._W_NS) if p.tag == self._w_tag("p")]
        blip = paragraphs[-1].find(".//{%s}blip" % a_ns)
        self.assertIsNotNone(blip)
        rel_id = blip.get("{%s}embed" % r_ns)
        rels = ET.fromstring(parts["rels"])
        target = None
        for rel in rels.findall("{%s}Relationship" % rel_ns):
            if rel.get("Id") == rel_id:
                target = rel.get("Target")
                break
        self.assertIsNotNone(target)
        media_path = target if target.startswith("word/") else "word/" + target
        return parts["media"][media_path]

    def _doc_texts(self, path: str) -> list:
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        return [self._et_text(p) for p in root.findall(".//w:p", self._W_NS)]

    def _first_section_texts(self, path: str) -> list:
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        body = root.find("w:body", self._W_NS)
        texts = []
        for child in body:
            if child.tag != self._w_tag("p"):
                continue
            texts.append(self._et_text(child))
            if child.find("w:pPr/w:sectPr", self._W_NS) is not None:
                break
        return texts


if __name__ == "__main__":
    unittest.main()
