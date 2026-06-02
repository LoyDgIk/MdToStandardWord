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


if __name__ == "__main__":
    unittest.main()
