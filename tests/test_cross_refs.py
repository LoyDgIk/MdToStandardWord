# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import re
import tempfile
import unittest
import unittest.mock as mock
import zipfile
import xml.etree.ElementTree as ET

from md2std import cli, docx_builder, md_parser, model, resources, word_postprocess
from md2std.docx.oxml import _bm_name


def _build_docx_xml(markdown: str) -> str:
    sdoc = md_parser.parse(markdown)
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        docx_builder.build_cover(sdoc, path)
        with zipfile.ZipFile(path) as zf:
            return zf.read("word/document.xml").decode("utf-8", errors="ignore")
    finally:
        if os.path.exists(path):
            os.remove(path)


def _build_docx_parts(markdown: str) -> dict:
    sdoc = md_parser.parse(markdown)
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        docx_builder.build_cover(sdoc, path)
        with zipfile.ZipFile(path) as zf:
            return {
                "document": zf.read("word/document.xml"),
                "styles": zf.read("word/styles.xml"),
                "numbering": zf.read("word/numbering.xml"),
            }
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
                "headers": {
                    name: zf.read(name)
                    for name in zf.namelist()
                    if name.startswith("word/header") and name.endswith(".xml")
                },
                "footers": {
                    name: zf.read(name)
                    for name in zf.namelist()
                    if name.startswith("word/footer") and name.endswith(".xml")
                },
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

    def test_standard_reference_registration_supports_aliases_and_foreign_years(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "符合{{std:GB/T 1.1}}、{{std:JIS S 6006}}、"
            "{{std:ISO 3160-2}}和{{std:EN 71—3:2019}}。\n\n"
            "# 规范性引用文件\n\n"
            "{{std:GB/T 1.1—2020}} GB/T 1.1—2020  标准化工作导则\n\n"
            "{{std:JIS S 6006:2007}} JIS S 6006:2007  铅笔、彩色铅笔及其笔芯\n\n"
            "{{std:ISO 3160-2:2015}} ISO 3160-2:2015  表壳体及其附件 金合金覆盖层 第2部分:纯度、厚度、耐腐蚀性能和附着力的测试\n\n"
            "{{std:EN 71—3}} EN 71—3:2019  玩具安全 第3部分:特定元素的迁移\n"
        )

        scope = next(
            blk for blk in doc.body
            if isinstance(blk, model.Paragraph) and blk.text.startswith("符合")
        )
        refs = [
            sp.target
            for sp in scope.spans
            if isinstance(sp, model.RefSpan) and sp.ref_type == "std"
        ]
        self.assertEqual(
            ["GB/T 1.1—2020", "JIS S 6006:2007", "ISO 3160-2:2015", "EN 71—3"],
            refs,
        )

    def test_implicit_normative_reference_warns_but_matches(self):
        with self.assertWarnsRegex(UserWarning, "旧版自动识别"):
            md_parser.parse(
                "# 范围\n\n"
                "符合{{std:GB/T 11615}}。\n\n"
                "# 规范性引用文件\n\n"
                "GB/T 11615  地热资源地质勘查规范\n"
            )

    def test_domestic_standard_year_separator_warns_but_matches(self):
        with self.assertWarnsRegex(UserWarning, "年份连接号"):
            md_parser.parse(
                "# 范围\n\n"
                "符合{{std:GB/T 1.1—2020}}。\n\n"
                "# 规范性引用文件\n\n"
                "{{std:GB/T 1.1-2020}} GB/T 1.1-2020  标准化工作导则\n"
            )

    def test_extended_markdown_subscript_and_superscript_render(self):
        doc = md_parser.parse("# 范围\n\nH~2~O 的 2^10^ 倍，T<sub>r</sub> 与 m<sup>3</sup>。\n")
        paragraph = next(b for b in doc.body if isinstance(b, model.Paragraph))

        subscript_text = "".join(s.text for s in paragraph.spans if s.subscript)
        superscript_text = "".join(s.text for s in paragraph.spans if s.superscript)

        self.assertEqual(subscript_text, "2r")
        self.assertEqual(superscript_text, "103")

        xml = _build_docx_xml("# 范围\n\nH~2~O 的 2^10^ 倍。\n")
        self.assertIn('<w:vertAlign w:val="subscript"', xml)
        self.assertIn('<w:vertAlign w:val="superscript"', xml)

    def test_explicit_term_marker_normalizes_to_term_model(self):
        doc = md_parser.parse(
            "# 术语和定义\n\n"
            "{术语：地热温泉 | geothermal hot spring}\n\n"
            "出水温度不低于25 ℃的地下热水。\n\n"
            "注：用于资源开发利用语境。\n\n"
            "[来源：GB/T 11615—2010，3.1，有修改]\n"
        )

        term = next(b for b in doc.body if isinstance(b, model.Term))

        self.assertEqual(term.term, "地热温泉")
        self.assertEqual(term.term_en, "geothermal hot spring")
        self.assertEqual("".join(s.text for s in term.term_en_spans), "geothermal hot spring")
        self.assertFalse(any(s.italic for s in term.term_en_spans))
        self.assertEqual("".join(s.text for s in term.definition), "出水温度不低于25 ℃的地下热水。")
        self.assertEqual(len(term.notes), 1)
        self.assertEqual("".join(s.text for s in term.notes[0].spans), "用于资源开发利用语境。")
        self.assertIsNotNone(term.source)
        self.assertEqual(term.source.text, "[来源：GB/T 11615—2010，3.1，有修改]")

    def test_term_english_spans_preserve_markdown_italic(self):
        doc = md_parser.parse(
            "# 术语和定义\n\n"
            "{术语：大肠埃希氏菌 | *Escherichia coli*}\n\n"
            "一种常见指示菌。\n\n"
            "## 耐热大肠菌群  thermotolerant coliform bacteria\n\n"
            "在规定条件下可生长的菌群。\n"
        )

        terms = [b for b in doc.body if isinstance(b, model.Term)]

        self.assertEqual(terms[0].term_en, "Escherichia coli")
        self.assertEqual("".join(s.text for s in terms[0].term_en_spans), "Escherichia coli")
        self.assertTrue(any(s.italic for s in terms[0].term_en_spans))
        self.assertEqual(terms[1].term_en, "thermotolerant coliform bacteria")
        self.assertFalse(any(s.italic for s in terms[1].term_en_spans))

    def test_inline_double_dollar_formula_renders_as_omml(self):
        doc = md_parser.parse("# 范围\n\n变量 $$T_r$$ 与 $$Q_e$$ 应统一说明。\n")
        paragraph = next(b for b in doc.body if isinstance(b, model.Paragraph))
        formulas = [s for s in paragraph.spans if isinstance(s, model.FormulaSpan)]

        self.assertEqual([s.text for s in formulas], ["T_r", "Q_e"])

        xml = _build_docx_xml("# 范围\n\n变量 $$T_r$$ 与 H~2~O。\n")
        self.assertIn("<m:oMath", xml)
        self.assertNotIn("$$T_r$$", xml)
        self.assertIn('<w:vertAlign w:val="subscript"', xml)

    def test_note_and_example_markers_are_not_kept_in_content_spans(self):
        doc = md_parser.parse("# 范围\n\n注：温度分级用于初判。\n\n示例1：按代表性温度判定。\n")
        note = next(b for b in doc.body if isinstance(b, model.Note))
        example = next(b for b in doc.body if isinstance(b, model.Example))

        self.assertEqual("".join(s.text for s in note.spans), "温度分级用于初判。")
        self.assertIsNone(note.index)
        self.assertEqual("".join(s.text for s in example.spans), "按代表性温度判定。")
        self.assertEqual(example.index, 1)

        xml = _build_docx_xml("# 范围\n\n注：温度分级用于初判。\n")
        self.assertIn(">温度分级用于初判。<", xml)
        self.assertNotIn(">注：温度分级用于初判。<", xml)

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
            "<table><tr><td>最高设计车速（$$v_{\\text{max}}$$）</td>"
            "<td>测试车速</td></tr>"
            "<tr><td>$$v_{\\text{max}}$$≤45</td><td>80%</td></tr>"
            "<tr><td colspan=\"2\">注：按临近分度线取值。</td></tr></table>"
        )

        table = next(b for b in doc.body if isinstance(b, model.TableModel))
        self.assertEqual(table.header, ["最高设计车速（vmax）", "测试车速"])
        self.assertEqual(table.header_parts[0][1].kind, "formula")
        self.assertEqual(table.header_parts[0][1].text, "v_{\\text{max}}")
        self.assertEqual(table.rows[-1], ["注：按临近分度线取值。"])
        self.assertEqual(table.row_colspans[-1], [2])

    def test_html_table_rowspan_borders_empty_and_same_are_structured(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:merge} 合并边框表\n\n"
            "<table data-border-outer=\"thick\" data-border-inner=\"thin\">"
            "<tr><th rowspan=\"2\" data-border-right=\"thick\">类别</th><th colspan=\"2\">指标</th></tr>"
            "<tr><th>值</th><th data-border-bottom=\"none\">备注</th></tr>"
            "<tr><td data-align=\"left\">一类</td><td data-align=\"right\"></td><td data-align=\"decimal\">同上</td></tr>"
            "</table>"
        )

        table = next(b for b in doc.body if isinstance(b, model.TableModel))

        self.assertEqual(table.border_outer, "thick")
        self.assertEqual(table.border_inner, "thin")
        self.assertEqual(table.cell_rows[0][0].text, "类别")
        self.assertEqual(table.cell_rows[0][0].rowspan, 2)
        self.assertEqual(table.cell_rows[0][0].borders["right"], "thick")
        self.assertEqual(table.cell_rows[0][1].colspan, 2)
        self.assertEqual(table.cell_rows[1][1].borders["bottom"], "none")
        self.assertEqual(table.rows[-1], ["一类", "", "同上"])
        self.assertEqual(table.cell_rows[-1][1].text, "")
        self.assertEqual(table.cell_rows[-1][2].text, "同上")
        self.assertEqual(table.cell_rows[-1][0].align, "left")
        self.assertEqual(table.cell_rows[-1][1].align, "right")
        self.assertEqual(table.cell_rows[-1][2].align, "decimal")

    def test_table_cell_alignment_markers_are_structured(self):
        html_doc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:align} 对齐表\n\n"
            "<table><tr><th>项目</th><th>值</th></tr>"
            "<tr><td data-align=\"left\">说明</td><td data-align=\"right\">12.5</td></tr>"
            "</table>"
        )
        html_table = next(b for b in html_doc.body if isinstance(b, model.TableModel))
        self.assertEqual(html_table.cell_rows[1][0].align, "left")
        self.assertEqual(html_table.cell_rows[1][1].align, "right")

        gfm_doc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:gfm-align} GFM对齐表\n\n"
            "| 左 | 中 | 右 |\n"
            "| :--- | :---: | ---: |\n"
            "| A | B | C |\n"
        )
        gfm_table = next(b for b in gfm_doc.body if isinstance(b, model.TableModel))
        self.assertEqual([cell.align for cell in gfm_table.cell_rows[0]], ["left", "center", "right"])

    def test_html_table_rejects_invalid_spans_and_borders(self):
        with self.assertRaisesRegex(ValueError, "rowspan 必须为正整数"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:bad} 坏表\n\n"
                "<table><tr><td rowspan=\"0\">A</td></tr></table>"
            )
        with self.assertRaisesRegex(ValueError, "rowspan 超出表格行数"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:bad} 坏表\n\n"
                "<table><tr><td rowspan=\"2\">A</td></tr></table>"
            )
        with self.assertRaisesRegex(ValueError, "data-border-left 只支持"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:bad} 坏表\n\n"
                "<table><tr><td data-border-left=\"wide\">A</td></tr></table>"
            )
        with self.assertRaisesRegex(ValueError, "data-align 只支持"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:bad} 坏表\n\n"
                "<table><tr><td data-align=\"justify\">A</td></tr></table>"
            )

    def test_explicit_example_end_keeps_multi_block_example_content(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "示例：\n\n"
            "第一段示例内容。\n\n"
            "{表：#tbl:example} 示例表\n\n"
            "<table><tr><td rowspan=\"2\">A</td><td>1</td></tr><tr><td>同上</td></tr></table>\n\n"
            "第二段示例内容。\n\n"
            "{示例结束}\n"
        )

        body_types = [type(block).__name__ for block in doc.body]

        self.assertEqual(body_types, [
            "Heading", "Example", "ExampleContent", "TableModel", "ExampleContent",
        ])
        self.assertEqual(doc.body[2].text, "第一段示例内容。")
        self.assertEqual(doc.body[3].cell_rows[0][0].rowspan, 2)
        self.assertEqual(doc.body[4].text, "第二段示例内容。")

    def test_isolated_example_end_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "没有对应"):
            md_parser.parse("# 范围\n\n{示例结束}\n")

    def test_table_cell_footnote_ref_marker_is_structured(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:sample} 表题\n\n"
            "| 类型 | 内圆直径〔脚注〕 |\n"
            "| --- | --- |\n"
            "| A | 100 |\n\n"
            "{脚注} 表脚注的内容\n"
        )

        table = next(b for b in doc.body if isinstance(b, model.TableModel))
        self.assertEqual(table.header, ["类型", "内圆直径"])
        self.assertEqual(table.header_parts[1][-1].kind, "footnote_ref")
        self.assertEqual(table.header_parts[1][-1].text, "")
        self.assertEqual(table.footnotes[0].text, "表脚注的内容")

    def test_legacy_manual_footnote_labels_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "表格内脚注引用请写成"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:sample} 表题\n\n"
                "| 类型 | 内圆直径{脚注a} |\n"
                "| --- | --- |\n"
                "| A | 100 |\n"
            )
        with self.assertRaisesRegex(ValueError, "表格内脚注引用请写成"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:sample} 表题\n\n"
                "| 类型 | 内圆直径{脚注} |\n"
                "| --- | --- |\n"
                "| A | 100 |\n"
            )
        with self.assertRaisesRegex(ValueError, "图表附加项内容不要写进"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:sample} 表题\n\n"
                "| 类型 | 内圆直径 |\n"
                "| --- | --- |\n"
                "| A | 100 |\n\n"
                "{表脚注a：表脚注的内容}\n"
            )
        with self.assertRaisesRegex(ValueError, "图表附加项内容不要写进"):
            md_parser.parse(
                "# 范围\n\n"
                "![流程图 {#fig:flow}](missing.png)\n\n"
                "{图脚注a：图脚注的内容}\n"
            )

    def test_table_cell_note_markers_are_structured(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:sample} 表题\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| 段内容〔注：见{{tbl:sample:label}}，变量为$$v_{\\text{max}}$$。〕〔注：保留**强调**内容。〕 |\n"
        )

        table = next(b for b in doc.body if isinstance(b, model.TableModel))
        parts = table.row_parts[0][0]
        note_parts = [part for part in parts if part.kind == "note"]

        self.assertEqual(table.rows[0][0], "段内容")
        self.assertEqual(len(note_parts), 2)
        self.assertEqual("".join(s.text for s in note_parts[0].spans), "见{{tbl:sample:label}}，变量为v_{\\text{max}}。")
        self.assertTrue(any(isinstance(s, model.RefSpan) for s in note_parts[0].spans))
        self.assertTrue(any(isinstance(s, model.FormulaSpan) for s in note_parts[0].spans))
        self.assertEqual("".join(s.text for s in note_parts[1].spans), "保留强调内容。")
        self.assertEqual("".join(s.text for s in note_parts[1].spans if s.bold), "强调")

    def test_table_cell_note_marker_allows_standalone_notes_and_requires_continuity(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:sample} 表题\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| 〔注：单独注。〕 |\n"
        )

        table = next(b for b in doc.body if isinstance(b, model.TableModel))
        self.assertEqual(table.rows[0][0], "")
        self.assertEqual(len(table.row_parts[0][0]), 1)
        self.assertEqual(table.row_parts[0][0][0].kind, "note")
        self.assertEqual("".join(s.text for s in table.row_parts[0][0][0].spans), "单独注。")

        with self.assertRaisesRegex(ValueError, "内联普通注请写成"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:sample} 表题\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| {注：孤立注} |\n"
            )
        with self.assertRaisesRegex(ValueError, "表格单元格注必须连续写在被注释内容之后"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:sample} 表题\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| 段内容〔注：注内容〕后续文本 |\n"
            )

    def test_body_note_marker_is_regular_note_and_braced_note_is_rejected(self):
        doc = md_parser.parse("# 范围\n\n注：单条注。\n\n注：第二条注。\n")
        notes = [b for b in doc.body if isinstance(b, model.Note)]

        self.assertEqual(["".join(s.text for s in note.spans) for note in notes], ["单条注。", "第二条注。"])
        with self.assertRaisesRegex(ValueError, "正文注不要写进"):
            md_parser.parse("# 范围\n\n{注：旧写法。}\n")

    def test_table_addons_bind_to_previous_table_and_keep_cell_note_spans(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:classify} 温泉利用分类\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| 医疗保健〔注：见{{fig:flow:label}}。〕〔注：保留**强调**内容。〕 |\n\n"
            "{单位} 单位为{{fig:flow:label}}。\n\n"
            "{脚注} 表脚注见{{fig:flow:label}}。\n\n"
            "{来源} 资料来自{{fig:flow:label}}。\n\n"
            "![流程图 {#fig:flow}](missing.png)\n"
        )

        table = next(b for b in doc.body if isinstance(b, model.TableModel))
        note_parts = [part for part in table.row_parts[0][0] if part.kind == "note"]

        self.assertEqual(len(note_parts), 2)
        self.assertIsNotNone(table.unit)
        self.assertEqual("".join(s.text for s in table.unit.spans), "单位为{{fig:flow:label}}。")
        self.assertTrue(any(isinstance(s, model.RefSpan) for s in table.unit.spans))
        self.assertEqual("".join(s.text for s in note_parts[0].spans), "见{{fig:flow:label}}。")
        self.assertTrue(any(isinstance(s, model.RefSpan) for s in note_parts[0].spans))
        self.assertEqual("".join(s.text for s in note_parts[1].spans), "保留强调内容。")
        self.assertEqual("".join(s.text for s in note_parts[1].spans if s.bold), "强调")
        self.assertEqual(len(table.footnotes), 1)
        self.assertEqual("".join(s.text for s in table.footnotes[0].spans), "表脚注见{{fig:flow:label}}。")
        self.assertTrue(any(isinstance(s, model.RefSpan) for s in table.footnotes[0].spans))
        self.assertIsNotNone(table.source)
        self.assertTrue(any(isinstance(s, model.RefSpan) for s in table.source.spans))

    def test_figure_addons_bind_to_previous_figure(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "![流程图 {#fig:flow}](missing.png)\n\n"
            "{单位} 单位为{{tbl:classify:label}}。\n\n"
            "{图标引} 见{{tbl:classify:label}}。\n\n"
            "{图标引} 保留**强调**内容。\n\n"
            "{分图组:2}\n\n"
            "![分图a](missing-a.png)\n\n"
            "![分图b](missing-b.png)\n\n"
            "{图段} 段（可包含要求型条款）〔注：图中的注引用{{tbl:classify:label}}。〕\n\n"
            "{脚注} 图脚注。\n\n"
            "{来源} 资料来自项目组。\n\n"
            "{表：#tbl:classify} 温泉利用分类\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| 医疗保健 |\n"
        )

        figure = next(b for b in doc.body if isinstance(b, model.Figure))

        self.assertIsNotNone(figure.unit)
        self.assertEqual("".join(s.text for s in figure.unit.spans), "单位为{{tbl:classify:label}}。")
        self.assertTrue(any(isinstance(s, model.RefSpan) for s in figure.unit.spans))
        self.assertEqual([item.index for item in figure.key_items], ["1", "2"])
        self.assertEqual("".join(s.text for s in figure.key_items[0].spans), "见{{tbl:classify:label}}。")
        self.assertTrue(any(isinstance(s, model.RefSpan) for s in figure.key_items[0].spans))
        self.assertEqual("".join(s.text for s in figure.key_items[1].spans if s.bold), "强调")
        self.assertEqual(figure.subfigure_columns, 2)
        self.assertEqual([(item.path, item.caption) for item in figure.subfigures], [
            ("missing-a.png", "分图a"),
            ("missing-b.png", "分图b"),
        ])
        self.assertEqual(figure.body_paragraphs[0].text, "段（可包含要求型条款）")
        self.assertEqual("".join(s.text for s in figure.body_paragraphs[0].notes[0].spans), "图中的注引用{{tbl:classify:label}}。")
        self.assertTrue(any(isinstance(s, model.RefSpan) for s in figure.body_paragraphs[0].notes[0].spans))
        self.assertEqual(len(figure.footnotes), 1)
        self.assertEqual(figure.footnotes[0].text, "图脚注。")
        self.assertIsNotNone(figure.source)
        self.assertEqual(figure.source.text, "资料来自项目组。")

    def test_image_paths_resolve_from_markdown_file_location(self):
        source_path = os.path.abspath(os.path.join("examples", "relative-input.md"))
        doc = md_parser.parse(
            "# 范围\n\n"
            "![普通图 {#fig:single}](images/subfigure-a.png)\n\n"
            "{图：#fig:subparts} 组合图\n\n"
            "{分图组:2}\n\n"
            "![第一张分图题 ](images/subfigure-a.png)\n\n"
            "![第二张分图题](images/subfigure-b.png)\n",
            source_path=source_path,
        )
        figures = [b for b in doc.body if isinstance(b, model.Figure)]

        self.assertEqual(
            figures[0].path,
            os.path.abspath(os.path.join("examples", "images", "subfigure-a.png")),
        )
        self.assertEqual(figures[1].subfigure_columns, 2)
        self.assertEqual(
            [(os.path.basename(item.path), item.caption) for item in figures[1].subfigures],
            [("subfigure-a.png", "第一张分图题"), ("subfigure-b.png", "第二张分图题")],
        )
        self.assertEqual(
            figures[1].subfigures[0].path,
            os.path.abspath(os.path.join("examples", "images", "subfigure-a.png")),
        )

    def test_figure_table_addon_markers_require_matching_adjacent_target(self):
        with self.assertRaisesRegex(ValueError, "表注语法已移除"):
            md_parser.parse("# 范围\n\n{表注：孤立表注。}\n")
        with self.assertRaisesRegex(ValueError, "单位必须紧跟表格或图片"):
            md_parser.parse("# 范围\n\n{单位} 单位为毫米。\n")
        with self.assertRaisesRegex(ValueError, "脚注必须紧跟表格或图片"):
            md_parser.parse("# 范围\n\n{脚注} 孤立脚注。\n")
        with self.assertRaisesRegex(ValueError, "图表附加项内容不要写进"):
            md_parser.parse("# 范围\n\n{表单位：单位为毫米。}\n")
        with self.assertRaisesRegex(ValueError, "图注语法已移除"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:classify} 温泉利用分类\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| 医疗保健 |\n\n"
                "{图注：类型不匹配。}\n"
            )
        with self.assertRaisesRegex(ValueError, "图标引序号说明必须紧跟图片"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:classify} 温泉利用分类\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| 医疗保健 |\n\n"
                "{图标引} 类型不匹配。\n"
            )
        with self.assertRaisesRegex(ValueError, "图段必须紧跟图片"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:classify} 温泉利用分类\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| 医疗保健 |\n\n"
                "{图段} 类型不匹配。\n"
            )
        with self.assertRaisesRegex(ValueError, "分图组必须紧跟图片"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:classify} 温泉利用分类\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| 医疗保健 |\n\n"
                "{分图组:2}\n\n"
                "![类型不匹配](missing.png)\n"
            )
        with self.assertRaisesRegex(ValueError, "分图语法已改为分图组"):
            md_parser.parse(
                "# 范围\n\n"
                "![流程图 {#fig:flow}](missing.png)\n\n"
                "{分图：missing.png | 旧写法。}\n"
            )
        with self.assertRaisesRegex(ValueError, "分图组内容不要写进"):
            md_parser.parse(
                "# 范围\n\n"
                "![流程图 {#fig:flow}](missing.png)\n\n"
                "{分图组：2 | ![旧写法](missing.png)}\n"
            )
        with self.assertRaisesRegex(ValueError, "图标引不再手写编号"):
            md_parser.parse(
                "# 范围\n\n"
                "![流程图 {#fig:flow}](missing.png)\n\n"
                "{图标引1：旧写法。}\n"
            )
        with self.assertRaisesRegex(ValueError, "图段内容不要写进"):
            md_parser.parse(
                "# 范围\n\n"
                "![流程图 {#fig:flow}](missing.png)\n\n"
                "{图段：旧写法。}\n"
            )
        with self.assertRaisesRegex(ValueError, "表注语法已移除"):
            md_parser.parse(
                "# 范围\n\n"
                "![流程图 {#fig:flow}](missing.png)\n\n"
                "{表注：类型不匹配。}\n"
            )
        with self.assertRaisesRegex(ValueError, "单位只能写一次"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:classify} 温泉利用分类\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| 医疗保健 |\n\n"
                "{单位} 单位为毫米。\n\n"
                "{单位} 单位为厘米。\n"
            )

    def test_addon_cross_references_are_validated(self):
        with self.assertRaisesRegex(ValueError, "未知交叉引用"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:classify} 温泉利用分类\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| 医疗保健〔注：见{{fig:missing:label}}。〕 |\n"
            )
        with self.assertRaisesRegex(ValueError, "未知交叉引用"):
            md_parser.parse(
                "# 范围\n\n"
                "{表：#tbl:classify} 温泉利用分类\n\n"
                "| 项目 |\n"
                "| --- |\n"
                "| 医疗保健 |\n\n"
                "{单位} 单位为{{fig:missing:label}}。\n"
            )

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

    def test_untitled_clause_uses_explicit_level_marker(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "本文件规定了地热温泉资源开发利用要求。\n\n"
            "# 规范性引用文件\n\n"
            "本文件没有规范性引用文件。\n\n"
            "# 术语和定义\n\n"
            "本文件没有需要界定的术语和定义。\n\n"
            "# 要求\n\n"
            "## 一般要求\n\n"
            "{无标题条:3} 开发地热温泉资源前，应完成地质勘查。\n"
        )

        clause = next(b for b in doc.body if isinstance(b, model.UntitledClause))
        self.assertEqual(clause.level, 3)
        self.assertEqual(clause.text, "开发地热温泉资源前，应完成地质勘查。")

    def test_untitled_clause_is_rejected_in_foundational_chapters(self):
        cases = (
            ("范围", "# 范围\n\n{无标题条:2} 本文件规定了测试要求。\n"),
            ("规范性引用文件", "# 规范性引用文件\n\n{无标题条:2} GB/T 1.1  标准化工作导则\n"),
            ("术语和定义", "# 术语和定义\n\n{无标题条:2} 测试术语。\n"),
            ("符号和缩略语", "# 符号和缩略语\n\n{无标题条:2} A 为面积。\n"),
        )
        for chapter, text in cases:
            with self.subTest(chapter=chapter):
                with self.assertRaisesRegex(ValueError, r"\{无标题条:n\}.*普通段落"):
                    md_parser.parse(text)

    def test_legacy_untitled_clause_number_warns_and_stays_plain_paragraph(self):
        with self.assertWarnsRegex(UserWarning, "旧式无标题条手写编号"):
            doc = md_parser.parse("# 范围\n\n4.2.1  开发地热温泉资源前，应完成地质勘查。")

        paragraph = next(b for b in doc.body if isinstance(b, model.Paragraph))
        self.assertEqual(paragraph.text, "4.2.1  开发地热温泉资源前，应完成地质勘查。")

    def test_metadata_defaults_and_cover_options_parse(self):
        default_doc = md_parser.parse("# 范围\n\n正文。\n")
        enabled_doc = md_parser.parse(
            "---\n"
            "odd_even_pages: true\n"
            "cover_form_protection: true\n"
            "draft_version: 征求意见稿\n"
            "consistency_degree: MOD\n"
            "record_number: 1234-2026\n"
            "important_notice: 涉及人身安全的整体提示。\n"
            "symbols_lead: 下列符号适用于本文件。\n"
            "---\n\n"
            "# 范围\n\n正文。\n"
        )
        chinese_key_doc = md_parser.parse(
            "---\n"
            "草案版次: 报批稿\n"
            "---\n\n"
            "# 范围\n\n正文。\n"
        )

        self.assertFalse(default_doc.meta.odd_even_pages)
        self.assertFalse(default_doc.meta.cover_form_protection)
        self.assertEqual(default_doc.meta.draft_version, "")
        self.assertTrue(enabled_doc.meta.odd_even_pages)
        self.assertTrue(enabled_doc.meta.cover_form_protection)
        self.assertEqual(enabled_doc.meta.draft_version, "征求意见稿")
        self.assertEqual(enabled_doc.meta.consistency_degree, "MOD")
        self.assertEqual(enabled_doc.meta.record_number, "1234-2026")
        self.assertEqual(enabled_doc.meta.important_notice, "涉及人身安全的整体提示。")
        self.assertEqual(enabled_doc.meta.symbols_lead, "下列符号适用于本文件。")
        self.assertEqual(chinese_key_doc.meta.draft_version, "报批稿")

    def test_page_break_markers_parse_as_body_blocks(self):
        doc = md_parser.parse(
            "# 范围\n\n"
            "分页前。\n\n"
            "<!-- pagebreak -->\n\n"
            "分页后。\n\n"
            "\\pagebreak\n\n"
            "末段。\n"
        )

        self.assertEqual(
            [type(block) for block in doc.body],
            [
                model.Heading,
                model.Paragraph,
                model.PageBreak,
                model.Paragraph,
                model.PageBreak,
                model.Paragraph,
            ],
        )

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

    def test_normative_reference_registration_marker_is_not_emitted(self):
        xml = _build_cover_docx_xml(
            "---\n"
            "title: 规范性引用测试\n"
            "---\n\n"
            "# 范围\n\n"
            "应符合{{std:ISO 3160-2}}的规定。\n\n"
            "# 规范性引用文件\n\n"
            "{{std:ISO 3160-2:2015}} ISO 3160-2:2015  表壳体及其附件 金合金覆盖层 第2部分:纯度、厚度、耐腐蚀性能和附着力的测试\n"
        )

        self.assertNotIn("{{std:ISO 3160-2:2015}}", xml)
        self.assertIn("ISO 3160-2:2015", xml)
        bookmark = _bm_name("ISO 3160-2:2015")
        self.assertIn('w:name="%s"' % bookmark, xml)
        self.assertIn("REF %s" % bookmark, xml)

    def test_docx_uses_template_caption_numbering_styles(self):
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            "见{{tbl:main:label}}，完整题名为{{tbl:main:full}}，编号{{tbl:main}}，按{{eq:rate:label}}计算。\n\n"
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
            "![附录B图{#fig:appB}](missing.png)\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n"
        )
        xml = parts["document"].decode("utf-8", errors="ignore")
        root = ET.fromstring(parts["document"])
        styles = parts["styles"]

        self.assertNotIn("SEQ 表", xml)
        self.assertNotIn("SEQ 图", xml)
        self.assertIn(" SEQ 公式 \\* ARABIC \\r 1 ", xml)
        self.assertIn(" REF _Ref", xml)
        self.assertIn("\\h \\r", xml)
        self.assertIn("\\h \\n", xml)
        self.assertIn('w:name="_Ref', xml)

        body_table = self._et_paragraph_containing(root, "主表")
        body_figure = self._et_paragraph_containing(root, "主图")
        appendix_table = self._et_paragraph_containing(root, "附录B表")
        appendix_figure = self._et_paragraph_containing(root, "附录B图")
        self.assertEqual(
            self._paragraph_style(body_table),
            self._style_id_by_name(styles, "标准文件_正文表标题"),
        )
        self.assertEqual(
            self._paragraph_style(body_figure),
            self._style_id_by_name(styles, "标准文件_正文图标题"),
        )
        self.assertEqual(
            self._paragraph_style(appendix_table),
            self._style_id_by_name(styles, "标准文件_附录表标题"),
        )
        self.assertEqual(
            self._paragraph_style(appendix_figure),
            self._style_id_by_name(styles, "标准文件_附录图标题"),
        )
        for paragraph in (body_table, body_figure, appendix_table, appendix_figure):
            self.assertIsNotNone(paragraph.find("w:pPr/w:numPr", self._W_NS))
            self.assertNotEqual(self._num_id_value(paragraph), "0")
        self.assertEqual(
            self._numbering_level_format(parts, "标准文件_正文表标题"),
            ("表%1", "nothing"),
        )
        self.assertEqual(
            self._numbering_level_format(parts, "标准文件_正文图标题"),
            ("图%1", "nothing"),
        )
        self.assertEqual(
            self._numbering_level_format(parts, "标准文件_附录表标题"),
            ("表%1.%2", "nothing"),
        )
        self.assertEqual(
            self._numbering_level_format(parts, "标准文件_附录图标题"),
            ("图%1.%2", "nothing"),
        )

        appendix_table_label = self._style_id_by_name(styles, "标准文件_附录表标号")
        appendix_figure_label = self._style_id_by_name(styles, "标准文件_附录图标号")
        hidden_table_labels = self._et_paragraphs_with_style(root, appendix_table_label)
        hidden_figure_labels = self._et_paragraphs_with_style(root, appendix_figure_label)
        self.assertEqual(len(hidden_table_labels), 2)
        self.assertEqual(len(hidden_figure_labels), 2)
        for paragraph in hidden_table_labels + hidden_figure_labels:
            self.assertIsNotNone(paragraph.find("w:pPr/w:rPr/w:vanish", self._W_NS))
        self.assertNotIn("SEQ 表A", xml)
        self.assertNotIn("SEQ 表B", xml)
        self.assertNotIn("SEQ 图表", xml)
        self.assertNotIn("SEQ 公式A", xml)
        self.assertNotIn("SEQ 公式B", xml)
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

    def test_formula_visible_number_refs_do_not_inherit_hidden_formatting(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "按{{eq:rate:label}}计算。\n\n"
            "$$H=(T_r-T_0)/G+h$${#eq:rate}\n"
        )

        self.assertGreaterEqual(xml.count("\\* CHARFORMAT"), 2)
        self.assertIn(" REF _Ref", xml)
        self.assertIn("\\h \\* CHARFORMAT", xml)

    def test_term_chinese_and_english_terms_are_bold(self):
        parts = _build_docx_parts(
            "# 术语和定义\n\n"
            "## 地热温泉  geothermal hot spring\n\n"
            "出水温度不低于25 ℃的地下热水天然露头或人工揭露点。\n"
        )
        root = ET.fromstring(parts["document"])
        paragraphs = root.findall(".//w:p", self._W_NS)
        paragraph = self._et_paragraph_containing(root, "geothermal hot spring")
        self.assertIsNotNone(paragraph)
        term_index = paragraphs.index(paragraph)
        number_paragraph = paragraphs[term_index - 1]
        bold_texts = [
            "".join(t.text or "" for t in run.findall(".//w:t", self._W_NS))
            for run in paragraph.findall("w:r", self._W_NS)
            if run.find("w:rPr/w:b", self._W_NS) is not None
        ]

        self.assertEqual(self._et_text(number_paragraph), "")
        self.assertIsNotNone(number_paragraph.find("w:pPr/w:numPr", self._W_NS))
        self.assertEqual(
            self._paragraph_style(number_paragraph),
            self._style_id_by_name(parts["styles"], "标准文件_术语条一"),
        )
        self.assertIsNone(paragraph.find("w:pPr/w:numPr", self._W_NS))
        self.assertEqual(
            self._paragraph_style(paragraph),
            self._style_id_by_name(parts["styles"], "标准文件_段"),
        )
        self.assertIn("地热温泉", bold_texts)
        self.assertIn("geothermal hot spring", bold_texts)
        english_run = next(
            run for run in paragraph.findall("w:r", self._W_NS)
            if "".join(t.text or "" for t in run.findall(".//w:t", self._W_NS)) == "geothermal hot spring"
        )
        self.assertIsNone(english_run.find("w:rPr/w:i", self._W_NS))

    def test_term_latin_scientific_name_is_bold_and_italic(self):
        parts = _build_docx_parts(
            "# 术语和定义\n\n"
            "{术语：大肠埃希氏菌 | *Escherichia coli*}\n\n"
            "一种常见指示菌。\n\n"
            "{术语：耐热大肠菌群 | thermotolerant coliform bacteria}\n\n"
            "在规定条件下可生长的菌群。\n"
        )
        root = ET.fromstring(parts["document"])
        latin_paragraph = self._et_paragraph_containing(root, "Escherichia coli")
        ordinary_paragraph = self._et_paragraph_containing(root, "thermotolerant coliform bacteria")

        latin_run = next(
            run for run in latin_paragraph.findall("w:r", self._W_NS)
            if "".join(t.text or "" for t in run.findall(".//w:t", self._W_NS)) == "Escherichia coli"
        )
        ordinary_run = next(
            run for run in ordinary_paragraph.findall("w:r", self._W_NS)
            if "".join(t.text or "" for t in run.findall(".//w:t", self._W_NS)) == "thermotolerant coliform bacteria"
        )

        self.assertIsNotNone(latin_run.find("w:rPr/w:b", self._W_NS))
        self.assertIsNotNone(latin_run.find("w:rPr/w:i", self._W_NS))
        self.assertIsNotNone(ordinary_run.find("w:rPr/w:b", self._W_NS))
        self.assertIsNone(ordinary_run.find("w:rPr/w:i", self._W_NS))

    def test_explicit_term_marker_emits_term_definition_note_and_source(self):
        xml = _build_docx_xml(
            "# 术语和定义\n\n"
            "{术语：地热温泉 | geothermal hot spring}\n\n"
            "出水温度不低于25 ℃的地下热水。\n\n"
            "注：用于资源开发利用语境。\n\n"
            "[来源：GB/T 11615—2010，3.1，有修改]\n"
        )
        root = ET.fromstring(xml)
        texts = [self._et_text(p) for p in root.findall(".//w:p", self._W_NS)]

        lead_index = texts.index("下列术语和定义适用于本文件。")
        term_index = next(i for i, text in enumerate(texts) if "geothermal hot spring" in text)
        definition_index = texts.index("出水温度不低于25 ℃的地下热水。")
        note_index = texts.index("用于资源开发利用语境。")
        source_index = texts.index("[来源：GB/T 11615—2010，3.1，有修改]")

        self.assertLess(lead_index, term_index)
        self.assertLess(term_index, definition_index)
        self.assertLess(definition_index, note_index)
        self.assertLess(note_index, source_index)

        term_paragraph = self._et_paragraph_containing(root, "geothermal hot spring")
        bold_texts = [
            "".join(t.text or "" for t in run.findall(".//w:t", self._W_NS))
            for run in term_paragraph.findall("w:r", self._W_NS)
            if run.find("w:rPr/w:b", self._W_NS) is not None
        ]
        self.assertIn("地热温泉", bold_texts)
        self.assertIn("geothermal hot spring", bold_texts)
        source_paragraph = self._et_paragraph_containing(root, "GB/T 11615")
        self.assertIsNotNone(source_paragraph.find(".//w:rStyle", self._W_NS))

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

    def test_table_addons_emit_after_table_with_declared_styles(self):
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            "![说明图 {#fig:flow}](missing.png)\n\n"
            "{表：#tbl:main} 测试表\n\n"
            "| 项目 | 说明〔脚注〕 |\n"
            "| --- | --- |\n"
            "| A | 长文本说明应左对齐。〔注：单元格注引用{{fig:flow:label}}。〕〔注：第二条注。〕 |\n\n"
            "{单位} 单位为毫米\n\n"
            "{脚注} 表脚注的内容。\n\n"
            "{来源} 表资料来源。\n"
        )
        xml = parts["document"].decode("utf-8", errors="ignore")
        root = ET.fromstring(parts["document"])

        caption_pos = xml.index("测试表")
        unit_pos = xml.index("单位为毫米")
        table_pos = xml.index("<w:tbl", caption_pos)
        table_end = xml.index("</w:tbl>", table_pos)
        note_pos = xml.index("单元格注引用")
        second_note_pos = xml.index("第二条注")
        footnote_pos = xml.index("表脚注的内容")
        source_pos = xml.index("表资料来源")

        self.assertLess(caption_pos, unit_pos)
        self.assertLess(unit_pos, table_pos)
        self.assertLess(table_pos, note_pos)
        self.assertLess(note_pos, second_note_pos)
        self.assertLess(second_note_pos, footnote_pos)
        self.assertLess(footnote_pos, source_pos)
        self.assertLess(source_pos, table_end)
        self.assertIn('<w:gridSpan w:val="2"', xml[table_pos:table_end])
        caption_para = self._et_paragraph_containing(root, "测试表")
        unit_para = self._et_paragraph_containing(root, "单位为毫米")
        note_para = self._et_paragraph_containing(root, "单元格注引用")
        footnote_para = self._et_paragraph_containing(root, "表脚注的内容")
        source_para = self._et_paragraph_containing(root, "表资料来源")
        note_style = self._style_id_by_name(parts["styles"], "标准文件_注×：")
        footnote_style = self._style_id_by_name(parts["styles"], "标准文件_图表脚注")
        footnote_content_style = self._style_id_by_name(parts["styles"], "标准文件_图表脚注内容")
        source_style = self._style_id_by_name(parts["styles"], "标准文件_图表说明")
        note_row = self._et_table_row_containing(root, "单元格注引用")
        footnote_row = self._et_table_row_containing(root, "表脚注的内容")
        source_row = self._et_table_row_containing(root, "表资料来源")

        self.assertIsNotNone(caption_para.find("w:pPr/w:keepNext", self._W_NS))
        self.assertIsNotNone(unit_para.find("w:pPr/w:keepNext", self._W_NS))
        self.assertEqual(self._jc_value(unit_para), "right")
        self.assertIn("第二条注", self._et_text(note_row))
        self.assertIsNot(note_row, source_row)
        self.assertIsNot(note_row, footnote_row)
        self.assertIsNot(footnote_row, source_row)
        self.assertNotIn("〔脚注〕", xml[table_pos:table_end])
        self.assertEqual(self._paragraph_style(note_para), note_style)
        self.assertEqual(self._paragraph_style(footnote_para), footnote_style)
        self.assertEqual(self._paragraph_style(source_para), source_style)
        self.assertIsNone(note_para.find(f'.//w:rStyle[@w:val="{footnote_content_style}"]', self._W_NS))
        self.assertIsNone(footnote_para.find(f'.//w:rStyle[@w:val="{footnote_content_style}"]', self._W_NS))
        self.assertIn(f'w:rStyle w:val="{footnote_content_style}"', xml[table_pos:table_end])
        self.assertNotIn("注 1：", self._et_text(note_row))
        self.assertNotIn("注 2：", self._et_text(note_row))
        self.assertNotIn("a 表脚注", self._et_text(footnote_row))
        self.assertIn(" REF _Ref", xml[xml.rfind("<w:p", 0, note_pos):xml.index("</w:p>", note_pos)])

    def test_figure_addons_emit_after_caption_with_declared_styles(self):
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            "{表：#tbl:main} 引用表\n\n"
            "| 项目 |\n"
            "| --- |\n"
            "| A |\n\n"
            "![测试图 {#fig:flow}](missing.png)\n\n"
            "{单位} 单位为毫米\n\n"
            "{图标引} 说明的内容\n\n"
            "{图标引} 说明的内容\n\n"
            "{图段} 段（可包含要求型条款）〔注：图中的注的内容〕\n\n"
            "{脚注} 图脚注的内容。\n\n"
            "{来源} 图资料来源。\n"
        )
        xml = parts["document"].decode("utf-8", errors="ignore")
        root = ET.fromstring(parts["document"])

        unit_pos = xml.index("单位为毫米")
        image_pos = xml.index("[缺少图片：missing.png]")
        key_lead_pos = xml.index("标引序号说明")
        key_item_pos = xml.index("说明的内容", key_lead_pos)
        body_para_pos = xml.index("段（可包含要求型条款）")
        note_pos = xml.index("图中的注的内容")
        footnote_pos = xml.index("图脚注的内容")
        source_pos = xml.index("图资料来源")
        caption_pos = xml.index("测试图")

        self.assertLess(unit_pos, image_pos)
        self.assertLess(image_pos, key_lead_pos)
        self.assertLess(key_lead_pos, key_item_pos)
        self.assertLess(key_item_pos, body_para_pos)
        self.assertLess(body_para_pos, note_pos)
        self.assertLess(note_pos, footnote_pos)
        self.assertLess(footnote_pos, source_pos)
        self.assertLess(source_pos, caption_pos)
        unit_para = self._et_paragraph_containing(root, "单位为毫米")
        key_lead_para = self._et_paragraph_containing(root, "标引序号说明")
        body_para = self._et_paragraph_containing(root, "段（可包含要求型条款）")
        note_para = self._et_paragraph_containing(root, "图中的注的内容")
        footnote_para = self._et_paragraph_containing(root, "图脚注的内容")
        source_para = self._et_paragraph_containing(root, "图资料来源")
        para_style = self._style_id_by_name(parts["styles"], "标准文件_段")
        note_style = self._style_id_by_name(parts["styles"], "标准文件_注：")
        footnote_style = self._style_id_by_name(parts["styles"], "标准文件_图表脚注")
        footnote_content_style = self._style_id_by_name(parts["styles"], "标准文件_图表脚注内容")
        source_style = self._style_id_by_name(parts["styles"], "标准文件_图表说明")

        self.assertEqual(self._jc_value(unit_para), "right")
        self.assertIsNotNone(unit_para.find("w:pPr/w:keepNext", self._W_NS))
        self.assertEqual(self._paragraph_style(key_lead_para), para_style)
        self.assertEqual(self._paragraph_style(body_para), para_style)
        self.assertEqual(self._paragraph_style(note_para), note_style)
        self.assertEqual(self._paragraph_style(footnote_para), footnote_style)
        self.assertEqual(self._paragraph_style(source_para), source_style)
        self.assertIsNone(footnote_para.find(f'.//w:rStyle[@w:val="{footnote_content_style}"]', self._W_NS))
        self.assertIsNotNone(source_para.find("w:pPr/w:keepNext", self._W_NS))

    def test_figure_image_is_constrained_to_body_width(self):
        image_path = os.path.abspath(os.path.join("examples", "images", "subfigure-a.png"))
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            f"![宽图 {{#fig:wide}}]({image_path})\n"
        )
        root = ET.fromstring(parts["document"])
        inline = root.find(".//wp:inline", {
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        })
        self.assertIsNotNone(inline)
        extent = inline.find("wp:extent", {
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        })
        self.assertIsNotNone(extent)
        self.assertLessEqual(int(extent.get("cx")), 6120130)
        image_para = self._et_paragraph_with_drawing(root)
        spacing = image_para.find("w:pPr/w:spacing", self._W_NS)
        self.assertIsNotNone(spacing)
        self.assertEqual(spacing.get(self._w_tag("lineRule")), "auto")

    def test_subfigures_are_composed_from_separate_images(self):
        first = os.path.abspath(os.path.join("examples", "images", "subfigure-a.png"))
        second = os.path.abspath(os.path.join("examples", "images", "subfigure-b.png"))
        first_md = first.replace("\\", "/")
        second_md = second.replace("\\", "/")
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            "{图：#fig:subparts} 组合图\n\n"
            "{分图组:2}\n\n"
            f"![分图题一]({first_md})\n\n"
            f"![分图题二]({second_md})\n"
        )
        xml = parts["document"].decode("utf-8", errors="ignore")
        root = ET.fromstring(parts["document"])
        wp_ns = {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
        inlines = root.findall(".//wp:inline", wp_ns)

        self.assertNotIn("[缺少图片：", xml)
        self.assertGreaterEqual(len(inlines), 2)
        self.assertIn("a)　分图题一", xml)
        self.assertIn("b)　分图题二", xml)

    def test_single_table_cell_note_uses_single_note_style(self):
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            "{表：#tbl:single-note} 单注表\n\n"
            "<table><tr><td colspan=\"2\">段（可包含要求型条款）〔注：单条注含变量$$v_{\\text{max}}$$。〕</td></tr></table>\n"
        )
        root = ET.fromstring(parts["document"])
        body_para = self._et_paragraph_containing(root, "段（可包含要求型条款）")
        note_para = self._et_paragraph_containing(root, "单条注含变量")
        xml = parts["document"].decode("utf-8", errors="ignore")

        self.assertEqual(
            self._paragraph_style(body_para),
            self._style_id_by_name(parts["styles"], "标准文件_段"),
        )
        self.assertEqual(
            self._paragraph_style(note_para),
            self._style_id_by_name(parts["styles"], "标准文件_注："),
        )
        self.assertIn("<m:oMath", xml)
        self.assertNotIn("v_{\\text{max}}", xml)
        self.assertIsNone(note_para.find("w:pPr/w:numPr", self._W_NS))
        self.assertEqual(self._ind_value(body_para, "firstLine"), "199")
        self.assertEqual(self._ind_value(body_para, "firstLineChars"), "95")

    def test_official_table_sample_keeps_body_row_and_notes_in_one_cell(self):
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            "{表：#tbl:sample} 表题\n\n"
            "<table><tr><th>类型</th><th>长度</th><th>内圆直径〔脚注〕</th><th>外圆直径</th></tr>"
            "<tr><td>A</td><td>230</td><td>100</td><td>125</td></tr>"
            "<tr><td>……</td><td>……</td><td>……</td><td>……</td></tr>"
            "<tr><td colspan=\"4\">段（可包含要求型条款）"
            "〔注：表中的注的内容〕〔注：表中的注的内容〕</td></tr></table>\n\n"
            "{单位} 单位为毫米\n\n"
            "{脚注} 表脚注的内容\n"
        )
        xml = parts["document"].decode("utf-8", errors="ignore")
        root = ET.fromstring(parts["document"])
        table_pos = xml.index("<w:tbl", xml.index("表题"))
        table_end = xml.index("</w:tbl>", table_pos)
        table_xml = xml[table_pos:table_end]

        self.assertNotIn("〔脚注〕", table_xml)
        self.assertEqual(table_xml.count("<w:tr>"), 5)
        body_note_row = self._et_table_row_containing(root, "段（可包含要求型条款）")
        footnote_row = self._et_table_row_containing(root, "表脚注的内容")
        body_para = self._et_paragraph_containing(root, "段（可包含要求型条款）")
        first_note_para = self._et_paragraph_containing(root, "表中的注的内容")
        body_style = self._style_id_by_name(parts["styles"], "标准文件_段")
        note_style = self._style_id_by_name(parts["styles"], "标准文件_注×：")
        footnote_style = self._style_id_by_name(parts["styles"], "标准文件_图表脚注")

        self.assertEqual(self._paragraph_style(body_para), body_style)
        self.assertEqual(self._paragraph_style(first_note_para), note_style)
        self.assertEqual(self._ind_value(body_para, "firstLine"), "199")
        self.assertEqual(self._ind_value(body_para, "firstLineChars"), "95")
        self.assertNotIn("注 1：", self._et_text(body_note_row))
        self.assertNotIn("注 2：", self._et_text(body_note_row))
        self.assertNotIn("表脚注的内容", self._et_text(body_note_row))
        self.assertIsNot(body_note_row, footnote_row)
        footnote_para = self._et_paragraph_containing(root, "表脚注的内容")
        self.assertEqual(self._paragraph_style(footnote_para), footnote_style)

    def test_standalone_table_cell_note_does_not_emit_empty_body_paragraph(self):
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            "{表：#tbl:standalone-note} 单独注表\n\n"
            "<table><tr><td colspan=\"2\">〔注：单独注的内容〕</td></tr></table>\n"
        )
        root = ET.fromstring(parts["document"])
        row = self._et_table_row_containing(root, "单独注的内容")
        paras = row.findall(".//w:p", self._W_NS)
        note_para = self._et_paragraph_containing(root, "单独注的内容")
        note_style = self._style_id_by_name(parts["styles"], "标准文件_注：")

        self.assertEqual(len(paras), 1)
        self.assertEqual(self._paragraph_style(note_para), note_style)
        self.assertEqual(self._et_text(row), "单独注的内容")

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

    def test_body_page_break_marker_emits_real_word_page_break(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "分页前。\n\n"
            "<!-- pagebreak -->\n\n"
            "分页后。\n"
        )

        self.assertIn('<w:br w:type="page"/>', xml)

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

    def test_caption_bookmarks_attach_to_style_numbered_paragraphs(self):
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

        for anchor_id, title in (("main", "正文表"), ("appB", "附录B表")):
            para = self._paragraph_containing(xml, title)
            label_name = docx_builder._native_ref_name("tbl", anchor_id, "label")
            num_name = docx_builder._native_ref_name("tbl", anchor_id, "num")
            full_name = docx_builder._native_ref_name("tbl", anchor_id, "full")
            text_name = docx_builder._native_ref_name("tbl", anchor_id, "text")
            label_id = self._bookmark_id(para, label_name)
            num_id = self._bookmark_id(para, num_name)
            full_id = self._bookmark_id(para, full_name)
            text_id = self._bookmark_id(para, text_name)

            self.assertNotIn("SEQ 表", para)
            self.assertNotIn("<w:t>表</w:t>", para)
            self.assertIn("<w:numPr>", para)
            title_pos = para.index(f"<w:t>{title}</w:t>")
            for bid in (label_id, num_id, full_id, text_id):
                self.assertGreater(para.index(f'<w:bookmarkEnd w:id="{bid}"/>'), title_pos)

    def test_nested_ordered_list_uses_template_multilevel_numbering(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "1. 车辆装备：\n"
            "   1. 为整备质量再加上驾驶员。\n"
            "   2. 所配轮胎气压为正常行驶用气压。\n"
        )
        root = ET.fromstring(xml)
        top = self._et_paragraph_containing(root, "车辆装备")
        nested = self._et_paragraph_containing(root, "整备质量")

        self.assertIsNotNone(top)
        self.assertIsNotNone(nested)
        self.assertIsNotNone(top.find('w:pPr/w:pStyle[@w:val="174"]', self._W_NS))
        self.assertIsNotNone(top.find('w:pPr/w:numPr/w:ilvl[@w:val="0"]', self._W_NS))
        self.assertIsNotNone(nested.find('w:pPr/w:pStyle[@w:val="109"]', self._W_NS))
        self.assertIsNotNone(nested.find('w:pPr/w:numPr/w:ilvl[@w:val="1"]', self._W_NS))
        self.assertEqual(self._num_id_value(top), self._num_id_value(nested))

    def test_separate_ordered_lists_restart_numbering(self):
        parts = _build_docx_parts(
            "# 范围\n\n"
            "## 发生异常时\n\n"
            "1. 停止运行。\n"
            "2. 设置警戒区域。\n\n"
            "## 解除异常后\n\n"
            "1. 自动通风。\n"
            "2. 恢复确认。\n"
        )
        document = ET.fromstring(parts["document"])
        numbering = ET.fromstring(parts["numbering"])
        first = self._et_paragraph_containing(document, "停止运行")
        second = self._et_paragraph_containing(document, "自动通风")

        self.assertIsNotNone(first)
        self.assertIsNotNone(second)
        first_num_id = self._num_id_value(first)
        second_num_id = self._num_id_value(second)
        self.assertNotEqual(first_num_id, second_num_id)
        for num_id in (first_num_id, second_num_id):
            restart = numbering.find(
                f'w:num[@w:numId="{num_id}"]/w:lvlOverride[@w:ilvl="0"]/w:startOverride[@w:val="1"]',
                self._W_NS,
            )
            self.assertIsNotNone(restart)

    def test_body_basic_elements_emit_notice_and_default_leads(self):
        xml = _build_cover_docx_xml(
            "---\n"
            "title: 基础要素测试标准\n"
            "important_notice: 涉及人身安全的整体提示。\n"
            "symbols_lead: 下列符号适用于本文件。\n"
            "---\n\n"
            "# 范围\n\n"
            "正文。\n\n"
            "# 规范性引用文件\n\n"
            "{{std:GB/T 1.1}} GB/T 1.1  标准化工作导则\n\n"
            "# 术语和定义\n\n"
            "## 地热资源  geothermal resources\n\n"
            "赋存于地球内部的热能资源。\n\n"
            "# 符号和缩略语\n\n"
            "A —— 面积。\n"
        )
        root = ET.fromstring(xml)
        texts = [self._et_text(p) for p in root.findall(".//w:p", self._W_NS)]

        title_index = texts.index("基础要素测试标准")
        notice_index = texts.index("重要提示：涉及人身安全的整体提示。")
        scope_index = texts.index("范围")

        self.assertLess(title_index, notice_index)
        self.assertLess(notice_index, scope_index)
        self.assertIn(
            "下列文件中的内容通过文中的规范性引用而构成本文件必不可少的条款。"
            "其中，注日期的引用文件，仅该日期对应的版本适用于本文件；"
            "不注日期的引用文件，其最新版本（包括所有的修改单）适用于本文件。",
            texts,
        )
        self.assertIn("下列术语和定义适用于本文件。", texts)
        self.assertIn("下列符号适用于本文件。", texts)

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

    def _et_paragraphs_with_style(self, root, style_id: str):
        paragraphs = []
        for paragraph in root.findall(".//w:p", self._W_NS):
            pstyle = paragraph.find("w:pPr/w:pStyle", self._W_NS)
            if pstyle is not None and pstyle.get(self._w_tag("val")) == style_id:
                paragraphs.append(paragraph)
        return paragraphs

    def _et_table_row_containing(self, root, text: str):
        for row in root.findall(".//w:tr", self._W_NS):
            if text in self._et_text(row):
                return row
        return None

    def _et_paragraph_with_drawing(self, root):
        for paragraph in root.findall(".//w:p", self._W_NS):
            if paragraph.find(".//w:drawing", self._W_NS) is not None:
                return paragraph
        return None

    def _et_text(self, element) -> str:
        return "".join(t.text or "" for t in element.findall(".//w:t", self._W_NS))

    def _num_id_value(self, paragraph) -> str:
        num_id = paragraph.find("w:pPr/w:numPr/w:numId", self._W_NS)
        self.assertIsNotNone(num_id)
        return num_id.get(self._w_tag("val"))

    def _paragraph_style(self, paragraph) -> str:
        pstyle = paragraph.find("w:pPr/w:pStyle", self._W_NS)
        self.assertIsNotNone(pstyle)
        return pstyle.get(self._w_tag("val"))

    def _style_id_by_name(self, styles_xml: bytes, name: str) -> str:
        root = ET.fromstring(styles_xml)
        for style in root.findall("w:style", self._W_NS):
            name_el = style.find("w:name", self._W_NS)
            if name_el is not None and name_el.get(self._w_tag("val")) == name:
                return style.get(self._w_tag("styleId"))
        self.fail("找不到样式：%s" % name)

    def _numbering_level_format(self, parts: dict, style_name: str) -> tuple:
        styles_root = ET.fromstring(parts["styles"])
        numbering_root = ET.fromstring(parts["numbering"])
        style_num_id = None
        style_ilvl = "0"
        for style in styles_root.findall("w:style", self._W_NS):
            name_el = style.find("w:name", self._W_NS)
            if name_el is None or name_el.get(self._w_tag("val")) != style_name:
                continue
            num_id = style.find("w:pPr/w:numPr/w:numId", self._W_NS)
            ilvl = style.find("w:pPr/w:numPr/w:ilvl", self._W_NS)
            self.assertIsNotNone(num_id)
            style_num_id = num_id.get(self._w_tag("val"))
            style_ilvl = ilvl.get(self._w_tag("val")) if ilvl is not None else "0"
            break
        self.assertIsNotNone(style_num_id)
        abstract_id = None
        for num in numbering_root.findall("w:num", self._W_NS):
            if num.get(self._w_tag("numId")) != style_num_id:
                continue
            abstract = num.find("w:abstractNumId", self._W_NS)
            self.assertIsNotNone(abstract)
            abstract_id = abstract.get(self._w_tag("val"))
            break
        self.assertIsNotNone(abstract_id)
        for abstract in numbering_root.findall("w:abstractNum", self._W_NS):
            if abstract.get(self._w_tag("abstractNumId")) != abstract_id:
                continue
            for level in abstract.findall("w:lvl", self._W_NS):
                if level.get(self._w_tag("ilvl")) != style_ilvl:
                    continue
                lvl_text = level.find("w:lvlText", self._W_NS)
                suff = level.find("w:suff", self._W_NS)
                return (
                    lvl_text.get(self._w_tag("val")) if lvl_text is not None else "",
                    suff.get(self._w_tag("val")) if suff is not None else "",
                )
        self.fail("找不到编号级别：%s" % style_name)

    def _jc_value(self, paragraph) -> str:
        jc = paragraph.find("w:pPr/w:jc", self._W_NS)
        return jc.get(self._w_tag("val")) if jc is not None else ""

    def _ind_value(self, paragraph, name: str) -> str:
        ind = paragraph.find("w:pPr/w:ind", self._W_NS)
        return ind.get(self._w_tag(name)) if ind is not None else ""

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

    def test_cover_backend_section_titles_match_template_direct_formatting(self):
        xml = _build_cover_docx_xml(
            "---\n"
            "title: 标题格式测试标准\n"
            "introduction: |\n"
            "  引言内容。\n"
            "---\n"
            "# 范围\n\n"
            "正文。\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n\n"
            "# 索引\n\n"
            "## B\n\n"
            "- 标准：1\n",
            kind="group",
        )
        root = ET.fromstring(xml)

        title_specs = {
            "目次": ("360", [("目", "320"), ("次", None)]),
            "前言": ("360", [("前", "320"), ("言", None)]),
            "引言": ("360", [("引", "320"), ("言", None)]),
            "参考文献": ("120", [("参考文", "105"), ("献", None)]),
            "索引": ("120", [("索", "210"), ("引", None)]),
        }
        for title, (after, runs) in title_specs.items():
            with self.subTest(title=title):
                paragraph = self._et_paragraph_exact(root, title)
                self.assertIsNotNone(paragraph)
                spacing = paragraph.find("w:pPr/w:spacing", self._W_NS)
                self.assertIsNotNone(spacing)
                self.assertEqual(spacing.get(self._w_tag("after")), after)
                actual_runs = paragraph.findall("w:r", self._W_NS)
                self.assertEqual(
                    [
                        (
                            self._et_text(run),
                            self._run_spacing(run),
                        )
                        for run in actual_runs
                        if self._et_text(run)
                    ],
                    runs,
                )

    def test_national_cover_backend_section_titles_match_template_spacing(self):
        xml = _build_cover_docx_xml(
            "---\n"
            "standard_type: 国家标准\n"
            "title: 国家标题格式测试标准\n"
            "introduction: |\n"
            "  引言内容。\n"
            "---\n"
            "# 范围\n\n"
            "正文。\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n",
            kind="national",
        )
        root = ET.fromstring(xml)

        toc = self._et_paragraph_exact(root, "目次")
        preface = self._et_paragraph_exact(root, "前言")
        references = self._et_paragraph_exact(root, "参考文献")
        self.assertEqual(
            toc.find("w:pPr/w:spacing", self._W_NS).get(self._w_tag("after")),
            "468",
        )
        preface_spacing = preface.find("w:pPr/w:spacing", self._W_NS)
        self.assertEqual(preface_spacing.get(self._w_tag("before")), "900")
        self.assertEqual(preface_spacing.get(self._w_tag("after")), "468")
        self.assertEqual(
            references.find("w:pPr/w:spacing", self._W_NS).get(self._w_tag("after")),
            "156",
        )

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

    def test_cover_backend_page_numbering_sections_use_template_headers(self):
        parts = _build_cover_docx_parts(
            "---\n"
            "title: 分节页码测试标准\n"
            "---\n"
            "# 范围\n\n"
            "正文。\n\n"
            "# 附录 规范性 第一附录\n\n"
            "附录正文。\n\n"
            "# 附录 资料性 第二附录\n\n"
            "第二附录正文。\n\n"
            "# 参考文献\n\n"
            "GB/T 1.1—2020　标准化工作导则\n\n"
            "# 索引\n\n"
            "## B\n\n"
            "- 标准：1\n"
        )
        sections = self._sections(parts["document"])

        self.assertGreaterEqual(len(sections), 8)
        self.assertEqual(self._pg_num(sections[1]), ("upperRoman", "1"))
        self.assertEqual(self._pg_num(sections[2]), ("upperRoman", ""))
        self.assertEqual(self._pg_num(sections[3]), ("", "1"))
        for section in sections[1:]:
            refs = self._section_refs(section)
            self.assertIn(("headerReference", "default"), refs)
            self.assertIn(("footerReference", "default"), refs)

    def test_cover_backend_odd_even_pages_setting_is_metadata_controlled(self):
        default_parts = _build_cover_docx_parts("# 范围\n\n正文。\n")
        enabled_parts = _build_cover_docx_parts(
            "---\n"
            "odd_even_pages: true\n"
            "---\n\n"
            "# 范围\n\n正文。\n"
        )
        default_settings = ET.fromstring(default_parts["settings"])
        enabled_settings = ET.fromstring(enabled_parts["settings"])
        enabled_sections = self._sections(enabled_parts["document"])

        self.assertIsNone(default_settings.find("w:evenAndOddHeaders", self._W_NS))
        self.assertIsNotNone(enabled_settings.find("w:evenAndOddHeaders", self._W_NS))
        self.assertEqual(self._pg_num(enabled_sections[0]), ("", "0"))
        self.assertEqual(self._pg_num(enabled_sections[1]), ("upperRoman", "1"))
        refs = self._section_refs(enabled_sections[1])
        self.assertIn(("headerReference", "even"), refs)
        self.assertIn(("footerReference", "even"), refs)

    def test_cover_backend_odd_even_headers_and_footers_use_standard_styles(self):
        for kind, expected in (
            ("group", {
                "odd_header": "标准文件_页眉奇数页",
                "even_header": "标准文件_页眉偶数页",
                "odd_footer": "标准文件_页脚奇数页",
                "even_footer": "标准文件_页脚偶数页",
            }),
            ("national", {
                "odd_header": "标准文件_页眉奇数页",
                "even_header": "标准文件_页眉偶数页",
                "odd_footer": "标准文件_页脚奇数页",
                "even_footer": "标准文件_页脚偶数页",
            }),
        ):
            with self.subTest(kind=kind):
                parts = _build_cover_docx_parts(
                    "---\n"
                    "odd_even_pages: true\n"
                    "---\n\n"
                    "# 范围\n\n正文。\n",
                    kind=kind,
                )
                sections = self._sections(parts["document"])
                targets = self._section_ref_targets(parts["rels"], sections[1])

                odd_header = "word/" + targets[("headerReference", "default")]
                even_header = "word/" + targets[("headerReference", "even")]
                odd_footer = "word/" + targets[("footerReference", "default")]
                even_footer = "word/" + targets[("footerReference", "even")]

                self.assertEqual(
                    self._first_paragraph_style_name(parts["styles"], parts["headers"][odd_header]),
                    expected["odd_header"],
                )
                self.assertEqual(
                    self._first_paragraph_style_name(parts["styles"], parts["headers"][even_header]),
                    expected["even_header"],
                )
                self.assertEqual(
                    self._first_paragraph_style_name(parts["styles"], parts["footers"][odd_footer]),
                    expected["odd_footer"],
                )
                self.assertEqual(
                    self._first_paragraph_style_name(parts["styles"], parts["footers"][even_footer]),
                    expected["even_footer"],
                )
                self.assertEqual(self._first_paragraph_direct_jc(parts["headers"][odd_header]), "")
                self.assertEqual(self._first_paragraph_direct_jc(parts["headers"][even_header]), "")
                self.assertEqual(self._first_paragraph_direct_jc(parts["footers"][odd_footer]), "")
                self.assertEqual(self._first_paragraph_direct_jc(parts["footers"][even_footer]), "")

    def test_cover_backend_tunes_dash_and_reference_numbering_indents(self):
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            "勘测工作应重点查明以下内容：\n\n"
            "- 热储层的埋深、厚度、岩性组合、孔隙度和渗透率；\n\n"
            "# 参考文献\n\n"
            "GB/T 11615　地热资源地质勘查规范\n"
        )

        self.assertEqual(self._numbering_ind_by_style(parts["numbering"], "92"), {
            "left": "600",
            "hanging": "200",
        })
        self.assertEqual(self._numbering_ind_by_style(parts["numbering"], "64"), {
            "left": "620",
            "hanging": "420",
        })
        self.assertIsNone(self._style_ind_by_id(parts["styles"], "92"))
        self.assertIsNone(self._style_ind_by_id(parts["styles"], "93"))

    def test_cover_backend_makes_body_spacing_match_original_templates(self):
        group_parts = _build_cover_docx_parts("# 范围\n\n正文。\n", kind="group")
        national_parts = _build_cover_docx_parts("# 范围\n\n正文。\n", kind="national")

        group_body = self._style_spacing_by_name(group_parts["styles"], "标准文件_段")
        group_chapter = self._style_spacing_by_name(group_parts["styles"], "标准文件_章标题")
        group_clause = self._style_spacing_by_name(group_parts["styles"], "标准文件_一级条标题")
        group_table_caption = self._style_spacing_by_name(group_parts["styles"], "标准文件_正文表标题")
        group_note = self._style_spacing_by_name(group_parts["styles"], "标准文件_注：")
        national_body = self._style_spacing_by_name(national_parts["styles"], "标准文件_段")
        national_chapter = self._style_spacing_by_name(national_parts["styles"], "标准文件_章标题")

        self.assertEqual(group_body, {
            "before": "0",
            "beforeLines": None,
            "after": "0",
            "afterLines": None,
            "line": "400",
            "lineRule": "exact",
        })
        self.assertEqual(group_chapter, {
            "before": "240",
            "beforeLines": None,
            "after": "240",
            "afterLines": None,
            "line": "400",
            "lineRule": "exact",
        })
        self.assertEqual(group_clause, {
            "before": "50",
            "beforeLines": "50",
            "after": "50",
            "afterLines": "50",
            "line": "400",
            "lineRule": "exact",
        })
        self.assertEqual(group_table_caption, group_clause)
        self.assertEqual(group_note, {
            "before": "0",
            "beforeLines": None,
            "after": "0",
            "afterLines": None,
            "line": "300",
            "lineRule": "exact",
        })
        self.assertEqual(national_body, group_body)
        self.assertEqual(national_chapter, {
            "before": "312",
            "beforeLines": None,
            "after": "312",
            "afterLines": None,
            "line": "400",
            "lineRule": "exact",
        })

    def test_cover_backend_carries_template_direct_paragraph_format_and_run_hint(self):
        parts = _build_cover_docx_parts(
            "# 范围\n\n"
            "正文English。\n\n"
            "## 条标题\n\n"
            "条正文。\n"
        )
        root = ET.fromstring(parts["document"])
        body_para = self._et_paragraph_containing(root, "正文English。")
        clause_para = self._et_paragraph_containing(root, "条标题")

        body_ind = body_para.find("w:pPr/w:ind", self._W_NS)
        self.assertIsNotNone(body_ind)
        self.assertEqual(body_ind.get(self._w_tag("firstLine")), "420")
        self.assertIsNone(body_ind.get(self._w_tag("firstLineChars")))

        body_run_fonts = body_para.find(".//w:rPr/w:rFonts", self._W_NS)
        self.assertIsNotNone(body_run_fonts)
        self.assertEqual(body_run_fonts.get(self._w_tag("hint")), "eastAsia")

        clause_spacing = clause_para.find("w:pPr/w:spacing", self._W_NS)
        self.assertIsNotNone(clause_spacing)
        self.assertEqual(clause_spacing.get(self._w_tag("before")), "156")
        self.assertEqual(clause_spacing.get(self._w_tag("after")), "156")

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

    def test_cover_backend_replaces_consistency_degree_placeholder(self):
        xml = _build_cover_docx_xml(
            "---\n"
            "standard_type: 国家标准\n"
            "number: GB/T 99999-2026\n"
            "title: 一致性标识测试标准\n"
            "title_en: Consistency degree test standard\n"
            "consistency_degree: MOD\n"
            "---\n\n"
            "# 范围\n\n"
            "正文。\n",
            kind="national",
        )

        self.assertIn("（MOD）", xml)
        self.assertNotIn("点击此处添加与国际标准一致性程度的标识", xml)

    def test_cover_backend_keeps_seq_groups_and_appendix_scope(self):
        parts = _build_cover_docx_parts(
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
        xml = parts["document"].decode("utf-8", errors="ignore")
        root = ET.fromstring(parts["document"])
        body_table = self._et_paragraph_containing(root, "主表")
        appendix_table = self._et_paragraph_containing(root, "附录B表")
        table_label_style = self._style_id_by_name(parts["styles"], "标准文件_附录表标号")
        figure_label_style = self._style_id_by_name(parts["styles"], "标准文件_附录图标号")

        self.assertNotIn("SEQ 表", xml)
        self.assertNotIn("SEQ 图", xml)
        self.assertIn(" SEQ 公式 \\* ARABIC \\r 1 ", xml)
        self.assertIn("\\h \\r", xml)
        self.assertEqual(
            self._paragraph_style(body_table),
            self._style_id_by_name(parts["styles"], "标准文件_正文表标题"),
        )
        self.assertEqual(
            self._paragraph_style(appendix_table),
            self._style_id_by_name(parts["styles"], "标准文件_附录表标题"),
        )
        self.assertNotEqual(self._num_id_value(body_table), "0")
        self.assertNotEqual(self._num_id_value(appendix_table), "0")
        self.assertEqual(len(self._et_paragraphs_with_style(root, table_label_style)), 2)
        self.assertEqual(len(self._et_paragraphs_with_style(root, figure_label_style)), 2)
        self.assertNotIn("<w:t>A.</w:t>", xml)
        self.assertNotIn("<w:t>B.</w:t>", xml)
        self.assertNotIn("SEQ TableA", xml)
        self.assertNotIn("SEQ TableB", xml)
        self.assertNotIn("SEQ Equation", xml)

    def test_cover_backend_uses_cover_blueprint_parts(self):
        parts = _build_cover_docx_parts("# 范围\n\n正文。\n", kind="group")
        with zipfile.ZipFile(docx_builder._default_cover_path("group")) as zf:
            self.assertEqual(self._style_names(parts["styles"]), self._style_names(zf.read("word/styles.xml")))
            self.assertNotEqual(parts["document"], zf.read("word/document.xml"))

    def test_cover_backend_preserves_legacy_form_dropdowns_without_default_protection(self):
        for kind in ("group", "national"):
            with self.subTest(kind=kind):
                parts = _build_cover_docx_parts("# 范围\n\n正文。\n", kind=kind)
                dropdown_fields = self._legacy_dropdown_fields(parts["document"])
                protection = self._document_protection(parts["settings"])
                sections = self._sections(parts["document"])

                self.assertIn(b"FORMDROPDOWN", parts["document"])
                self.assertIn(b"<w:ddList", parts["document"])
                self.assertEqual(protection.get("edit"), "forms")
                self.assertEqual(protection.get("enforcement"), "0")
                for section in sections:
                    self.assertEqual(self._section_form_prot(section), "0")
                self.assertEqual(len(dropdown_fields), 2)
                self.assertEqual(dropdown_fields[0]["name"], "下拉1")
                self.assertEqual(dropdown_fields[0]["result"], "1")
                self.assertEqual(dropdown_fields[0]["selected"], "草案版次选择")
                self.assertEqual(dropdown_fields[0]["entries"][0], " ")
                self.assertIn("草案版次选择", dropdown_fields[0]["entries"])
                self.assertEqual(dropdown_fields[1]["name"], "下拉2")
                self.assertEqual(dropdown_fields[1]["result"], "1")
                self.assertEqual(dropdown_fields[1]["entries"][0], " ")
                self.assertIn("在提交反馈意见时，请将您知道的相关专利连同支持性文件一并附上。", dropdown_fields[1]["entries"])

    def test_cover_backend_metadata_enables_cover_form_protection(self):
        parts = _build_cover_docx_parts(
            "---\n"
            "cover_form_protection: true\n"
            "---\n\n"
            "# 范围\n\n正文。\n"
        )
        protection = self._document_protection(parts["settings"])
        sections = self._sections(parts["document"])

        self.assertEqual(protection.get("edit"), "forms")
        self.assertEqual(protection.get("enforcement"), "1")
        self.assertNotIn("hash", protection)
        self.assertNotIn("salt", protection)
        self.assertEqual(self._section_form_prot(sections[0]), "")
        for section in sections[1:]:
            self.assertEqual(self._section_form_prot(section), "0")

    def test_cover_backend_metadata_sets_draft_version_dropdown(self):
        parts = _build_cover_docx_parts(
            "---\n"
            "draft_version: 征求意见稿\n"
            "---\n\n"
            "# 范围\n\n正文。\n"
        )
        dropdown_fields = self._legacy_dropdown_fields(parts["document"])

        self.assertEqual(dropdown_fields[0]["name"], "下拉1")
        self.assertEqual(dropdown_fields[0]["result"], "3")
        self.assertEqual(dropdown_fields[0]["selected"], "（征求意见稿）")

    def test_cover_backend_end_line_uses_packaged_cover_image(self):
        cases = [
            ("group", docx_builder._default_cover_path("group"), "word/media/image1.jpeg"),
            ("national", docx_builder._default_cover_path("national"), "word/media/image3.jpg"),
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
            "<table><tr><td>最高设计车速（$$v_{\\text{max}}$$）</td><td>测试车速</td></tr>"
            "<tr><td>$$v_{\\text{max}}$$≤45</td><td>80%$$v_{\\text{max}}$$</td></tr>"
            "<tr><td colspan=\"2\">注：按临近分度线取值。</td></tr></table>"
        )

        self.assertIn("<w:gridSpan w:val=\"2\"/>", xml)
        self.assertIn("<m:oMath", xml)
        self.assertIn("注：按临近分度线取值。", xml)
        self.assertIn("<w:vAlign w:val=\"center\"/>", xml)
        self.assertIn("<w:jc w:val=\"center\"/>", xml)
        self.assertIn("<w:jc w:val=\"left\"/>", xml)

    def test_docx_table_cell_alignment_defaults_and_overrides(self):
        parts = _build_docx_parts(
            "# 范围\n\n"
            "{表：#tbl:align} 对齐表\n\n"
            "<table>"
            "<tr><th>表头</th><th>数值</th><th>说明</th><th>显式左</th><th>显式右</th><th>小数列</th></tr>"
            "<tr><td>A</td><td>12.5</td><td>应符合运行、维护和记录要求。</td>"
            "<td data-align=\"left\">短值</td><td data-align=\"right\">98.6</td>"
            "<td data-align=\"decimal\">3.14</td></tr>"
            "</table>"
        )
        root = ET.fromstring(parts["document"])

        def cell_jc(text: str) -> str:
            cell = self._et_cell_containing(root, text)
            self.assertIsNotNone(cell)
            para = cell.find("w:p", self._W_NS)
            self.assertIsNotNone(para)
            return self._jc_value(para)

        self.assertEqual(cell_jc("表头"), "center")
        self.assertEqual(cell_jc("12.5"), "center")
        self.assertEqual(cell_jc("应符合运行、维护和记录要求。"), "left")
        self.assertEqual(cell_jc("短值"), "left")
        self.assertEqual(cell_jc("98.6"), "right")
        self.assertEqual(cell_jc("3.14"), "right")

    def test_docx_table_header_cells_default_bottom_border_is_thick(self):
        parts = _build_docx_parts(
            "# 范围\n\n"
            "{表：#tbl:gfm} 管道表\n\n"
            "| 管道表头 | 值 |\n"
            "| --- | --- |\n"
            "| A | 1 |\n\n"
            "{表：#tbl:html} HTML表\n\n"
            "<table>"
            "<tr><th>默认粗底线</th><th data-border-bottom=\"none\">显式无底线</th></tr>"
            "<tr><td>A</td><td>1</td></tr>"
            "</table>\n"
        )
        root = ET.fromstring(parts["document"])
        gfm_header_cell = self._et_cell_containing(root, "管道表头")
        html_header_cell = self._et_cell_containing(root, "默认粗底线")
        override_cell = self._et_cell_containing(root, "显式无底线")

        self.assertEqual(self._cell_border_attrs(gfm_header_cell, "bottom"), ("single", "8"))
        self.assertEqual(self._cell_border_attrs(html_header_cell, "bottom"), ("single", "8"))
        self.assertEqual(self._cell_border_attrs(override_cell, "bottom"), ("nil", "0"))

    def test_docx_html_table_rowspan_borders_empty_and_same(self):
        xml = _build_docx_xml(
            "# 范围\n\n"
            "示例：\n\n"
            "第一段示例内容。\n\n"
            "{表：#tbl:merge} 合并边框表\n\n"
            "<table data-border-outer=\"thick\" data-border-inner=\"thin\">"
            "<tr><th rowspan=\"2\" data-border-right=\"thick\">类别</th><th colspan=\"2\">指标</th></tr>"
            "<tr><th>值</th><th data-border-bottom=\"none\">备注</th></tr>"
            "<tr><td>一类</td><td></td><td>同上</td></tr>"
            "</table>\n\n"
            "第二段示例内容。\n\n"
            "{示例结束}\n"
        )

        self.assertIn('<w:vMerge w:val="restart"', xml)
        self.assertIn("<w:vMerge/>", xml)
        self.assertIn('<w:gridSpan w:val="2"', xml)
        self.assertIn('<w:right w:val="single" w:sz="8"', xml)
        self.assertIn('<w:bottom w:val="nil" w:sz="0"', xml)
        self.assertIn(">同上<", xml)
        self.assertIn("第一段示例内容", xml)
        self.assertIn("第二段示例内容", xml)
        self.assertGreaterEqual(xml.count('w:pStyle w:val="182"'), 2)

    def test_docx_moderate_table_is_not_split_as_continued_table(self):
        rows = "".join("| %d | 值%d |\n" % (i, i) for i in range(1, 8))
        parts = _build_docx_parts(
            "# 范围\n\n"
            "{表：#tbl:moderate} 中等表\n\n"
            "| 项 | 值 |\n"
            "| --- | --- |\n" +
            rows
        )
        xml = parts["document"].decode("utf-8", errors="ignore")
        root = ET.fromstring(parts["document"])
        caption = self._et_paragraph_containing(root, "中等表")

        self.assertNotIn("SEQ 表", xml)
        self.assertEqual(
            self._paragraph_style(caption),
            self._style_id_by_name(parts["styles"], "标准文件_正文表标题"),
        )
        self.assertNotEqual(self._num_id_value(caption), "0")
        self.assertNotIn("（续）", xml)
        self.assertEqual(xml.count("<w:tblHeader w:val=\"true\"/>"), 1)

    def test_docx_long_table_is_not_split_before_render_pagination(self):
        rows = "".join("| %d | 值%d |\n" % (i, i) for i in range(1, 14))
        parts = _build_docx_parts(
            "# 范围\n\n"
            "{表：#tbl:long} 长表\n\n"
            "| 项 | 值 |\n"
            "| --- | --- |\n" +
            rows
        )
        xml = parts["document"].decode("utf-8", errors="ignore")
        root = ET.fromstring(parts["document"])
        caption = self._et_paragraph_containing(root, "长表")

        self.assertNotIn("SEQ 表", xml)
        self.assertEqual(
            self._paragraph_style(caption),
            self._style_id_by_name(parts["styles"], "标准文件_正文表标题"),
        )
        self.assertIsNotNone(caption.find("w:pPr/w:keepNext", self._W_NS))
        self.assertNotEqual(self._num_id_value(caption), "0")
        self.assertNotIn("（续）", xml)
        self.assertEqual(xml.count("<w:tblHeader w:val=\"true\"/>"), 1)

    def test_word_postprocess_splits_continuation_table_from_measured_plan(self):
        sdoc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:split} 测试表\n\n"
            "| 唯一表头 | 值 |\n"
            "| --- | --- |\n"
            "| 一 | 1 |\n"
            "| 二 | 2 |\n"
            "| 三 | 3 |\n"
            "| 四 | 4 |\n"
        )
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            docx_builder.build_cover(sdoc, path)
            changed = word_postprocess._apply_table_continuations(path, [
                word_postprocess._TableContinuationPlan(
                    table_index=2,
                    row_breaks=[4],
                    header_count=1,
                    caption_text="表1　测试表（续）",
                )
            ])
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
        finally:
            if os.path.exists(path):
                os.remove(path)

        self.assertTrue(changed)
        self.assertIn("表1　测试表（续）", xml)
        self.assertIn('w:pStyle w:val="185"', xml)
        self.assertNotIn("SEQ 表", xml)
        self.assertEqual(xml.count("唯一表头"), 2)
        root = ET.fromstring(xml)
        self.assertEqual(len(root.findall(".//w:tblHeader", self._W_NS)), 2)
        continuation = self._et_paragraph_containing(root, "表1　测试表（续）")
        self.assertIsNotNone(continuation)
        self.assertIsNone(continuation.find("w:pPr/w:pageBreakBefore", self._W_NS))
        self.assertIsNotNone(continuation.find("w:pPr/w:keepNext", self._W_NS))
        spacing = continuation.find("w:pPr/w:spacing", self._W_NS)
        self.assertIsNotNone(spacing)
        self.assertEqual(spacing.get(self._w_tag("before")), "50")
        self.assertEqual(spacing.get(self._w_tag("after")), "50")
        self.assertEqual(spacing.get(self._w_tag("line")), "400")
        self.assertEqual(spacing.get(self._w_tag("lineRule")), "exact")

    def test_word_postprocess_applies_only_first_measured_break_per_iteration(self):
        sdoc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:split} 测试表\n\n"
            "| 唯一表头 | 值 |\n"
            "| --- | --- |\n"
            "| 一 | 1 |\n"
            "| 二 | 2 |\n"
            "| 三 | 3 |\n"
            "| 四 | 4 |\n"
            "| 五 | 5 |\n"
            "| 六 | 6 |\n"
        )
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            docx_builder.build_cover(sdoc, path)
            changed = word_postprocess._apply_table_continuations(path, [
                word_postprocess._TableContinuationPlan(
                    table_index=2,
                    row_breaks=[4, 6],
                    header_count=1,
                    caption_text="表1　测试表（续）",
                )
            ])
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
        finally:
            if os.path.exists(path):
                os.remove(path)

        self.assertTrue(changed)
        self.assertEqual(xml.count("表1　测试表（续）"), 1)
        self.assertEqual(xml.count("唯一表头"), 2)
        root = ET.fromstring(xml)
        split_tables = [
            table for table in root.findall(".//w:tbl", self._W_NS)
            if "唯一表头" in self._et_text(table)
        ]
        self.assertEqual(len(split_tables), 2)
        self.assertIn("二", self._et_text(split_tables[0]))
        self.assertNotIn("三", self._et_text(split_tables[0]))
        self.assertIn("三", self._et_text(split_tables[1]))
        self.assertIn("六", self._et_text(split_tables[1]))

    def test_word_postprocess_detects_row_visually_pushed_to_next_page(self):
        class FakeRangeValues:
            def __init__(self, start, end):
                self.Start = start
                self.End = end

        class FakeRow:
            def __init__(self, start, end):
                self.Range = FakeRangeValues(start, end)

        class FakeRows:
            def __init__(self, rows):
                self._rows = rows

            def __call__(self, index):
                return self._rows[index - 1]

        class FakeTable:
            def __init__(self, rows):
                self.Rows = FakeRows(rows)

        class FakeRange:
            def __init__(self, page):
                self._page = page

            def Information(self, _key):
                return self._page

        class FakeDoc:
            def __init__(self, pages):
                self._pages = pages

            def Range(self, start, _end):
                return FakeRange(self._pages[start])

        rows = [
            FakeRow(10, 11),
            FakeRow(20, 21),
            FakeRow(30, 31),
            FakeRow(40, 41),
            FakeRow(50, 51),
            FakeRow(60, 61),
            FakeRow(70, 71),
        ]
        pages = {
            10: 10, 11: 10,
            20: 10, 21: 10,
            30: 10, 31: 10,
            40: 10, 41: 10,
            50: 10, 51: 10,
            60: 10, 61: 11,
            70: 11, 71: 11,
        }

        breaks = word_postprocess._table_row_page_breaks(
            FakeDoc(pages),
            FakeTable(rows),
            row_count=len(rows),
            header_count=1,
        )

        self.assertEqual(breaks, [6])

    def test_word_postprocess_detects_vertical_merge_table_breaks_from_cells(self):
        class FakeRangeValues:
            def __init__(self, start, end):
                self.Start = start
                self.End = end

        class FakeRows:
            def __call__(self, _index):
                raise RuntimeError("vertical merge")

        class FakeCell:
            def __init__(self, row_index, start, end):
                self.RowIndex = row_index
                self.Range = FakeRangeValues(start, end)

        class FakeCells:
            def __init__(self, cells):
                self._cells = cells
                self.Count = len(cells)

            def Item(self, index):
                return self._cells[index - 1]

        class FakeRange:
            def __init__(self, page=None, cells=None):
                self._page = page
                self.Cells = cells

            def Information(self, _key):
                return self._page

        class FakeTable:
            def __init__(self, cells):
                self.Rows = FakeRows()
                self.Range = FakeRange(cells=FakeCells(cells))

        class FakeDoc:
            def __init__(self, pages):
                self._pages = pages

            def Range(self, start, _end):
                return FakeRange(page=self._pages[start])

        cells = [
            FakeCell(1, 10, 11),
            FakeCell(1, 12, 13),
            FakeCell(2, 20, 21),
            FakeCell(2, 22, 23),
            FakeCell(3, 30, 31),
            FakeCell(3, 32, 33),
        ]
        pages = {
            10: 1, 11: 1, 12: 1, 13: 1,
            20: 1, 21: 1, 22: 1, 23: 1,
            30: 2, 31: 2, 32: 2, 33: 2,
        }

        breaks = word_postprocess._table_row_page_breaks(
            FakeDoc(pages),
            FakeTable(cells),
            row_count=3,
            header_count=1,
        )

        self.assertEqual(breaks, [3])

    def test_word_postprocess_moves_break_after_vmerge_continuation(self):
        rows = [
            ET.fromstring("<w:tr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tc/></w:tr>"),
            ET.fromstring("<w:tr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr></w:tc></w:tr>"),
            ET.fromstring("<w:tr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tc><w:tcPr><w:vMerge/></w:tcPr></w:tc></w:tr>"),
            ET.fromstring("<w:tr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tc/></w:tr>"),
        ]

        self.assertEqual(word_postprocess._safe_vmerge_breaks(rows, [3]), [4])

    def test_word_postprocess_applies_one_plan_before_remeasuring(self):
        first = word_postprocess._TableContinuationPlan(
            table_index=1,
            row_breaks=[4],
            header_count=1,
            caption_text="表1　前表（续）",
        )
        stale_second = word_postprocess._TableContinuationPlan(
            table_index=2,
            row_breaks=[6],
            header_count=1,
            caption_text="表2　后表（续）",
        )
        fresh_second = word_postprocess._TableContinuationPlan(
            table_index=3,
            row_breaks=[5],
            header_count=1,
            caption_text="表2　后表（续）",
        )
        measured = [[first, stale_second], [fresh_second], []]
        applied = []

        def fake_measure(_word, _path):
            return measured.pop(0)

        def fake_apply(_path, plans):
            applied.append(plans)
            return True

        def fake_measure_layout(word, path):
            return fake_measure(word, path), []

        with mock.patch("md2std.word_postprocess._measure_table_layout_plans", fake_measure_layout), \
             mock.patch("md2std.word_postprocess._apply_table_continuations", fake_apply):
            word_postprocess._postprocess_document(object(), "dummy.docx")

        self.assertEqual(applied, [[first], [fresh_second]])

    def test_word_postprocess_dedupes_identical_continuation_plans(self):
        first = word_postprocess._TableContinuationPlan(
            table_index=1,
            row_breaks=[4],
            header_count=1,
            caption_text="表1　前表（续）",
        )
        duplicate = word_postprocess._TableContinuationPlan(
            table_index=1,
            row_breaks=[4],
            header_count=1,
            caption_text="表1　前表（续）",
        )
        second = word_postprocess._TableContinuationPlan(
            table_index=2,
            row_breaks=[6],
            header_count=1,
            caption_text="表2　后表（续）",
        )

        plans = word_postprocess._dedupe_table_continuation_plans([first, duplicate, second])

        self.assertEqual(plans, [first, second])

    def test_word_postprocess_detects_self_spanning_rows(self):
        spans = [
            (1, 3, 3),
            (2, 3, 4),
            (3, 4, 4),
            (4, 4, 6),
        ]

        self.assertEqual(word_postprocess._self_spanning_row_indices(spans), [2, 4])

    def test_word_postprocess_stops_on_repeated_continuation_plan(self):
        plan = word_postprocess._TableContinuationPlan(
            table_index=1,
            row_breaks=[4],
            header_count=1,
            caption_text="表1　测试表（续）",
        )
        measured = [[plan], [plan]]
        applied = []

        def fake_measure(_word, _path):
            return measured.pop(0)

        def fake_apply(_path, plans):
            applied.append(plans)
            return True

        def fake_measure_layout(word, path):
            return fake_measure(word, path), []

        with mock.patch("md2std.word_postprocess._measure_table_layout_plans", fake_measure_layout), \
             mock.patch("md2std.word_postprocess._apply_table_continuations", fake_apply):
            word_postprocess._postprocess_document(object(), "dummy.docx")

        self.assertEqual(applied, [[plan]])

    def test_word_postprocess_applies_horizontal_split_when_no_vertical_plan(self):
        horizontal = word_postprocess._HorizontalTableSplitPlan(
            table_index=1,
            caption_text="表1　宽表（续）",
        )
        measured = [([], [horizontal]), ([], [])]
        applied = []

        def fake_measure(_word, _path):
            return measured.pop(0)

        def fake_horizontal(_path, plans):
            applied.append(plans)
            return len(applied) == 1

        with mock.patch("md2std.word_postprocess._measure_table_layout_plans", fake_measure), \
             mock.patch("md2std.word_postprocess._apply_horizontal_table_splits", fake_horizontal):
            word_postprocess._postprocess_document(object(), "dummy.docx")

        self.assertEqual(applied, [[horizontal]])

    def test_word_postprocess_splits_wide_table_horizontally(self):
        sdoc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:wide} 宽表\n\n"
            "| A | B | C | D | E | F | G | H | I | J |\n"
            "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |\n"
            "| a | b | c | d | e | f | g | h | i | j |\n"
        )
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            docx_builder.build_cover(sdoc, path)
            changed = word_postprocess._apply_horizontal_table_splits(path, [
                word_postprocess._HorizontalTableSplitPlan(
                    table_index=2,
                    caption_text="表1　宽表（续）",
                )
            ])
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
        finally:
            if os.path.exists(path):
                os.remove(path)

        self.assertTrue(changed)
        self.assertIn("表1　宽表（续）", xml)
        root = ET.fromstring(xml)
        split_tables = [
            table for table in root.findall(".//w:tbl", self._W_NS)
            if "A" in self._et_text(table) and ("I" in self._et_text(table) or "J" in self._et_text(table))
        ]
        self.assertEqual(len(split_tables), 2)
        self.assertIn("I", self._et_text(split_tables[0]))
        self.assertNotIn("J", self._et_text(split_tables[0]))
        self.assertIn("A", self._et_text(split_tables[1]))
        self.assertIn("J", self._et_text(split_tables[1]))
        self.assertNotIn("B", self._et_text(split_tables[1]))

    def test_word_postprocess_splits_when_first_segment_has_one_body_row(self):
        sdoc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:split} 测试表\n\n"
            "| 唯一表头 | 值 |\n"
            "| --- | --- |\n"
            "| 一 | 1 |\n"
            "| 二 | 2 |\n"
            "| 三 | 3 |\n"
        )
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            docx_builder.build_cover(sdoc, path)
            changed = word_postprocess._apply_table_continuations(path, [
                word_postprocess._TableContinuationPlan(
                    table_index=2,
                    row_breaks=[3],
                    header_count=1,
                    caption_text="表1　测试表（续）",
                )
            ])
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
        finally:
            if os.path.exists(path):
                os.remove(path)

        self.assertTrue(changed)
        self.assertIn("表1　测试表（续）", xml)
        self.assertEqual(xml.count("唯一表头"), 2)

    def test_word_postprocess_splits_when_tail_segment_has_one_body_row(self):
        sdoc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:split} 测试表\n\n"
            "| 唯一表头 | 值 |\n"
            "| --- | --- |\n"
            "| 一 | 1 |\n"
            "| 二 | 2 |\n"
            "| 三 | 3 |\n"
        )
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            docx_builder.build_cover(sdoc, path)
            changed = word_postprocess._apply_table_continuations(path, [
                word_postprocess._TableContinuationPlan(
                    table_index=2,
                    row_breaks=[4],
                    header_count=1,
                    caption_text="表1　测试表（续）",
                )
            ])
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
        finally:
            if os.path.exists(path):
                os.remove(path)

        self.assertTrue(changed)
        self.assertIn("表1　测试表（续）", xml)
        self.assertEqual(xml.count("唯一表头"), 2)

    def test_word_postprocess_keeps_measured_break_for_short_tail(self):
        sdoc = md_parser.parse(
            "# 范围\n\n"
            "{表：#tbl:split} 测试表\n\n"
            "| 唯一表头 | 值 |\n"
            "| --- | --- |\n"
            "| 一 | 1 |\n"
            "| 二 | 2 |\n"
            "| 三 | 3 |\n"
            "| 四 | 4 |\n"
            "| 五 | 5 |\n"
        )
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            docx_builder.build_cover(sdoc, path)
            changed = word_postprocess._apply_table_continuations(path, [
                word_postprocess._TableContinuationPlan(
                    table_index=2,
                    row_breaks=[6],
                    header_count=1,
                    caption_text="表1　测试表（续）",
                )
            ])
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
        finally:
            if os.path.exists(path):
                os.remove(path)

        self.assertTrue(changed)
        self.assertIn("表1　测试表（续）", xml)
        self.assertEqual(xml.count("唯一表头"), 2)
        root = ET.fromstring(xml)
        continuation = self._et_paragraph_containing(root, "表1　测试表（续）")
        self.assertIsNone(continuation.find("w:pPr/w:pageBreakBefore", self._W_NS))
        self.assertIsNotNone(continuation.find("w:pPr/w:keepNext", self._W_NS))
        spacing = continuation.find("w:pPr/w:spacing", self._W_NS)
        self.assertIsNotNone(spacing)
        self.assertEqual(spacing.get(self._w_tag("before")), "50")
        self.assertEqual(spacing.get(self._w_tag("after")), "50")
        self.assertEqual(spacing.get(self._w_tag("line")), "400")
        self.assertEqual(spacing.get(self._w_tag("lineRule")), "exact")
        split_tables = [
            table for table in root.findall(".//w:tbl", self._W_NS)
            if "唯一表头" in self._et_text(table)
        ]
        self.assertEqual(len(split_tables), 2)
        self.assertIn("四", self._et_text(split_tables[0]))
        self.assertNotIn("五", self._et_text(split_tables[0]))
        self.assertIn("五", self._et_text(split_tables[1]))

    def test_cover_blueprints_keep_complete_cover_section(self):
        pairs = [
            (self._first_existing_path(
                *resources.template_candidates("template_group.docx"),
            ), docx_builder._default_cover_path("group")),
            (self._first_existing_path(
                *resources.template_candidates("template_national.docx"),
            ), docx_builder._default_cover_path("national")),
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

    def _et_paragraph_containing(self, root, text: str):
        for paragraph in root.findall(".//w:p", self._W_NS):
            if text in self._et_text(paragraph):
                return paragraph
        return None

    def _et_paragraphs_with_style(self, root, style_id: str):
        paragraphs = []
        for paragraph in root.findall(".//w:p", self._W_NS):
            pstyle = paragraph.find("w:pPr/w:pStyle", self._W_NS)
            if pstyle is not None and pstyle.get(self._w_tag("val")) == style_id:
                paragraphs.append(paragraph)
        return paragraphs

    def _et_paragraph_exact(self, root, text: str):
        for paragraph in root.findall(".//w:p", self._W_NS):
            if self._et_text(paragraph) == text:
                return paragraph
        return None

    def _et_cell_containing(self, root, text: str):
        for cell in root.findall(".//w:tc", self._W_NS):
            if text in self._et_text(cell):
                return cell
        return None

    def _cell_border_attrs(self, cell, edge: str) -> tuple:
        border = cell.find(f"w:tcPr/w:tcBorders/w:{edge}", self._W_NS)
        self.assertIsNotNone(border)
        return (
            border.get(self._w_tag("val")),
            border.get(self._w_tag("sz")),
        )

    def _jc_value(self, paragraph) -> str:
        jc = paragraph.find("w:pPr/w:jc", self._W_NS)
        return jc.get(self._w_tag("val")) if jc is not None else ""

    def _num_id_value(self, paragraph) -> str:
        num_id = paragraph.find("w:pPr/w:numPr/w:numId", self._W_NS)
        self.assertIsNotNone(num_id)
        return num_id.get(self._w_tag("val"))

    def _paragraph_style(self, paragraph) -> str:
        pstyle = paragraph.find("w:pPr/w:pStyle", self._W_NS)
        self.assertIsNotNone(pstyle)
        return pstyle.get(self._w_tag("val"))

    def _style_id_by_name(self, styles_xml: bytes, name: str) -> str:
        root = ET.fromstring(styles_xml)
        for style in root.findall("w:style", self._W_NS):
            name_el = style.find("w:name", self._W_NS)
            if name_el is not None and name_el.get(self._w_tag("val")) == name:
                return style.get(self._w_tag("styleId"))
        self.fail("找不到样式：%s" % name)

    def _style_spacing_by_name(self, styles_xml: bytes, name: str) -> dict:
        root = ET.fromstring(styles_xml)
        for style in root.findall("w:style", self._W_NS):
            name_el = style.find("w:name", self._W_NS)
            if name_el is None or name_el.get(self._w_tag("val")) != name:
                continue
            spacing = style.find("w:pPr/w:spacing", self._W_NS)
            self.assertIsNotNone(spacing)
            return {
                "before": spacing.get(self._w_tag("before")),
                "beforeLines": spacing.get(self._w_tag("beforeLines")),
                "after": spacing.get(self._w_tag("after")),
                "afterLines": spacing.get(self._w_tag("afterLines")),
                "line": spacing.get(self._w_tag("line")),
                "lineRule": spacing.get(self._w_tag("lineRule")),
            }
        self.fail("找不到样式：%s" % name)

    def _w_tag(self, name: str) -> str:
        return "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + name

    def _run_spacing(self, run) -> str:
        spacing = run.find("w:rPr/w:spacing", self._W_NS)
        return spacing.get(self._w_tag("val")) if spacing is not None else None

    def _sections(self, document_xml: bytes) -> list:
        root = ET.fromstring(document_xml)
        body = root.find("w:body", self._W_NS)
        sections = []
        for paragraph in body.findall("w:p", self._W_NS):
            sect = paragraph.find("w:pPr/w:sectPr", self._W_NS)
            if sect is not None:
                sections.append(sect)
        final_sect = body.find("w:sectPr", self._W_NS)
        if final_sect is not None:
            sections.append(final_sect)
        return sections

    def _pg_num(self, section) -> tuple:
        pg = section.find("w:pgNumType", self._W_NS)
        if pg is None:
            return "", ""
        return pg.get(self._w_tag("fmt"), ""), pg.get(self._w_tag("start"), "")

    def _section_refs(self, section) -> set:
        refs = set()
        for child in section:
            local = child.tag.rsplit("}", 1)[-1]
            if local in ("headerReference", "footerReference"):
                refs.add((local, child.get(self._w_tag("type"))))
        return refs

    def _section_ref_targets(self, rels_xml: bytes, section) -> dict:
        rels_root = ET.fromstring(rels_xml)
        rels_ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
        rel_targets = {
            rel.get("Id"): rel.get("Target")
            for rel in rels_root.findall("r:Relationship", rels_ns)
        }
        targets = {}
        for child in section:
            local = child.tag.rsplit("}", 1)[-1]
            if local not in ("headerReference", "footerReference"):
                continue
            rel_id = child.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            targets[(local, child.get(self._w_tag("type")))] = rel_targets[rel_id]
        return targets

    def _first_paragraph_style_name(self, styles_xml: bytes, part_xml: bytes) -> str:
        styles_root = ET.fromstring(styles_xml)
        style_names = {}
        for style in styles_root.findall(".//w:style", self._W_NS):
            name = style.find("w:name", self._W_NS)
            if name is not None:
                style_names[style.get(self._w_tag("styleId"))] = name.get(self._w_tag("val"))

        part_root = ET.fromstring(part_xml)
        pstyle = part_root.find(".//w:p/w:pPr/w:pStyle", self._W_NS)
        self.assertIsNotNone(pstyle)
        return style_names[pstyle.get(self._w_tag("val"))]

    def _first_paragraph_direct_jc(self, part_xml: bytes) -> str:
        part_root = ET.fromstring(part_xml)
        jc = part_root.find(".//w:p/w:pPr/w:jc", self._W_NS)
        return jc.get(self._w_tag("val")) if jc is not None else ""

    def _legacy_dropdown_fields(self, document_xml: bytes) -> list:
        root = ET.fromstring(document_xml)
        fields = []
        for fld_char in root.findall(".//w:fldChar", self._W_NS):
            ffdata = fld_char.find("w:ffData", self._W_NS)
            if ffdata is None:
                continue
            ddlist = ffdata.find("w:ddList", self._W_NS)
            if ddlist is None:
                continue
            name = ffdata.find("w:name", self._W_NS)
            result = ddlist.find("w:result", self._W_NS)
            entries = [
                item.get(self._w_tag("val"))
                for item in ddlist.findall("w:listEntry", self._W_NS)
            ]
            raw_result = result.get(self._w_tag("val")) if result is not None else ""
            selected_index = int(raw_result) if raw_result.isdigit() else -1
            fields.append({
                "name": name.get(self._w_tag("val")) if name is not None else "",
                "result": raw_result,
                "selected": entries[selected_index] if 0 <= selected_index < len(entries) else "",
                "entries": entries,
            })
        return fields

    def _document_protection(self, settings_xml: bytes) -> dict:
        root = ET.fromstring(settings_xml)
        protection = root.find(".//w:documentProtection", self._W_NS)
        self.assertIsNotNone(protection)
        return {
            key.rsplit("}", 1)[-1]: value
            for key, value in protection.attrib.items()
        }

    def _section_form_prot(self, section) -> str:
        form_prot = section.find("w:formProt", self._W_NS)
        return form_prot.get(self._w_tag("val")) if form_prot is not None else ""

    def _numbering_ind_by_style(self, numbering_xml: bytes, style_id: str) -> dict:
        root = ET.fromstring(numbering_xml)
        for lvl in root.findall(".//w:lvl", self._W_NS):
            pstyle = lvl.find("w:pStyle", self._W_NS)
            if pstyle is None or pstyle.get(self._w_tag("val")) != style_id:
                continue
            ind = lvl.find("w:pPr/w:ind", self._W_NS)
            self.assertIsNotNone(ind)
            return {
                key.rsplit("}", 1)[-1]: value
                for key, value in ind.attrib.items()
            }
        raise AssertionError("找不到编号样式：%s" % style_id)

    def _style_ind_by_id(self, styles_xml: bytes, style_id: str):
        root = ET.fromstring(styles_xml)
        for style in root.findall(".//w:style", self._W_NS):
            if style.get(self._w_tag("styleId")) != style_id:
                continue
            return style.find("w:pPr/w:ind", self._W_NS)
        raise AssertionError("找不到样式：%s" % style_id)

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


class CliPostprocessTest(unittest.TestCase):
    def test_legacy_template_backend_flags_are_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            input_path = os.path.join(td, "input.md")
            with open(input_path, "w", encoding="utf-8") as f:
                f.write("# 范围\n\n正文。\n")

            with mock.patch("sys.stderr"), self.assertRaises(SystemExit) as cm:
                cli.main([input_path, "--backend", "template"])

        self.assertEqual(cm.exception.code, 2)

    def test_cover_form_protection_flag_is_forwarded_and_can_override_metadata(self):
        with tempfile.TemporaryDirectory() as td:
            input_path = os.path.join(td, "input.md")
            output_path = os.path.join(td, "output.docx")
            with open(input_path, "w", encoding="utf-8") as f:
                f.write(
                    "---\n"
                    "cover_form_protection: true\n"
                    "---\n\n"
                    "# 范围\n\n正文。\n"
                )

            with mock.patch("md2std.docx_builder.build_cover", return_value=output_path) as build_cover:
                self.assertEqual(cli.main([input_path, "-o", output_path]), 0)
                self.assertIsNone(build_cover.call_args.kwargs["cover_form_protection"])

                self.assertEqual(cli.main([input_path, "-o", output_path, "--cover-form-protection"]), 0)
                self.assertTrue(build_cover.call_args.kwargs["cover_form_protection"])

                self.assertEqual(cli.main([input_path, "-o", output_path, "--no-cover-form-protection"]), 0)
                self.assertFalse(build_cover.call_args.kwargs["cover_form_protection"])

    def test_word_com_postprocess_is_called_only_when_flag_is_enabled(self):
        with tempfile.TemporaryDirectory() as td:
            input_path = os.path.join(td, "input.md")
            output_path = os.path.join(td, "output.docx")
            with open(input_path, "w", encoding="utf-8") as f:
                f.write("# 范围\n\n正文。\n")

            with mock.patch("md2std.docx_builder.build_cover", return_value=output_path), \
                 mock.patch("md2std.word_postprocess.postprocess_with_word_com", return_value=output_path) as post:
                self.assertEqual(cli.main([input_path, "-o", output_path]), 0)
                post.assert_not_called()

                self.assertEqual(cli.main([input_path, "-o", output_path, "--word-com-postprocess"]), 0)
                post.assert_called_once_with(output_path)

    def test_cli_reports_untitled_clause_in_scope_before_building(self):
        with tempfile.TemporaryDirectory() as td:
            input_path = os.path.join(td, "input.md")
            output_path = os.path.join(td, "output.docx")
            with open(input_path, "w", encoding="utf-8") as f:
                f.write("# 范围\n\n{无标题条:2} 本文件规定了测试要求。\n")

            with mock.patch("sys.stderr"), \
                 mock.patch("md2std.docx_builder.build_cover") as build_cover, \
                 self.assertRaises(SystemExit) as cm:
                cli.main([input_path, "-o", output_path])

        self.assertEqual(cm.exception.code, 2)
        build_cover.assert_not_called()

    def test_word_postprocess_replaces_target_only_after_temp_copy_success(self):
        with tempfile.TemporaryDirectory() as td:
            target = os.path.join(td, "target.docx")
            with open(target, "wb") as f:
                f.write(b"original")

            class FakeWord:
                Hwnd = 0

                def __init__(self):
                    self.quit_count = 0

                def Quit(self):
                    self.quit_count += 1

            word = FakeWord()
            seen_work_paths = []

            def fake_process(_word, work_path):
                seen_work_paths.append(work_path)
                self.assertNotEqual(os.path.abspath(target), work_path)
                with open(work_path, "wb") as f:
                    f.write(b"processed")

            with mock.patch("md2std.word_postprocess._postprocess_document", fake_process):
                word_postprocess._postprocess_with_word_instance(
                    os.path.abspath(target),
                    word,
                    timeout_seconds=0,
                )

            with open(target, "rb") as f:
                self.assertEqual(f.read(), b"processed")
            self.assertEqual(word.quit_count, 1)
            self.assertEqual(self._temp_wordcom_files(td), [])
            self.assertEqual(len(seen_work_paths), 1)

    def test_word_postprocess_failure_keeps_target_and_removes_temp_copy(self):
        with tempfile.TemporaryDirectory() as td:
            target = os.path.join(td, "target.docx")
            with open(target, "wb") as f:
                f.write(b"original")

            class FakeWord:
                Hwnd = 0

                def Quit(self):
                    pass

            def fake_process(_word, work_path):
                with open(work_path, "wb") as f:
                    f.write(b"partial")
                raise RuntimeError("boom")

            with mock.patch("md2std.word_postprocess._postprocess_document", fake_process):
                with self.assertRaisesRegex(RuntimeError, "boom"):
                    word_postprocess._postprocess_with_word_instance(
                        os.path.abspath(target),
                        FakeWord(),
                        timeout_seconds=0,
                    )

            with open(target, "rb") as f:
                self.assertEqual(f.read(), b"original")
            self.assertEqual(self._temp_wordcom_files(td), [])

    def _temp_wordcom_files(self, directory: str) -> list:
        return sorted(
            name for name in os.listdir(directory)
            if name.startswith(".md2std-wordcom-")
        )


if __name__ == "__main__":
    unittest.main()
