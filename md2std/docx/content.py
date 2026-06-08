# -*- coding: utf-8 -*-
"""DOCX content emitters for front matter, body, annexes, references, and index."""

from __future__ import annotations

import io
import os
import re
from typing import List, Optional

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .. import boilerplate as bp
from .. import model
from .. import styles as S
from .oxml import *
from .state import _needs_seq_reset, _next_bm_id

def _emit_foreword_content(anchor: Paragraph, doc, meta: model.Meta):
    fw = meta.foreword

    def add(text):
        if text:
            _new_paragraph_before(anchor, doc, S.S_PARA, text=text)

    def add_extra_note(note):
        if isinstance(note, (list, tuple)):
            for item in note:
                text = str(item).strip()
                if text.startswith("- "):
                    text = text[2:].strip()
                if text:
                    _new_numbered_style_paragraph(anchor, doc, S.S_LIST_DASH, text=text)
            return
        add(str(note).strip())

    add(bp.FOREWORD_FIRST)
    if fw.multipart_note:
        add(fw.multipart_note)
    if fw.replace_changes:
        # 技术变化导语 + 破折号列项
        if meta.replaces:
            lead = ("本文件代替%s，与%s相比，除结构调整和编辑性改动外，"
                    "主要技术变化如下：" % (meta.replaces, meta.replaces))
        else:
            lead = "与上一版相比，除结构调整和编辑性改动外，主要技术变化如下："
        add(lead)
        for ch in fw.replace_changes:
            _new_numbered_style_paragraph(anchor, doc, S.S_LIST_DASH, text=ch)
    if fw.patent_note:
        add(bp.PATENT_NOTE)
    # 提出 / 归口
    if fw.proposer and (not fw.owner or fw.owner == fw.proposer):
        add(bp.PROPOSE_OWNER_SAME.format(org=fw.proposer))
    else:
        if fw.proposer:
            add(bp.PROPOSE.format(org=fw.proposer))
        if fw.owner:
            add(bp.OWNER.format(org=fw.owner))
    if fw.draft_orgs:
        add(bp.DRAFT_ORGS.format(orgs="、".join(fw.draft_orgs)))
    if fw.drafters:
        add(bp.DRAFTERS.format(people="、".join(fw.drafters)))
    if fw.history:
        add(bp.HISTORY_LEAD)
        add(fw.history)
    for note in fw.extra_notes:
        add_extra_note(note)


def _emit_introduction_content(anchor: Paragraph, doc, meta: model.Meta):
    for line in meta.introduction.splitlines():
        if line.strip():
            _new_paragraph_before(anchor, doc, S.S_PARA, text=line.strip())


_TERM_SPLIT_RE = re.compile(r"^(.*?)[\s　]{2,}(.+)$")
# 规范性引用文件条目："标准号  标准名称"（标准号如 GB 5749 / GB/T 5750.3 / DZ/T 0225）
_NORMREF_RE = re.compile(r"^\s*([A-Z][A-Z/]*\s+\d[\w.\-—–]*)(?:\s{2,}|　+)(.+)$")


def _emit_normative_ref(anchor: Paragraph, doc, spans):
    """规范性引用文件条目：把标准号加书签，便于交叉引用插入"GB 5749"。"""
    text = "".join(s.text for s in spans).strip()
    m = _NORMREF_RE.match(text)
    if not m:
        _new_paragraph_before(anchor, doc, S.S_PARA, spans=spans)
        return
    stdno, name = m.group(1).strip(), m.group(2).strip()
    para = _new_paragraph_before(anchor, doc, S.S_PARA)
    bid = _COUNTER.bm
    _COUNTER.bm += 1
    _bookmark_start(para, _bm_name(stdno), bid)
    para.add_run(stdno)
    _bookmark_end(para, bid)
    para.add_run("  " + name)


def _emit_source(anchor: Paragraph, doc, source: model.Source):
    para = _new_paragraph_before(anchor, doc, S.S_PARA)
    run = para.add_run(source.text)
    try:
        run.style = doc.styles[S.S_SOURCE]
    except KeyError:
        pass


def _split_term_text(text: str):
    m = _TERM_SPLIT_RE.match((text or "").strip())
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return (text or "").strip(), ""


def _emit_term_title(anchor: Paragraph, doc, term: str, term_en: str = ""):
    para = _new_paragraph_before(anchor, doc, S.S_TERM_1)
    _set_numbering(para, S.NUM_BODY, 2)
    para.add_run().add_break()          # 让自动编号 3.1 单独占一行
    r = para.add_run(term)
    r.bold = True
    if term_en:
        para.add_run("　")
        er = para.add_run(term_en)
        er.bold = True
    return para


def _emit_term(anchor: Paragraph, doc, spans):
    """术语条目：编号(3.1)单独成行，下一行加粗中文术语 + 英文对应词。

    自动编号接入 numId=2 ilvl=2（术语为章 3 下的一级条）。
    """
    cn, en = _split_term_text("".join(s.text for s in spans))
    return _emit_term_title(anchor, doc, cn, en)


def _emit_term_entry(anchor: Paragraph, doc, term: model.Term):
    _emit_term_title(anchor, doc, term.term, term.term_en)
    if term.definition:
        _new_paragraph_before(anchor, doc, S.S_PARA, spans=term.definition)
    _emit_note_group(anchor, doc, term.notes)
    if term.source is not None:
        _emit_source(anchor, doc, term.source)


def _note_group_requires_numbering(notes: List[model.Note]) -> bool:
    return len(notes) > 1 or any(note.index is not None for note in notes)


def _emit_note_group(anchor: Paragraph, doc, notes: List[model.Note]):
    if not notes:
        return
    if not _note_group_requires_numbering(notes):
        _new_paragraph_before(anchor, doc, S.S_NOTE, spans=notes[0].spans)
        return
    num_id = _new_numbering_instance_from_style(doc, S.S_NOTE_X)
    for note in notes:
        _new_numbered_style_paragraph(
            anchor,
            doc,
            S.S_NOTE_X,
            spans=note.spans,
            num_id_override=num_id,
        )


def _emit_body_block(anchor: Paragraph, doc, blk, in_terms=False, in_normrefs=False,
                     appendix_letter=None, list_num_id: Optional[int] = None):
    """把一个正文块插入到 anchor 之前。"""
    if isinstance(blk, model.Heading):
        if in_terms and blk.level == 2:
            _emit_term(anchor, doc, blk.spans)
            return
        style = S.HEADING_STYLE_BY_LEVEL.get(blk.level, S.S_PARA)
        _new_paragraph_before(anchor, doc, style, spans=blk.spans)
    elif isinstance(blk, model.Term):
        _emit_term_entry(anchor, doc, blk)
    elif isinstance(blk, model.Paragraph):
        if in_normrefs:
            _emit_normative_ref(anchor, doc, blk.spans)
        else:
            _new_paragraph_before(anchor, doc, S.S_PARA, spans=blk.spans)
    elif isinstance(blk, model.UntitledClause):
        # 无标题条：套用"X级无标题"样式 + 接入正文多级列表(numId=2)自动编号，
        # Markdown 只声明层级；ilvl = 层级（3 -> 渲染为 "4.2.1"）。
        seg = blk.segments
        style = S.UNTITLED_STYLE_BY_SEGMENTS.get(seg, S.S_PARA)
        para = _new_paragraph_before(anchor, doc, style, spans=list(blk.spans))
        if seg in S.UNTITLED_STYLE_BY_SEGMENTS:
            _set_numbering(para, S.NUM_BODY, seg)
    elif isinstance(blk, model.Note):
        _emit_note_group(anchor, doc, [blk])
    elif isinstance(blk, model.Example):
        style = S.S_EXAMPLE
        _new_paragraph_before(anchor, doc, style, spans=blk.spans)
    elif isinstance(blk, model.ExampleContent):
        _new_paragraph_before(anchor, doc, S.S_EXAMPLE_CONTENT, spans=blk.spans)
    elif isinstance(blk, model.PageBreak):
        _emit_page_break(anchor, doc)
    elif isinstance(blk, model.Source):
        _emit_source(anchor, doc, blk)
    elif isinstance(blk, model.ListBlock):
        if blk.ordered:
            style = S.ORDERED_LIST_STYLE_BY_LEVEL.get(blk.level, S.S_LIST_NUMBER_3)
        else:
            style = S.S_LIST_DASH
        for it in blk.items:
            _new_numbered_style_paragraph(
                anchor, doc, style, spans=it.spans,
                num_id_override=list_num_id if blk.ordered else None,
            )
    elif isinstance(blk, model.TableModel):
        _emit_table(anchor, doc, blk, appendix_letter=appendix_letter)
    elif isinstance(blk, model.Figure):
        _emit_figure(anchor, doc, blk, appendix_letter=appendix_letter)
    elif isinstance(blk, model.Formula):
        fstyle = S.S_FORMULA_APPENDIX if appendix_letter else S.S_FORMULA
        _emit_formula(anchor, doc, blk, style=fstyle, appendix_letter=appendix_letter)


def _emit_body_standard_title(anchor: Paragraph, doc, meta: model.Meta):
    title = (meta.title or "").strip()
    if not title:
        return
    para = _new_paragraph_before(anchor, doc, S.S_BODY_STANDARD_NAME)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, line in enumerate([x.strip() for x in title.splitlines() if x.strip()]):
        if i > 0:
            para.add_run().add_break()
        run = para.add_run(line)
        run.bold = True


_IMPORTANT_NOTICE_PREFIX_RE = re.compile(r"^(?:重要提示|危险|警告|注意)\s*[:：]")


def _emit_important_notice(anchor: Paragraph, doc, meta: model.Meta):
    lines = [x.strip() for x in (meta.important_notice or "").splitlines() if x.strip()]
    if not lines:
        return
    for i, line in enumerate(lines):
        text = line
        if i == 0 and not _IMPORTANT_NOTICE_PREFIX_RE.match(text):
            text = "重要提示：" + text
        _new_paragraph_before(anchor, doc, S.S_IMPORTANT_NOTICE, text=text)


def _emit_formula(anchor: Paragraph, doc, formula, style=None, appendix_letter=None):
    """公式段：用制表位——[Tab]公式[Tab]（序号）。

    模板"标准文件_正文公式"已配好：居中制表位让公式居中，右制表位(带点引导)推序号。
    序号括在"（""）"内，括号为纯文本；书签仅围住不带括号的编号文字。
    SEQ 域使用中文前缀"公式"，使公式出现在 Word 交叉引用"公式"域。
    """
    from .. import mathconv
    _emit_formula_caption_anchor(anchor, doc, formula, appendix_letter=appendix_letter)
    para = _new_paragraph_before(anchor, doc, style or S.S_FORMULA)
    # 居中制表位：公式居中
    para.add_run("\t")
    # 原生公式（OMML）
    omath = mathconv.latex_to_omml(formula.latex)
    if omath is not None:
        para._p.append(omath)
    else:
        para.add_run(formula.latex)
    para.add_run("\t")
    _emit_formula_number(para, formula, appendix_letter=appendix_letter)


def _add_bookmark_ends_after(run, bids):
    """按 bids 给出的最终顺序，把多个 bookmarkEnd 插到同一个 run 后。"""
    for bid in reversed([b for b in bids if b is not None]):
        _bookmark_end_after(run, bid)


def _emit_style_numbered_caption(para: Paragraph, doc, style_name: str,
                                 ref_type: str, anchor_id: str, title: str):
    """用模板标题样式的自动编号输出表/图题，并建立交叉引用书签。"""
    _normalize_caption_numbering_separator(doc, style_name)
    _set_numbering_from_style(para, doc, style_name)
    if title:
        para.add_run("　")
    title_run = para.add_run(title or "")
    if not anchor_id:
        return

    bookmark_ids = [
        (_native_ref_name(ref_type, anchor_id, "full"), _next_bm_id()),
        (_native_ref_name(ref_type, anchor_id, "label"), _next_bm_id()),
        (_native_ref_name(ref_type, anchor_id, "num"), _next_bm_id()),
        (_native_ref_name(ref_type, anchor_id, "text"), _next_bm_id()),
    ]
    for name, bid in bookmark_ids:
        _bookmark_start_before(title_run, name, bid)
    _add_bookmark_ends_after(title_run, [bid for _, bid in bookmark_ids])


def _apply_character_style(doc, run, style_name: str):
    try:
        run.style = doc.styles[style_name]
    except Exception:
        _apply_style_run_properties(doc, run, style_name)


def _add_character_styled_runs(paragraph: Paragraph, doc, style_name: str,
                               spans: List[model.Span]):
    from .. import mathconv
    for sp in spans:
        if isinstance(sp, model.RefSpan):
            _add_typed_ref(paragraph, sp)
            continue
        if isinstance(sp, model.FormulaSpan):
            omath = mathconv.latex_to_omml(sp.text)
            if omath is not None:
                paragraph._p.append(omath)
            else:
                run = paragraph.add_run(sp.text)
                _apply_character_style(doc, run, style_name)
            continue
        for i, piece in enumerate(sp.text.split("\n")):
            if i > 0:
                paragraph.add_run().add_break()
            if piece:
                run = paragraph.add_run(piece)
                _apply_character_style(doc, run, style_name)
                if sp.bold:
                    run.bold = True
                if sp.italic:
                    run.italic = True
                if getattr(sp, "subscript", False):
                    run.font.subscript = True
                if getattr(sp, "superscript", False):
                    run.font.superscript = True


def _footnote_ref_label(index: int) -> str:
    """Return a, b, ..., z, aa, ab... for generated table footnote references."""
    letters = []
    value = max(0, index)
    while True:
        value, rem = divmod(value, 26)
        letters.append(chr(ord("a") + rem))
        if value == 0:
            break
        value -= 1
    return "".join(reversed(letters))


def _emit_footnote_runs(para: Paragraph, doc, footnote: model.FigureTableFootnote,
                        num_id: Optional[int] = None):
    _set_paragraph_style(doc, para, S.S_FIG_TABLE_NOTE)
    _set_numbering_from_style(para, doc, S.S_FIG_TABLE_NOTE, num_id_override=num_id)
    _set_runs(para, footnote.spans)


def _emit_figure_table_footnote(anchor: Paragraph, doc, footnote: model.FigureTableFootnote,
                                num_id: Optional[int] = None):
    para = _new_paragraph_before(anchor, doc, S.S_FIG_TABLE_NOTE)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _emit_footnote_runs(para, doc, footnote, num_id=num_id)


def _emit_figure_table_source(anchor: Paragraph, doc, source: model.FigureTableSource):
    para = _new_paragraph_before(anchor, doc, S.S_FIG_TABLE_SOURCE)
    para.add_run("来源：")
    _set_runs(para, source.spans)


def _emit_figure_table_unit(anchor: Paragraph, doc, unit: model.FigureTableSource):
    para = _new_paragraph_before(anchor, doc, S.S_FIG_TABLE_SOURCE)
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_runs(para, unit.spans)


def _set_keep_next_on_paragraph_element(p):
    if p is None or p.tag != qn("w:p"):
        return
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    if ppr.find(qn("w:keepNext")) is None:
        ppr.append(OxmlElement("w:keepNext"))


def _keep_previous_block_with_paragraph(para: Paragraph):
    prev = para._p.getprevious()
    if prev is None:
        return
    if prev.tag == qn("w:p"):
        _set_keep_next_on_paragraph_element(prev)
        return
    if prev.tag == qn("w:tbl"):
        paragraphs = list(prev.iter(qn("w:p")))
        if paragraphs:
            _set_keep_next_on_paragraph_element(paragraphs[-1])


def _collapse_hidden_paragraph(para: Paragraph):
    ppr = para._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "1")
    spacing.set(qn("w:lineRule"), "exact")


def _emit_hidden_appendix_caption_label(anchor: Paragraph, doc, style_name: str):
    para = _new_paragraph_before(anchor, doc, style_name)
    _set_numbering_from_style(para, doc, style_name)
    _make_paragraph_hidden(para)
    _collapse_hidden_paragraph(para)


def _available_figure_width_emu(doc) -> int:
    fallback = 9360 * 635
    for section in reversed(doc.sections):
        usable = int(section.page_width - section.left_margin - section.right_margin)
        if usable > 0:
            return usable
    return fallback


def _fit_inline_shape_to_width(shape, max_width_emu: int):
    if max_width_emu <= 0 or int(shape.width) <= max_width_emu:
        return
    scale = max_width_emu / float(shape.width)
    shape.width = max_width_emu
    shape.height = int(shape.height * scale)


def _format_figure_image_paragraph(para: Paragraph):
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ppr = para._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    spacing.set(qn("w:before"), "120")
    spacing.set(qn("w:after"), "120")


def _emu_to_twips(value: int) -> int:
    return max(1, int(value / 635))


def _set_table_width(table, width_emu: int):
    tblpr = table._tbl.find(qn("w:tblPr"))
    if tblpr is None:
        tblpr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tblpr)
    old = tblpr.find(qn("w:tblW"))
    if old is not None:
        tblpr.remove(old)
    tblw = OxmlElement("w:tblW")
    tblw.set(qn("w:w"), str(_emu_to_twips(width_emu)))
    tblw.set(qn("w:type"), "dxa")
    tblpr.append(tblw)


def _set_cell_width(cell, width_emu: int):
    tcpr = cell._tc.get_or_add_tcPr()
    old = tcpr.find(qn("w:tcW"))
    if old is not None:
        tcpr.remove(old)
    tcw = OxmlElement("w:tcW")
    tcw.set(qn("w:w"), str(_emu_to_twips(width_emu)))
    tcw.set(qn("w:type"), "dxa")
    tcpr.append(tcw)


def _set_table_no_borders(table):
    tblpr = table._tbl.find(qn("w:tblPr"))
    if tblpr is None:
        tblpr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tblpr)
    old = tblpr.find(qn("w:tblBorders"))
    if old is not None:
        tblpr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement("w:" + edge)
        node.set(qn("w:val"), "nil")
        node.set(qn("w:sz"), "0")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "auto")
        borders.append(node)
    tblpr.append(borders)


def _emit_subfigure_cell(cell, doc, subfig: model.FigureSubfigure, max_width_emu: int):
    _set_cell_vertical_center(cell)
    _set_cell_margins(cell, top=40, bottom=40, left=60, right=60)
    para = _reset_cell_to_single_paragraph(cell, doc, S.S_NORMAL)
    _format_figure_image_paragraph(para)
    if subfig.path and os.path.isfile(subfig.path):
        try:
            shape = para.add_run().add_picture(subfig.path)
            _fit_inline_shape_to_width(shape, max_width_emu)
            return
        except Exception:
            pass
    para.add_run("[缺少图片：%s]" % subfig.path)


def _emit_subfigure_caption_cell(cell, doc, label: str, caption: str):
    _set_cell_vertical_center(cell)
    _set_cell_margins(cell, top=0, bottom=80, left=60, right=60)
    para = _reset_cell_to_single_paragraph(cell, doc, S.S_NORMAL)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    text = label + ")"
    if caption:
        text += "　" + caption
    para.add_run(text)


def _emit_subfigures(anchor: Paragraph, doc, subfigures: List[model.FigureSubfigure],
                     columns: int = 0):
    if not subfigures:
        return
    cols = columns if columns > 0 else (1 if len(subfigures) == 1 else 2)
    cols = max(1, min(cols, len(subfigures)))
    usable_width = _available_figure_width_emu(doc)
    cell_width = int(usable_width / cols)
    image_max_width = max(1, cell_width - 220000)
    table = doc.add_table(rows=0, cols=cols)
    _set_table_width(table, usable_width)
    _set_table_no_borders(table)

    for start in range(0, len(subfigures), cols):
        chunk = subfigures[start:start + cols]
        image_row = table.add_row()
        caption_row = table.add_row()
        for c_i in range(cols):
            image_cell = image_row.cells[c_i]
            caption_cell = caption_row.cells[c_i]
            _set_cell_width(image_cell, cell_width)
            _set_cell_width(caption_cell, cell_width)
            if c_i < len(chunk):
                subfig = chunk[c_i]
                _emit_subfigure_cell(image_cell, doc, subfig, image_max_width)
                label = _footnote_ref_label(start + c_i)
                _emit_subfigure_caption_cell(caption_cell, doc, label, subfig.caption)
            else:
                image_cell.text = ""
                caption_cell.text = ""
    anchor._p.addprevious(table._tbl)


def _format_figure_body_paragraph(para: Paragraph):
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_paragraph_left_indent(para, 0)


def _emit_figure_key_items(anchor: Paragraph, doc, items: List[model.FigureKeyItem]):
    if not items:
        return
    lead = _new_paragraph_before(anchor, doc, S.S_PARA, text="标引序号说明：")
    _format_figure_body_paragraph(lead)
    for item in items:
        para = _new_paragraph_before(anchor, doc, S.S_PARA)
        _format_figure_body_paragraph(para)
        para.add_run(item.index)
        para.add_run("——")
        _set_runs(para, item.spans)


def _emit_figure_body_paragraphs(anchor: Paragraph, doc, paragraphs: List[model.FigureBodyParagraph]):
    for fig_para in paragraphs:
        para = _new_paragraph_before(anchor, doc, S.S_PARA, spans=fig_para.spans)
        _format_figure_body_paragraph(para)
        _emit_note_group(anchor, doc, fig_para.notes)


def _emit_figure_addons(anchor: Paragraph, doc, fig: model.Figure):
    _emit_figure_key_items(anchor, doc, fig.key_items)
    _emit_figure_body_paragraphs(anchor, doc, fig.body_paragraphs)
    footnote_num_id = (
        _new_numbering_instance_from_style(doc, S.S_FIG_TABLE_NOTE)
        if fig.footnotes else None
    )
    for footnote in fig.footnotes:
        _emit_figure_table_footnote(anchor, doc, footnote, num_id=footnote_num_id)
    if fig.source is not None:
        _emit_figure_table_source(anchor, doc, fig.source)


def _emit_formula_number(para: Paragraph, formula: model.Formula, appendix_letter=None):
    """输出公式右侧编号 `（1）` / `（A.1）`，显示值引用隐藏公式题注。"""
    para.add_run("（")
    placeholder = "1"
    if appendix_letter:
        placeholder = appendix_letter + ".1"
    if formula.anchor_id:
        _add_ref_bookmark(
            para,
            _native_ref_name("eq", formula.anchor_id, "num"),
            display_text=placeholder,
            charformat=True,
        )
    else:
        para.add_run(placeholder)
    para.add_run("）")


def _emit_formula_caption_anchor(anchor: Paragraph, doc, formula: model.Formula, appendix_letter=None):
    """插入一个隐藏公式题注段，供 Word 交叉引用列表和 REF 域使用。"""
    if not formula.anchor_id:
        return
    cap = _new_paragraph_before(anchor, doc, S.S_NORMAL)
    try:
        cap.style = doc.styles["Caption"]
    except KeyError:
        pass
    _make_paragraph_hidden(cap)

    label_run = cap.add_run("公式")
    _make_run_hidden(label_run._r)
    bm_label_id = _next_bm_id()
    bm_num_id = _next_bm_id()
    _bookmark_start_before(label_run, _native_ref_name("eq", formula.anchor_id, "label"), bm_label_id)
    space_run = cap.add_run(" ")
    _make_run_hidden(space_run._r)
    prefix_run = None
    if appendix_letter:
        prefix_run = cap.add_run(appendix_letter + ".")
        _make_run_hidden(prefix_run._r)
    result_run, seq_end_run = _add_seq(
        cap,
        SEQ_EQUATION,
        reset=_needs_seq_reset("eq", appendix_letter),
        hidden=True,
    )
    _bookmark_start_before(prefix_run or result_run, _native_ref_name("eq", formula.anchor_id, "num"), bm_num_id)
    _add_bookmark_ends_after(seq_end_run, [bm_num_id, bm_label_id])


# 不按行数预拆续表。是否跨页只有 Word/LibreOffice 完成分页后才知道；
# 生成阶段误拆会产生同页“续表”，因此默认只生成一个真实表格。
_TABLE_SPLIT_THRESHOLD = None
_TABLE_SPLIT_CHUNK_SIZE = None


def _set_cell_vertical_center(cell):
    tcpr = cell._tc.get_or_add_tcPr()
    old = tcpr.find(qn("w:vAlign"))
    if old is not None:
        tcpr.remove(old)
    valign = OxmlElement("w:vAlign")
    valign.set(qn("w:val"), "center")
    tcpr.append(valign)


def _set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcpr = cell._tc.get_or_add_tcPr()
    old = tcpr.find(qn("w:tcMar"))
    if old is not None:
        tcpr.remove(old)
    tcmar = OxmlElement("w:tcMar")
    for edge, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement("w:" + edge)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tcmar.append(node)
    tcpr.append(tcmar)


_BORDER_SIZE_BY_NAME = {
    "thin": 4,
    "thick": 8,
}


def _append_border(parent, edge: str, value: str):
    value = (value or "").strip().lower()
    node = OxmlElement("w:" + edge)
    if value == "none":
        node.set(qn("w:val"), "nil")
        node.set(qn("w:sz"), "0")
    else:
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), str(_BORDER_SIZE_BY_NAME.get(value, 4)))
    node.set(qn("w:space"), "0")
    node.set(qn("w:color"), "000000")
    parent.append(node)


def _set_table_borders(table, outer: str = "thick", inner: str = "thin"):
    tblpr = table._tbl.find(qn("w:tblPr"))
    if tblpr is None:
        tblpr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tblpr)
    old = tblpr.find(qn("w:tblBorders"))
    if old is not None:
        tblpr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        _append_border(borders, edge, inner if edge.startswith("inside") else outer)
    tblpr.append(borders)


def _set_cell_borders(cell, borders: dict):
    if not borders:
        return
    tcpr = cell._tc.get_or_add_tcPr()
    old = tcpr.find(qn("w:tcBorders"))
    if old is not None:
        tcpr.remove(old)
    node = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        value = borders.get(edge)
        if value:
            _append_border(node, edge, value)
    tcpr.append(node)


def _set_row_repeat_header(row):
    trpr = row._tr.get_or_add_trPr()
    if trpr.find(qn("w:tblHeader")) is None:
        hdr = OxmlElement("w:tblHeader")
        hdr.set(qn("w:val"), "true")
        trpr.append(hdr)


def _cell_is_long_text(text: str, colspan: int = 1) -> bool:
    text = (text or "").strip()
    return colspan > 1 or len(text) > 18 or any(mark in text for mark in "，。；：、（）()")


def _parts_from_text(text: str) -> List[model.TableCellPart]:
    return [model.TableCellPart("text", text or "")]


def _emit_table_cell_parts(paragraph: Paragraph, doc, parts: List[model.TableCellPart],
                           footnote_ref_state=None):
    from .. import mathconv
    if not parts:
        return
    for part in parts:
        if part.kind == "note":
            continue
        if part.kind == "footnote_ref":
            label = part.text
            if footnote_ref_state is not None:
                label = _footnote_ref_label(footnote_ref_state["next"])
                footnote_ref_state["next"] += 1
            if label:
                run = paragraph.add_run(label)
                _apply_character_style(doc, run, S.S_FIG_TABLE_NOTE_CONTENT)
                run.font.superscript = True
            continue
        if part.kind == "ref":
            _add_typed_ref(paragraph, model.RefSpan(
                text=part.text,
                ref_type=part.ref_type,
                target=part.target,
                mode=part.mode,
            ))
            continue
        if part.kind == "formula":
            omath = mathconv.latex_to_omml(part.text)
            if omath is not None:
                paragraph._p.append(omath)
            else:
                paragraph.add_run(part.text)
            continue
        for i, piece in enumerate((part.text or "").split("\n")):
            if i > 0:
                paragraph.add_run().add_break()
            if piece:
                run = paragraph.add_run(piece)
                if getattr(part, "subscript", False):
                    run.font.subscript = True
                if getattr(part, "superscript", False):
                    run.font.superscript = True


def _set_paragraph_style(doc, para: Paragraph, style_name: str):
    try:
        para.style = doc.styles[style_name]
    except KeyError:
        pass


def _set_paragraph_left_indent(para: Paragraph, left_twips: int):
    ppr = para._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:left"), str(left_twips))
    if ind.get(qn("w:firstLine")) is not None:
        del ind.attrib[qn("w:firstLine")]
    if ind.get(qn("w:firstLineChars")) is not None:
        del ind.attrib[qn("w:firstLineChars")]


def _set_paragraph_first_line_indent(para: Paragraph, first_line_twips: int, first_line_chars: int):
    ppr = para._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    for attr in ("left", "leftChars", "hanging", "hangingChars"):
        key = qn("w:" + attr)
        if ind.get(key) is not None:
            del ind.attrib[key]
    ind.set(qn("w:firstLine"), str(first_line_twips))
    ind.set(qn("w:firstLineChars"), str(first_line_chars))


def _reset_cell_to_single_paragraph(cell, doc, style_name: str) -> Paragraph:
    cell.text = ""
    para = cell.paragraphs[0]
    _set_paragraph_style(doc, para, style_name)
    return para


def _format_table_footer_paragraph(para: Paragraph, doc, indent: bool = True):
    _set_paragraph_style(doc, para, S.S_TABLE_CELL)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if indent:
        _set_paragraph_left_indent(para, 420)


def _split_table_cell_note_parts(parts: List[model.TableCellPart]):
    body_parts: List[model.TableCellPart] = []
    notes: List[List[model.Span]] = []
    for part in parts:
        if part.kind == "note":
            notes.append(part.spans)
        else:
            body_parts.append(part)
    return body_parts, notes


def _table_cell_parts_have_body(parts: List[model.TableCellPart]) -> bool:
    for part in parts:
        if part.kind == "text":
            if part.text.strip():
                return True
            continue
        if part.kind != "note":
            return True
    return False


def _format_table_cell_paragraph_with_notes(para: Paragraph, doc):
    _set_paragraph_style(doc, para, S.S_PARA)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_paragraph_first_line_indent(para, 199, 95)


def _emit_table_cell_notes(cell, doc, note_spans: List[List[model.Span]],
                           first_para: Optional[Paragraph] = None):
    if not note_spans:
        return
    numbered = len(note_spans) > 1
    num_id = _new_numbering_instance_from_style(doc, S.S_NOTE_X) if numbered else None
    for idx, spans in enumerate(note_spans):
        para = first_para if idx == 0 and first_para is not None else cell.add_paragraph()
        if numbered:
            _set_paragraph_style(doc, para, S.S_NOTE_X)
            _set_numbering_from_style(para, doc, S.S_NOTE_X, num_id_override=num_id)
        else:
            _set_paragraph_style(doc, para, S.S_NOTE)
        _set_runs(para, spans)


def _emit_table_cell_content(word_cell, doc, cell_model: model.TableCell,
                             footnote_ref_state=None):
    _set_cell_vertical_center(word_cell)
    _set_cell_margins(word_cell)
    _set_cell_borders(word_cell, cell_model.borders)
    cp = word_cell.paragraphs[0]
    parts = cell_model.parts or _parts_from_text(cell_model.text)
    body_parts, note_spans = _split_table_cell_note_parts(parts)
    if note_spans:
        if _table_cell_parts_have_body(body_parts):
            _format_table_cell_paragraph_with_notes(cp, doc)
            _emit_table_cell_parts(cp, doc, body_parts, footnote_ref_state)
            _emit_table_cell_notes(word_cell, doc, note_spans)
        else:
            _emit_table_cell_notes(word_cell, doc, note_spans, first_para=cp)
        return
    _set_paragraph_style(doc, cp, S.S_TABLE_CELL)
    cp.alignment = (
        WD_ALIGN_PARAGRAPH.LEFT
        if not cell_model.header and _cell_is_long_text(cell_model.text, cell_model.colspan)
        else WD_ALIGN_PARAGRAPH.CENTER
    )
    _emit_table_cell_parts(cp, doc, parts, footnote_ref_state)


def _emit_table_footer_footnotes(table, doc, footnotes: List[model.FigureTableFootnote]):
    if not footnotes:
        return
    num_id = _new_numbering_instance_from_style(doc, S.S_FIG_TABLE_NOTE)
    row = table.add_row()
    cell = row.cells[0]
    if len(row.cells) > 1:
        cell = cell.merge(row.cells[-1])
    _set_cell_vertical_center(cell)
    _set_cell_margins(cell)
    para = _reset_cell_to_single_paragraph(cell, doc, S.S_FIG_TABLE_NOTE)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _emit_footnote_runs(para, doc, footnotes[0], num_id=num_id)
    for footnote in footnotes[1:]:
        para = cell.add_paragraph()
        _set_paragraph_style(doc, para, S.S_FIG_TABLE_NOTE)
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _emit_footnote_runs(para, doc, footnote, num_id=num_id)


def _emit_table_footer_source(table, doc, source: model.FigureTableSource):
    row = table.add_row()
    cell = row.cells[0]
    if len(row.cells) > 1:
        cell = cell.merge(row.cells[-1])
    _set_cell_vertical_center(cell)
    _set_cell_margins(cell)
    para = _reset_cell_to_single_paragraph(cell, doc, S.S_FIG_TABLE_SOURCE)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.first_line_indent = 0
    para.add_run("来源：")
    _set_runs(para, source.spans)


def _emit_table_footer_addons(table, doc, tbl: model.TableModel):
    _emit_table_footer_footnotes(table, doc, tbl.footnotes)
    if tbl.source is not None:
        _emit_table_footer_source(table, doc, tbl.source)


def _emit_table_continuation_caption(anchor: Paragraph, doc, tbl: model.TableModel,
                                     appendix_letter=None):
    style = "标准文件_表格续"
    cap = _new_paragraph_before(anchor, doc, style)
    _set_numbering(cap, 0, 0)
    cap.paragraph_format.page_break_before = True
    if tbl.anchor_id:
        _add_ref_bookmark(
            cap,
            _native_ref_name("tbl", tbl.anchor_id, "label"),
            display_text="表?",
            switches="\\r",
        )
    else:
        cap.add_run("表")
    title = (tbl.caption or "").strip()
    if title:
        cap.add_run("　")
        cap.add_run(title)
    cap.add_run("（续）")


def _legacy_cell_rows(tbl: model.TableModel, rows, row_parts, row_colspans):
    cell_rows: List[List[model.TableCell]] = []
    if tbl.header:
        header_parts = tbl.header_parts or [_parts_from_text(text) for text in tbl.header]
        header_spans = tbl.header_colspans or [1 for _ in tbl.header]
        cell_rows.append([
            model.TableCell(
                text=text,
                parts=header_parts[i] if i < len(header_parts) else _parts_from_text(text),
                colspan=header_spans[i] if i < len(header_spans) else 1,
                header=True,
            )
            for i, text in enumerate(tbl.header)
        ])
    for r_i, row in enumerate(rows):
        parts_row = row_parts[r_i] if r_i < len(row_parts) else []
        spans = row_colspans[r_i] if r_i < len(row_colspans) and row_colspans[r_i] else [1 for _ in row]
        cell_rows.append([
            model.TableCell(
                text=text,
                parts=parts_row[i] if i < len(parts_row) else _parts_from_text(text),
                colspan=spans[i] if i < len(spans) else 1,
            )
            for i, text in enumerate(row)
        ])
    return cell_rows


def _table_cell_layout(cell_rows: List[List[model.TableCell]]):
    occupied = {}
    placements = []
    max_cols = 0
    for r_i, row in enumerate(cell_rows):
        row_placements = []
        c_i = 0
        for cell in row:
            while occupied.get((r_i, c_i)):
                c_i += 1
            rowspan = max(1, int(cell.rowspan or 1))
            colspan = max(1, int(cell.colspan or 1))
            row_placements.append((cell, r_i, c_i, rowspan, colspan))
            for rr in range(r_i, r_i + rowspan):
                for cc in range(c_i, c_i + colspan):
                    occupied[(rr, cc)] = True
            max_cols = max(max_cols, c_i + colspan)
            c_i += colspan
        placements.append(row_placements)
    return placements, max(1, max_cols)


def _emit_table_part(anchor: Paragraph, doc, tbl: model.TableModel, rows, row_parts, row_colspans,
                     appendix_letter=None, include_addons: bool = False):
    """输出一个表格片段。续表复用原表头，不新增 SEQ。"""
    cell_rows = tbl.cell_rows or _legacy_cell_rows(tbl, rows, row_parts, row_colspans)
    placements, ncols = _table_cell_layout(cell_rows)
    table = doc.add_table(rows=len(cell_rows), cols=ncols)
    try:
        table.style = doc.styles[S.S_TABLE_GRID]
    except KeyError:
        pass
    _set_table_borders(table, tbl.border_outer or "thick", tbl.border_inner or "thin")
    footnote_ref_state = {"next": 0}

    for r_i, word_row in enumerate(table.rows):
        if r_i < tbl.header_row_count:
            _set_row_repeat_header(word_row)
    for row_placements in placements:
        for cell_model, r_i, c_i, rowspan, colspan in row_placements:
            word_cell = table.cell(r_i, c_i)
            if rowspan > 1 or colspan > 1:
                word_cell = word_cell.merge(table.cell(r_i + rowspan - 1, c_i + colspan - 1))
            _reset_cell_to_single_paragraph(word_cell, doc, S.S_TABLE_CELL)
            _emit_table_cell_content(word_cell, doc, cell_model, footnote_ref_state)
    if include_addons:
        _emit_table_footer_addons(table, doc, tbl)
    anchor._p.addprevious(table._tbl)


def _emit_table(anchor: Paragraph, doc, tbl: model.TableModel, appendix_letter=None):
    """表格：标题编号由模板表题样式自动生成。"""
    style = S.S_APPENDIX_TABLE_CAPTION if appendix_letter else S.S_TABLE_CAPTION
    cap = _new_paragraph_before(anchor, doc, style)
    _emit_style_numbered_caption(cap, doc, style, "tbl", tbl.anchor_id, tbl.caption or "")
    if tbl.unit is not None:
        _emit_figure_table_unit(anchor, doc, tbl.unit)

    row_parts = tbl.row_parts or [
        [_parts_from_text(cell) for cell in row]
        for row in tbl.rows
    ]
    row_colspans = tbl.row_colspans or [
        [1 for _ in row]
        for row in tbl.rows
    ]
    if _TABLE_SPLIT_THRESHOLD is not None and len(tbl.rows) > _TABLE_SPLIT_THRESHOLD:
        start = 0
        first = True
        while start < len(tbl.rows):
            end = min(start + _TABLE_SPLIT_CHUNK_SIZE, len(tbl.rows))
            if not first:
                _emit_table_continuation_caption(anchor, doc, tbl, appendix_letter=appendix_letter)
            _emit_table_part(
                anchor,
                doc,
                tbl,
                tbl.rows[start:end],
                row_parts[start:end],
                row_colspans[start:end],
                appendix_letter=appendix_letter,
                include_addons=end == len(tbl.rows),
            )
            first = False
            start = end
        return

    _emit_table_part(
        anchor, doc, tbl, tbl.rows, row_parts, row_colspans,
        appendix_letter=appendix_letter,
        include_addons=True,
    )


def _emit_figure(anchor: Paragraph, doc, fig: model.Figure, appendix_letter=None):
    if fig.unit is not None:
        _emit_figure_table_unit(anchor, doc, fig.unit)
    if fig.subfigures:
        _emit_subfigures(anchor, doc, fig.subfigures, columns=fig.subfigure_columns)
    else:
        img_p = _new_paragraph_before(anchor, doc, S.S_NORMAL)
        _format_figure_image_paragraph(img_p)
        if fig.path and os.path.isfile(fig.path):
            try:
                shape = img_p.add_run().add_picture(fig.path)
                _fit_inline_shape_to_width(shape, _available_figure_width_emu(doc))
            except Exception:
                img_p.add_run("[图片无法加载：%s]" % fig.path)
        else:
            img_p.add_run("[缺少图片：%s]" % fig.path)
    _emit_figure_addons(anchor, doc, fig)
    # 图题编号由模板图题样式自动生成。
    style = S.S_APPENDIX_FIGURE_CAPTION if appendix_letter else S.S_FIGURE_CAPTION
    cap = _new_paragraph_before(anchor, doc, style)
    _keep_previous_block_with_paragraph(cap)
    _emit_style_numbered_caption(cap, doc, style, "fig", fig.anchor_id, fig.caption or "")


def _emit_body_blocks(anchor: Paragraph, doc, blocks: List[object], meta: Optional[model.Meta] = None):
    blocks = _inject_chapter_boilerplate(blocks, meta=meta)
    in_terms = False
    in_normrefs = False
    ordered_list_num_id = None
    i = 0
    while i < len(blocks):
        blk = blocks[i]
        if isinstance(blk, model.Heading) and blk.level == 1:
            t = blk.text.strip()
            in_terms = "术语" in t and "定义" in t
            in_normrefs = "规范性引用文件" in t
        if isinstance(blk, model.Note):
            j = i
            notes = []
            while j < len(blocks) and isinstance(blocks[j], model.Note):
                notes.append(blocks[j])
                j += 1
            _emit_note_group(anchor, doc, notes)
            ordered_list_num_id = None
            i = j
            continue
        if isinstance(blk, model.ListBlock) and blk.ordered:
            style = S.ORDERED_LIST_STYLE_BY_LEVEL.get(blk.level, S.S_LIST_NUMBER_3)
            if blk.level == 1 or ordered_list_num_id is None:
                ordered_list_num_id = _new_numbering_instance_from_style(doc, style)
            _emit_body_block(
                anchor, doc, blk, in_terms=in_terms, in_normrefs=in_normrefs,
                list_num_id=ordered_list_num_id,
            )
        else:
            _emit_body_block(anchor, doc, blk, in_terms=in_terms, in_normrefs=in_normrefs)
            ordered_list_num_id = None
        i += 1


def _chapter_blocks(body: List[object], start: int) -> List[object]:
    out = []
    for blk in body[start + 1:]:
        if isinstance(blk, model.Heading) and blk.level == 1:
            break
        out.append(blk)
    return out


def _first_block_text(blocks: List[object]) -> str:
    for blk in blocks:
        text = getattr(blk, "text", "")
        if text:
            return text.strip()
    return ""


def _starts_with_any(text: str, prefixes: List[str]) -> bool:
    return any((text or "").startswith(prefix) for prefix in prefixes)


def _lead_paragraph(text: str) -> model.Paragraph:
    return model.Paragraph(spans=[model.Span(text)])


def _inject_chapter_boilerplate(body: List[object], meta: Optional[model.Meta] = None) -> List[object]:
    """为 规范性引用文件 / 术语和定义 / 符号和缩略语 空章补默认导语。"""
    out: List[object] = []
    for i, blk in enumerate(body):
        out.append(blk)
        if isinstance(blk, model.Heading) and blk.level == 1:
            title = blk.text.strip()
            content = _chapter_blocks(body, i)
            has_content = bool(content)
            first_text = _first_block_text(content)
            if not has_content:
                if title == S.CH_NORMATIVE_REF:
                    out.append(_lead_paragraph(bp.NORMATIVE_REF_NONE))
                elif title == S.CH_TERMS:
                    out.append(_lead_paragraph(bp.TERMS_NONE))
            elif title == S.CH_NORMATIVE_REF and not _starts_with_any(first_text, [
                "下列文件中的内容通过文中的规范性引用",
                bp.NORMATIVE_REF_NONE,
            ]):
                out.append(_lead_paragraph(bp.NORMATIVE_REF_LEAD))
            elif title == S.CH_TERMS and not _starts_with_any(first_text, [
                "下列术语和定义适用于本文件",
                "界定的术语和定义适用于本文件",
                "界定的以及下列术语和定义适用于本文件",
                bp.TERMS_NONE,
            ]):
                out.append(_lead_paragraph(bp.TERMS_LEAD))
            elif title == S.CH_SYMBOLS and not _starts_with_any(first_text, [
                "下列符号适用于本文件",
                "下列缩略语适用于本文件",
                "下列符号和缩略语适用于本文件",
            ]):
                lead = (meta.symbols_lead if meta is not None else "") or bp.SYMBOLS_LEAD
                out.append(_lead_paragraph(lead))
    return out


def _emit_appendices(anchor: Paragraph, doc, appendices: List[model.Appendix],
                     page_break_before: bool = True, section_before_each: bool = False,
                     start_index: int = 0):
    for ai, appx in enumerate(appendices):
        if section_before_each:
            _new_section_break_before(anchor, doc)
        letter = chr(ord("A") + start_index + ai)
        # 附录标题块：同一段三行，避免把附录标识、性质、标题拆成段落符。
        nature = bp.APPENDIX_NORMATIVE if appx.nature == "normative" else bp.APPENDIX_INFORMATIVE
        head = _new_paragraph_before(anchor, doc, S.S_APPENDIX_MARK)  # 自动"附录A"
        head.paragraph_format.page_break_before = page_break_before
        head.add_run().add_break()
        nature_run = head.add_run(nature)
        _apply_style_run_properties(doc, nature_run, S.S_APPENDIX_NATURE)
        head.add_run().add_break()
        _add_styled_runs(head, doc, S.S_APPENDIX_TITLE, appx.title_spans)
        _emit_hidden_appendix_caption_label(anchor, doc, S.S_APPENDIX_TABLE_LABEL)
        _emit_hidden_appendix_caption_label(anchor, doc, S.S_APPENDIX_FIGURE_LABEL)
        ordered_list_num_id = None
        i = 0
        while i < len(appx.blocks):
            blk = appx.blocks[i]
            if isinstance(blk, model.Heading):
                style = S.APPENDIX_CLAUSE_STYLE_BY_LEVEL.get(blk.level, S.S_PARA)
                _new_paragraph_before(anchor, doc, style, spans=blk.spans)
                ordered_list_num_id = None
            elif isinstance(blk, model.Note):
                j = i
                notes = []
                while j < len(appx.blocks) and isinstance(appx.blocks[j], model.Note):
                    notes.append(appx.blocks[j])
                    j += 1
                _emit_note_group(anchor, doc, notes)
                ordered_list_num_id = None
                i = j
                continue
            elif isinstance(blk, model.ListBlock) and blk.ordered:
                style = S.ORDERED_LIST_STYLE_BY_LEVEL.get(blk.level, S.S_LIST_NUMBER_3)
                if blk.level == 1 or ordered_list_num_id is None:
                    ordered_list_num_id = _new_numbering_instance_from_style(doc, style)
                _emit_body_block(
                    anchor, doc, blk,
                    appendix_letter=letter,
                    list_num_id=ordered_list_num_id,
                )
            else:
                _emit_body_block(anchor, doc, blk, appendix_letter=letter)
                ordered_list_num_id = None
            i += 1


def _emit_index_item(anchor: Paragraph, doc, item: model.IndexItem):
    para = _new_paragraph_before(anchor, doc, S.S_INDEX_ITEM)
    para.add_run(item.term)
    para.add_run("\t")
    para.add_run(item.targets)


def _remove_paragraph(para: Paragraph):
    parent = para._p.getparent()
    parent.remove(para._p)


def _emit_cover_toc(anchor: Paragraph, doc, kind: str):
    _emit_section_title_before(anchor, doc, S.S_TOC_TITLE, "目次", kind=kind)
    _make_toc_field(anchor, doc)


def _emit_cover_foreword(anchor: Paragraph, doc, meta: model.Meta, kind: str):
    _emit_section_title_before(anchor, doc, S.S_PREFACE_TITLE, "前言", kind=kind)
    _emit_foreword_content(anchor, doc, meta)


def _emit_cover_introduction(anchor: Paragraph, doc, meta: model.Meta, kind: str):
    if not meta.introduction.strip():
        return
    _emit_section_title_before(anchor, doc, S.S_PREFACE_TITLE, "引言", kind=kind)
    _emit_introduction_content(anchor, doc, meta)


def _emit_cover_references(anchor: Paragraph, doc, sdoc: model.StandardDoc, kind: str):
    if not sdoc.references:
        return
    _emit_section_title_before(anchor, doc, S.S_REF_TITLE, "参考文献", kind=kind)
    for item in sdoc.references:
        _new_paragraph_before(anchor, doc, S.S_REF_ITEM, text=item)


def _emit_cover_index(anchor: Paragraph, doc, sdoc: model.StandardDoc, kind: str):
    if not sdoc.index_groups:
        return
    _emit_section_title_before(anchor, doc, S.S_INDEX_TITLE, "索引", kind=kind)
    for group in sdoc.index_groups:
        _new_paragraph_before(anchor, doc, S.S_INDEX_LETTER, text=group.letter)
        for item in group.items:
            _emit_index_item(anchor, doc, item)


def _emit_document_end_line(anchor: Paragraph, doc, image_bytes: bytes):
    para = _new_paragraph_before(anchor, doc, S.S_PARA)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.add_run().add_picture(io.BytesIO(image_bytes))


def _emit_cover_sections(doc, sdoc: model.StandardDoc, end_line_image: bytes, kind: str):
    anchor = doc.add_paragraph()
    refs = _body_page_refs(doc)
    include_even = sdoc.meta.odd_even_pages
    body_started = False

    def front_break(start: Optional[int] = None):
        _new_section_break_before(
            anchor,
            doc,
            refs=refs,
            page_fmt="upperRoman",
            page_start=start,
            include_even=include_even,
        )

    def body_break(start: Optional[int] = None):
        nonlocal body_started
        _new_section_break_before(
            anchor,
            doc,
            refs=refs,
            page_start=start,
            include_even=include_even,
        )
        body_started = True

    _emit_cover_toc(anchor, doc, kind)
    front_break(start=1)
    _emit_cover_foreword(anchor, doc, sdoc.meta, kind)
    if sdoc.meta.introduction.strip():
        front_break()
        _emit_cover_introduction(anchor, doc, sdoc.meta, kind)
    front_break()
    _emit_body_standard_title(anchor, doc, sdoc.meta)
    _emit_important_notice(anchor, doc, sdoc.meta)
    _emit_body_blocks(anchor, doc, sdoc.body, meta=sdoc.meta)
    for idx, appx in enumerate(sdoc.appendices):
        body_break(start=1 if not body_started else None)
        _emit_appendices(
            anchor,
            doc,
            [appx],
            page_break_before=False,
            section_before_each=False,
            start_index=idx,
        )
    if sdoc.references:
        body_break(start=1 if not body_started else None)
        _emit_cover_references(anchor, doc, sdoc, kind)
    if sdoc.index_groups:
        body_break(start=1 if not body_started else None)
        _emit_cover_index(anchor, doc, sdoc, kind)
    _emit_document_end_line(anchor, doc, end_line_image)
    _configure_final_section(
        doc,
        refs=refs,
        page_start=1 if not body_started else None,
        include_even=include_even,
    )
    _remove_paragraph(anchor)

__all__ = [name for name in globals() if name.startswith("_")]
