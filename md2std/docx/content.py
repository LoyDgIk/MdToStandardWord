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
    for note in term.notes:
        style = S.S_NOTE_X if note.index else S.S_NOTE
        _new_paragraph_before(anchor, doc, style, spans=note.spans)
    if term.source is not None:
        _emit_source(anchor, doc, term.source)


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
        style = S.S_NOTE_X if blk.index else S.S_NOTE
        _new_paragraph_before(anchor, doc, style, spans=blk.spans)
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


def _emit_visible_caption(para: Paragraph, ref_type: str, seq_name: str,
                          label_text: str, anchor_id: str, title: str,
                          appendix_letter=None):
    """输出 `表1　标题` / `图A.1　标题`，并建立 num/label/full 书签。"""
    label_run = para.add_run(label_text)
    bm_full_id = bm_label_id = bm_num_id = bm_text_id = None
    if anchor_id:
        bm_full_id = _next_bm_id()
        bm_label_id = _next_bm_id()
        bm_num_id = _next_bm_id()
        bm_text_id = _next_bm_id() if title else None
        _bookmark_start_before(label_run, _native_ref_name(ref_type, anchor_id, "full"), bm_full_id)
        _bookmark_start_before(label_run, _native_ref_name(ref_type, anchor_id, "label"), bm_label_id)

    prefix_run = None
    if appendix_letter:
        prefix_run = para.add_run(appendix_letter + ".")
    result_run, seq_end_run = _add_seq(para, seq_name, reset=_needs_seq_reset(ref_type, appendix_letter))

    if anchor_id:
        _bookmark_start_before(prefix_run or result_run, _native_ref_name(ref_type, anchor_id, "num"), bm_num_id)

    title_run = None
    if title:
        para.add_run("　")
        title_run = para.add_run(title)
        if anchor_id and bm_text_id is not None:
            _bookmark_start_before(title_run, _native_ref_name(ref_type, anchor_id, "text"), bm_text_id)

    if anchor_id:
        if title_run is not None:
            _add_bookmark_ends_after(seq_end_run, [bm_num_id, bm_label_id])
            if bm_text_id is not None:
                _bookmark_end_after(title_run, bm_text_id)
            _bookmark_end_after(title_run, bm_full_id)
        else:
            _add_bookmark_ends_after(seq_end_run, [bm_num_id, bm_label_id, bm_full_id])


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


def _emit_figure_table_note(anchor: Paragraph, doc, note: model.Note):
    para = _new_paragraph_before(anchor, doc, S.S_FIG_TABLE_NOTE)
    para.add_run("注%d：" % note.index if note.index else "注：")
    _add_character_styled_runs(para, doc, S.S_FIG_TABLE_NOTE_CONTENT, note.spans)


def _emit_figure_table_source(anchor: Paragraph, doc, source: model.FigureTableSource):
    para = _new_paragraph_before(anchor, doc, S.S_FIG_TABLE_SOURCE)
    para.add_run("来源：")
    _set_runs(para, source.spans)


def _emit_figure_table_addons(anchor: Paragraph, doc, obj):
    for note in getattr(obj, "notes", []) or []:
        _emit_figure_table_note(anchor, doc, note)
    source = getattr(obj, "source", None)
    if source is not None:
        _emit_figure_table_source(anchor, doc, source)


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


def _set_row_repeat_header(row):
    trpr = row._tr.get_or_add_trPr()
    if trpr.find(qn("w:tblHeader")) is None:
        hdr = OxmlElement("w:tblHeader")
        hdr.set(qn("w:val"), "true")
        trpr.append(hdr)


def _cell_is_long_text(text: str, colspan: int = 1) -> bool:
    text = (text or "").strip()
    return colspan > 1 or len(text) > 25 or "。" in text or "；" in text or "：" in text


def _parts_from_text(text: str) -> List[model.TableCellPart]:
    return [model.TableCellPart("text", text or "")]


def _emit_table_cell_parts(paragraph: Paragraph, parts: List[model.TableCellPart]):
    from .. import mathconv
    if not parts:
        return
    for part in parts:
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
                paragraph.add_run(piece)


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
        )
    else:
        cap.add_run("表")
    title = (tbl.caption or "").strip()
    if title:
        cap.add_run("　")
        cap.add_run(title)
    cap.add_run("（续）")


def _emit_table_part(anchor: Paragraph, doc, tbl: model.TableModel, rows, row_parts, row_colspans,
                     appendix_letter=None):
    """输出一个表格片段。续表复用原表头，不新增 SEQ。"""

    def row_width(row, spans):
        if spans:
            return sum(spans[:len(row)])
        return len(row)

    widths = []
    if tbl.header:
        widths.append(row_width(tbl.header, tbl.header_colspans))
    widths.extend(row_width(row, row_colspans[i] if i < len(row_colspans) else [])
                  for i, row in enumerate(rows))
    ncols = max(widths) if widths else 1
    table = doc.add_table(rows=0, cols=ncols)
    try:
        table.style = doc.styles[S.S_TABLE_GRID]
    except KeyError:
        pass

    all_rows = ([tbl.header] if tbl.header else []) + rows
    header_parts = tbl.header_parts or [_parts_from_text(text) for text in tbl.header]
    all_parts = ([header_parts] if tbl.header else []) + row_parts
    all_spans = ([tbl.header_colspans or [1 for _ in tbl.header]] if tbl.header else [])
    all_spans += [
        row_colspans[i] if i < len(row_colspans) and row_colspans[i]
        else [1 for _ in row]
        for i, row in enumerate(rows)
    ]

    for r_i, row in enumerate(all_rows):
        word_row = table.add_row()
        if r_i == 0 and tbl.header:
            _set_row_repeat_header(word_row)
        cells = word_row.cells
        spans = all_spans[r_i] if r_i < len(all_spans) else [1 for _ in row]
        parts_row = all_parts[r_i] if r_i < len(all_parts) else []
        c_pos = 0
        for c_i, txt in enumerate(row):
            if c_pos >= ncols:
                break
            colspan = spans[c_i] if c_i < len(spans) else 1
            colspan = max(1, min(colspan, ncols - c_pos))
            cell = cells[c_pos]
            if colspan > 1:
                cell = cell.merge(cells[c_pos + colspan - 1])
            _set_cell_vertical_center(cell)
            cp = cell.paragraphs[0]
            try:
                cp.style = doc.styles[S.S_TABLE_CELL]
            except KeyError:
                pass
            cp.alignment = (
                WD_ALIGN_PARAGRAPH.LEFT
                if _cell_is_long_text(txt, colspan)
                else WD_ALIGN_PARAGRAPH.CENTER
            )
            parts = parts_row[c_i] if c_i < len(parts_row) else _parts_from_text(txt)
            _emit_table_cell_parts(cp, parts)
            c_pos += colspan
    anchor._p.addprevious(table._tbl)


def _emit_table(anchor: Paragraph, doc, tbl: model.TableModel, appendix_letter=None):
    """表格：标题显式输出 `表N　标题`，模板样式只负责排版。"""
    style = S.S_APPENDIX_TABLE_CAPTION if appendix_letter else S.S_TABLE_CAPTION
    cap = _new_paragraph_before(anchor, doc, style)
    _set_numbering(cap, 0, 0)
    _emit_visible_caption(
        cap, "tbl", SEQ_TABLE, "表", tbl.anchor_id, tbl.caption or "",
        appendix_letter=appendix_letter,
    )

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
            )
            first = False
            start = end
        _emit_figure_table_addons(anchor, doc, tbl)
        return

    _emit_table_part(anchor, doc, tbl, tbl.rows, row_parts, row_colspans, appendix_letter=appendix_letter)
    _emit_figure_table_addons(anchor, doc, tbl)


def _emit_figure(anchor: Paragraph, doc, fig: model.Figure, appendix_letter=None):
    img_p = _new_paragraph_before(anchor, doc, S.S_NORMAL)
    if fig.path and os.path.isfile(fig.path):
        try:
            img_p.add_run().add_picture(fig.path)
        except Exception:
            img_p.add_run("[图片无法加载：%s]" % fig.path)
    else:
        img_p.add_run("[缺少图片：%s]" % fig.path)
    # 图题：标题显式输出 `图N　标题`，模板样式只负责排版。
    style = S.S_APPENDIX_FIGURE_CAPTION if appendix_letter else S.S_FIGURE_CAPTION
    cap = _new_paragraph_before(anchor, doc, style)
    _set_numbering(cap, 0, 0)
    _emit_visible_caption(
        cap, "fig", SEQ_FIGURE, "图", fig.anchor_id, fig.caption or "",
        appendix_letter=appendix_letter,
    )
    _emit_figure_table_addons(anchor, doc, fig)


def _emit_body_blocks(anchor: Paragraph, doc, blocks: List[object], meta: Optional[model.Meta] = None):
    blocks = _inject_chapter_boilerplate(blocks, meta=meta)
    in_terms = False
    in_normrefs = False
    ordered_list_num_id = None
    for blk in blocks:
        if isinstance(blk, model.Heading) and blk.level == 1:
            t = blk.text.strip()
            in_terms = "术语" in t and "定义" in t
            in_normrefs = "规范性引用文件" in t
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
        ordered_list_num_id = None
        for blk in appx.blocks:
            if isinstance(blk, model.Heading):
                style = S.APPENDIX_CLAUSE_STYLE_BY_LEVEL.get(blk.level, S.S_PARA)
                _new_paragraph_before(anchor, doc, style, spans=blk.spans)
                ordered_list_num_id = None
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
