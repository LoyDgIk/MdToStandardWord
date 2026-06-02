# -*- coding: utf-8 -*-
"""把 model.StandardDoc 构建进团体标准模板，输出标准文本 Word。

核心思路：复制模板 .docx，逐节回填。
- 封面字段：按样式名定位占位段落，替换 run 文本（保留格式）。
- 前言/引言/正文/参考文献/索引：每节由一个"承载 sectPr 的段落"作为终止段；
  清空该节内容段落（保留终止段以维持分节符与页眉页脚），再在终止段前插入新内容。
- 章条编号全部依赖模板的样式联动多级列表自动生成，标题只写文本、不写编号。
"""

from __future__ import annotations

import hashlib
import copy
import os
import re
import shutil
from typing import List, Optional

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import boilerplate as bp
from . import model
from . import styles as S


# --------------------------------------------------------------------------- #
# 低层 OXML 帮助
# --------------------------------------------------------------------------- #
def _carries_sectpr(p: Paragraph) -> bool:
    ppr = p._p.find(qn("w:pPr"))
    if ppr is None:
        return False
    return ppr.find(qn("w:sectPr")) is not None


def _set_numbering(para: Paragraph, num_id: int, ilvl: int):
    """在段落上显式设置列表编号（numId/ilvl），覆盖样式默认。

    num_id=0 表示关闭编号。用于让"无标题条"接入正文章条同一多级列表(numId=2)，
    从而与章/条编号连续（经 Word 验证：标准文件_X无标题 + numId=2 即可同步）。
    """
    p = para._p
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    old = ppr.find(qn("w:numPr"))
    if old is not None:
        ppr.remove(old)
    numpr = OxmlElement("w:numPr")
    ilvl_el = OxmlElement("w:ilvl")
    ilvl_el.set(qn("w:val"), str(ilvl))
    numid_el = OxmlElement("w:numId")
    numid_el.set(qn("w:val"), str(num_id))
    numpr.append(ilvl_el)
    numpr.append(numid_el)
    pstyle = ppr.find(qn("w:pStyle"))
    if pstyle is not None:
        pstyle.addnext(numpr)
    else:
        ppr.insert(0, numpr)



# --------------------------------------------------------------------------- #
# 域 / 书签（SEQ 自动编号、REF 交叉引用）
# ---------------------------------------------------------------------------
# SEQ 前缀与可见题注标签保持一致，避免 Word 只把数字识别为题注范围。
# 表/图/公式都使用可见 SEQ；模板标题样式只负责外观，不再负责编号。
# 书签：
#   _Ref... → Word 原生隐藏书签，分别围住编号、标签+编号、整项题注等范围。
# --------------------------------------------------------------------------- #

SEQ_TABLE    = "表"
SEQ_FIGURE   = "图"
SEQ_EQUATION = "公式"


def _bm_name(anchor_id: str) -> str:
    """把任意 id（可含中文）映射为合法的 Word 书签名。

    纯 ASCII 标识（如标准号 "GB 5749"）生成可读名 "Ref_GB5749"，便于在
    Word 原生交叉引用对话框里识别；含中文等则用哈希，保证有效且稳定。
    """
    norm = anchor_id.replace(" ", "").strip()
    ascii_safe = re.sub(r"[^0-9A-Za-z]+", "_", norm)
    if norm and re.match(r"^[0-9A-Za-z/._\-]+$", norm) and ascii_safe:
        return ("Ref_" + ascii_safe)[:40]
    return "Ref_" + hashlib.md5(norm.encode("utf-8")).hexdigest()[:16]


def _native_ref_name(ref_type: str, local_id: str, mode: str) -> str:
    """生成 Word 原生交叉引用常用的隐藏书签名 `_Ref#########`。"""
    key = "%s:%s:%s" % (ref_type, local_id, mode)
    n = int(hashlib.md5(key.encode("utf-8")).hexdigest()[:8], 16) % 900000000
    return "_Ref" + str(100000000 + n)


def _bookmark_start(para: Paragraph, name: str, bid: int):
    bs = OxmlElement("w:bookmarkStart")
    bs.set(qn("w:id"), str(bid))
    bs.set(qn("w:name"), name)
    para._p.append(bs)


def _bookmark_end(para: Paragraph, bid: int):
    be = OxmlElement("w:bookmarkEnd")
    be.set(qn("w:id"), str(bid))
    para._p.append(be)


def _add_ref_bookmark(para: Paragraph, bookmark_name: str, display_text: str = "?"):
    """插入 REF 域，引用指定书签。display_text 是域更新前的占位结果。"""
    r = para.add_run()
    fb = OxmlElement("w:fldChar"); fb.set(qn("w:fldCharType"), "begin"); r._r.append(fb)
    r = para.add_run()
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = " REF %s \\h " % bookmark_name
    r._r.append(it)
    r = para.add_run()
    fs = OxmlElement("w:fldChar"); fs.set(qn("w:fldCharType"), "separate"); r._r.append(fs)
    para.add_run(display_text)
    r = para.add_run()
    fe = OxmlElement("w:fldChar"); fe.set(qn("w:fldCharType"), "end"); r._r.append(fe)


def _hyperlink_run_text(text: str):
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    if text[:1].isspace() or text[-1:].isspace():
        t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _hyperlink_run_fld_char(kind: str):
    r = OxmlElement("w:r")
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), kind)
    r.append(fc)
    return r


def _hyperlink_run_instr(text: str):
    r = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = text
    r.append(it)
    return r


def _add_hyperlinked_ref_bookmark(para: Paragraph, hyperlink_anchor: str,
                                  ref_bookmark: str, prefix: str = "",
                                  suffix: str = "", display_text: str = "?"):
    """插入一个外层整体可点击、内部编号可更新的 REF。"""
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), hyperlink_anchor)
    link.set(qn("w:history"), "1")
    if prefix:
        link.append(_hyperlink_run_text(prefix))
    link.append(_hyperlink_run_fld_char("begin"))
    link.append(_hyperlink_run_instr(" REF %s \\h " % ref_bookmark))
    link.append(_hyperlink_run_fld_char("separate"))
    link.append(_hyperlink_run_text(display_text))
    link.append(_hyperlink_run_fld_char("end"))
    if suffix:
        link.append(_hyperlink_run_text(suffix))
    para._p.append(link)


def _add_ref(para: Paragraph, anchor_id: str, suffix: str = ""):
    """插入 REF 域，引用 Ref_{anchor_id}{suffix} 书签。"""
    _add_ref_bookmark(para, _bm_name(anchor_id) + suffix)


def _add_typed_ref(para: Paragraph, ref: model.RefSpan):
    if ref.ref_type == "std":
        _add_ref(para, ref.target)
        return
    if ref.ref_type == "eq" and ref.mode in ("label", "full"):
        _add_hyperlinked_ref_bookmark(
            para,
            _native_ref_name(ref.ref_type, ref.target, "label"),
            _native_ref_name(ref.ref_type, ref.target, "num"),
            prefix="式（",
            suffix="）",
            display_text="?",
        )
        return
    _add_ref_bookmark(para, _native_ref_name(ref.ref_type, ref.target, ref.mode))


def _make_run_hidden(run_elem):
    """给一个 <w:r> 加上 <w:vanish/> 使其成为隐藏文字。"""
    rpr = run_elem.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        run_elem.insert(0, rpr)
    rpr.append(OxmlElement("w:vanish"))


def _make_paragraph_hidden(para: Paragraph):
    """隐藏整个段落，保留其中字段供 Word 交叉引用对话框识别。"""
    ppr = para._p.get_or_add_pPr()
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ppr.append(rpr)
    if rpr.find(qn("w:vanish")) is None:
        rpr.append(OxmlElement("w:vanish"))


def _hide_run_if_needed(run, hidden: bool):
    if hidden:
        _make_run_hidden(run._r)
    return run


def _add_seq(para: Paragraph, seq_name: str, reset: bool = False, hidden: bool = False):
    """插入可见 SEQ 域。返回（SEQ 结果 run, 字段结束 run）。"""
    r = _hide_run_if_needed(para.add_run(), hidden)
    fb = OxmlElement("w:fldChar"); fb.set(qn("w:fldCharType"), "begin"); r._r.append(fb)
    r = _hide_run_if_needed(para.add_run(), hidden)
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    reset_part = " \\r 1" if reset else ""
    it.text = " SEQ %s \\* ARABIC%s " % (seq_name, reset_part)
    r._r.append(it)
    r = _hide_run_if_needed(para.add_run(), hidden)
    fs = OxmlElement("w:fldChar"); fs.set(qn("w:fldCharType"), "separate"); r._r.append(fs)
    result_run = _hide_run_if_needed(para.add_run("1"), hidden)
    end_run = _hide_run_if_needed(para.add_run(), hidden)
    fe = OxmlElement("w:fldChar"); fe.set(qn("w:fldCharType"), "end"); end_run._r.append(fe)
    return result_run, end_run


def _bookmark_start_before(run, name: str, bid: int):
    bs = OxmlElement("w:bookmarkStart")
    bs.set(qn("w:id"), str(bid))
    bs.set(qn("w:name"), name)
    run._r.addprevious(bs)


def _bookmark_end_after(run, bid: int):
    be = OxmlElement("w:bookmarkEnd")
    be.set(qn("w:id"), str(bid))
    run._r.addnext(be)


def _next_bm_id():
    bid = _COUNTER.bm
    _COUNTER.bm += 1
    return bid


def _add_hidden_seq(para: Paragraph, seq_name: str, num_prefix: str = "", separator: str = ""):
    """插入隐藏 SEQ 域（用于表/图等有样式自动编号的场景）。

    所有相关 run 带 <w:vanish/>，不可见但被 Word 域引擎识别，
    使条目出现在"表"/"图"交叉引用域中。

    num_prefix: 前缀文字（附录用，如 "A."），插在 SEQ 之前（隐藏）。
    separator:  SEQ 结果之后的分隔符（如 "　"），插在域结束之后（隐藏）。
    返回 SEQ 结果 run（隐藏）。
    """
    if num_prefix:
        r = para.add_run(num_prefix); _make_run_hidden(r._r)
    r = para.add_run(); _make_run_hidden(r._r)
    fb = OxmlElement("w:fldChar"); fb.set(qn("w:fldCharType"), "begin"); r._r.append(fb)
    r = para.add_run(); _make_run_hidden(r._r)
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = " SEQ %s \\* ARABIC " % seq_name
    r._r.append(it)
    r = para.add_run(); _make_run_hidden(r._r)
    fs = OxmlElement("w:fldChar"); fs.set(qn("w:fldCharType"), "separate"); r._r.append(fs)
    result_run = para.add_run("1"); _make_run_hidden(result_run._r)
    r = para.add_run(); _make_run_hidden(r._r)
    fe = OxmlElement("w:fldChar"); fe.set(qn("w:fldCharType"), "end"); r._r.append(fe)
    if separator:
        r = para.add_run(separator); _make_run_hidden(r._r)
    return result_run


def _emit_caption(para, seq_name, anchor_id, title, num_prefix="", hidden=False):
    """在段落上插入 SEQ + 双书签（:a 编号 / :b 编号+标题）。

    hidden=True 时用隐藏 SEQ（表/图——样式自带列表编号），
    hidden=False 时用可见 SEQ（公式——SEQ 是唯一编号来源）。

    num_prefix: 附录前缀（如 "A."），在 SEQ 结果前显示（隐藏）。
    """
    bm_full_id = None
    bm_num_id  = None
    if anchor_id:
        bm_full_id = _COUNTER.bm; _COUNTER.bm += 1
        bm_num_id  = _COUNTER.bm; _COUNTER.bm += 1

    # 书签 full 围住：prefix + SEQ + "　" + title → {:b} 编号+标题
    if bm_full_id is not None:
        _bookmark_start(para, _bm_name(anchor_id) + "_full", bm_full_id)

    # 前缀（附录用，如 "A."）+ SEQ 域 + 分隔符
    if hidden:
        result_run = _add_hidden_seq(para, seq_name, num_prefix=num_prefix, separator="　")
    else:
        if num_prefix:
            r = para.add_run(num_prefix); _make_run_hidden(r._r)
        result_run, seq_end_run = _add_seq(para, seq_name)

    # 书签 num 仅围住 SEQ 结果（含前缀）→ {:a} 纯编号
    # 关键：bookmarkStart 插到 result_run 之前，bookmarkEnd 插到字段结束之后，使文字落入书签
    if bm_num_id is not None:
        bm_start = OxmlElement("w:bookmarkStart")
        bm_start.set(qn("w:id"), str(bm_num_id))
        bm_start.set(qn("w:name"), _bm_name(anchor_id))
        result_run._r.addprevious(bm_start)
        bm_end = OxmlElement("w:bookmarkEnd")
        bm_end.set(qn("w:id"), str(bm_num_id))
        (seq_end_run if not hidden else result_run)._r.addnext(bm_end)

    # 标题文本（visible SEQ 时加显式分隔符；hidden 已在域后插入隐藏分隔符）
    if title:
        if not hidden:
            para.add_run("　")
        para.add_run(title)

    if bm_full_id is not None:
        _bookmark_end(para, bm_full_id)


# --------------------------------------------------------------------------- #
# 缩进辅助
# --------------------------------------------------------------------------- #
# 1/100 字符 -> twips 的近似换算（按五号 10.5pt，约 210 twips/字）
_CHAR_TWIPS = 2.1
_PT_TWIPS = 20  # 1 磅 = 20 twips


def _set_ind_pt(ind, left_pt=None, hanging_pt=None, first_line_pt=None):
    """用绝对磅值重设 w:ind，清掉字符级与冲突属性。"""
    for a in ("firstLine", "firstLineChars", "hanging", "hangingChars",
              "left", "leftChars", "start", "startChars", "end", "endChars"):
        k = qn("w:" + a)
        if ind.get(k) is not None:
            del ind.attrib[k]
    if left_pt is not None:
        ind.set(qn("w:left"), str(int(round(left_pt * _PT_TWIPS))))
    if hanging_pt is not None:
        ind.set(qn("w:hanging"), str(int(round(hanging_pt * _PT_TWIPS))))
    if first_line_pt is not None:
        ind.set(qn("w:firstLine"), str(int(round(first_line_pt * _PT_TWIPS))))


def _set_ind(ind, left_chars=None, hanging_chars=None, first_line_chars=None):
    """重设 w:ind 的字符级缩进，清掉冲突属性。"""
    for a in ("firstLine", "firstLineChars", "hanging", "hangingChars", "left", "leftChars"):
        k = qn("w:" + a)
        if ind.get(k) is not None:
            del ind.attrib[k]
    if left_chars is not None:
        ind.set(qn("w:leftChars"), str(left_chars))
        ind.set(qn("w:left"), str(int(left_chars * _CHAR_TWIPS)))
    if hanging_chars is not None:
        ind.set(qn("w:hangingChars"), str(hanging_chars))
        ind.set(qn("w:hanging"), str(int(hanging_chars * _CHAR_TWIPS)))
    if first_line_chars is not None:
        ind.set(qn("w:firstLineChars"), str(first_line_chars))
        ind.set(qn("w:firstLine"), str(int(first_line_chars * _CHAR_TWIPS)))


def _fix_style_indent(doc, style_name, left_pt, hanging_pt):
    """改写某段落样式的缩进，用绝对磅值对齐悬挂列项。"""
    try:
        st = doc.styles[style_name]
    except KeyError:
        return
    ppr = st.element.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    _set_ind_pt(ind, left_pt=left_pt, hanging_pt=hanging_pt)


def _set_runs(paragraph: Paragraph, spans: List[model.Span]):
    """按 spans 给段落添加 run（加粗/斜体）；RefSpan 转换为 REF 域。"""
    if not spans:
        return
    for sp in spans:
        if isinstance(sp, model.RefSpan):
            _add_typed_ref(paragraph, sp)
            continue
        for i, piece in enumerate(sp.text.split("\n")):
            if i > 0:
                paragraph.add_run().add_break()
            if piece:
                r = paragraph.add_run(piece)
                if sp.bold:
                    r.bold = True
                if sp.italic:
                    r.italic = True


def _apply_style_run_properties(doc, run, style_name: str):
    """把段落样式里的 rPr 复制到 run，用于同段软换行的附录标题局部格式。"""
    try:
        style = doc.styles[style_name]
    except KeyError:
        return
    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        return
    old = run._r.find(qn("w:rPr"))
    if old is not None:
        run._r.remove(old)
    run._r.insert(0, copy.deepcopy(rpr))


def _add_styled_runs(paragraph: Paragraph, doc, style_name: str, spans: List[model.Span]):
    for sp in spans:
        if isinstance(sp, model.RefSpan):
            _add_typed_ref(paragraph, sp)
            continue
        for i, piece in enumerate(sp.text.split("\n")):
            if i > 0:
                paragraph.add_run().add_break()
            if piece:
                run = paragraph.add_run(piece)
                _apply_style_run_properties(doc, run, style_name)
                if sp.bold:
                    run.bold = True
                if sp.italic:
                    run.italic = True


def _new_paragraph_before(anchor: Paragraph, doc, style_name: str,
                          text: str = "", spans: Optional[List[model.Span]] = None) -> Paragraph:
    """在 anchor 段落前插入一个套用 style_name 的新段落。"""
    new_p = OxmlElement("w:p")
    anchor._p.addprevious(new_p)
    para = Paragraph(new_p, anchor._parent)
    try:
        para.style = doc.styles[style_name]
    except KeyError:
        para.style = doc.styles[S.S_PARA]
    if spans is not None:
        _set_runs(para, spans)
    elif text:
        para.add_run(text)
    return para


def _set_field(doc, style_name: str, text: str) -> bool:
    """把第一个套用 style_name 的段落文本替换为 text，保留首个 run 的格式。返回是否命中。"""
    for p in doc.paragraphs:
        if p.style is not None and p.style.name == style_name:
            _replace_text_keep_format(p, text)
            return True
    return False


def _replace_text_keep_format(p: Paragraph, text: str):
    runs = p.runs
    if runs:
        runs[0].text = text
        for extra in runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        p.add_run(text)


def _find_para(doc, style_name: Optional[str] = None, text: Optional[str] = None,
               start: int = 0) -> Optional[int]:
    """在 doc.paragraphs 中查找匹配段落，返回索引。"""
    for i in range(start, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        if style_name is not None and (p.style is None or p.style.name != style_name):
            continue
        if text is not None and p.text.strip() != text:
            continue
        return i
    return None


def _next_sectpr_para(doc, start: int) -> Optional[int]:
    """从 start（含）起，找下一个承载 sectPr 的段落索引。"""
    for i in range(start, len(doc.paragraphs)):
        if _carries_sectpr(doc.paragraphs[i]):
            return i
    return None


def _remove_between(doc, start_p: Paragraph, end_p: Paragraph):
    """删除 body 中 start_p 与 end_p 之间的所有元素（不含两端）。

    会顺带删除区间内的 sdt（占位控件）等非段落元素。
    """
    body = start_p._p.getparent()
    start_el = start_p._p
    end_el = end_p._p
    el = start_el.getnext()
    while el is not None and el is not end_el:
        nxt = el.getnext()
        body.remove(el)
        el = nxt


# --------------------------------------------------------------------------- #
# TOC 域
# --------------------------------------------------------------------------- #
def _make_toc_field(anchor: Paragraph, doc):
    """在 anchor 前插入一个 TOC 域段落（目录 1-3 级，含页码）。

    使用 w:fldSimple，配合 settings 的 updateFields，Word 打开时自动生成目次。
    """
    para = _new_paragraph_before(anchor, doc, S.S_NORMAL)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), r' TOC \o "1-3" \h \z \u ')
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = '右键此处选择"更新域"以生成目次。'
    run.append(t)
    fld.append(run)
    para._p.append(fld)
    return para


def _enable_update_fields(doc):
    """让 Word 打开时提示更新域（用于 TOC 页码）。"""
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields")
        uf.set(qn("w:val"), "true")
        settings.append(uf)


# --------------------------------------------------------------------------- #
# 各节构建
# --------------------------------------------------------------------------- #
def _build_cover(doc, meta: model.Meta):
    # ICS / CCS 表（封面第一个表格的两行第二列）
    if doc.tables:
        t = doc.tables[0]
        try:
            if meta.ics:
                _set_cell_text(t.rows[0].cells[1], meta.ics)
            if meta.ccs:
                _set_cell_text(t.rows[1].cells[1], meta.ccs)
        except IndexError:
            pass

    if meta.standard_type:
        _set_field(doc, S.S_COVER_TYPE, meta.standard_type)
    if meta.number:
        _set_field(doc, S.S_COVER_NUMBER, meta.number)
    if meta.replaces:
        _set_field(doc, S.S_COVER_REPLACES, "代替 %s" % meta.replaces)
    if meta.title:
        _set_field(doc, S.S_COVER_NAME, meta.title)
    if meta.title_en:
        _set_field(doc, S.S_COVER_NAME_EN, meta.title_en)
    if meta.publish_date:
        _set_field(doc, S.S_COVER_PUBLISH, "%s发布" % meta.publish_date)
    if meta.implement_date:
        _set_field(doc, S.S_COVER_IMPLEMENT, "%s实施" % meta.implement_date)
    if meta.publisher:
        _set_field(doc, S.S_COVER_PUBLISHER, meta.publisher)


def _set_cell_text(cell, text):
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for extra in p.runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        p.add_run(text)


def _build_toc(doc):
    title_idx = _find_para(doc, style_name=S.S_TOC_TITLE, text="目次")
    if title_idx is None:
        return
    title_p = doc.paragraphs[title_idx]
    term_idx = _next_sectpr_para(doc, title_idx + 1)
    if term_idx is None:
        return
    term_p = doc.paragraphs[term_idx]
    _remove_between(doc, title_p, term_p)
    _make_toc_field(term_p, doc)


def _build_foreword(doc, meta: model.Meta):
    fw = meta.foreword
    idx = _find_para(doc, style_name=S.S_PREFACE_TITLE, text="前言")
    if idx is None:
        return
    title_p = doc.paragraphs[idx]
    term_idx = _next_sectpr_para(doc, idx + 1)
    if term_idx is None:
        return
    term_p = doc.paragraphs[term_idx]
    _remove_between(doc, title_p, term_p)

    def add(text):
        if text:
            _new_paragraph_before(term_p, doc, S.S_PARA, text=text)

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
            _new_paragraph_before(term_p, doc, S.S_LIST_DASH, text=ch)
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
        add(note)


def _build_introduction(doc, meta: model.Meta):
    idx = _find_para(doc, style_name=S.S_PREFACE_TITLE, text="引言")
    if idx is None:
        return
    title_p = doc.paragraphs[idx]
    term_idx = _next_sectpr_para(doc, idx + 1)
    if term_idx is None:
        return
    term_p = doc.paragraphs[term_idx]

    if meta.introduction.strip():
        _remove_between(doc, title_p, term_p)
        for line in meta.introduction.splitlines():
            if line.strip():
                _new_paragraph_before(term_p, doc, S.S_PARA, text=line.strip())
    else:
        # 无引言：删除标题 + 区间内容 + 本节终止段（含其 sectPr），整节合并入正文
        _remove_between(doc, title_p, term_p)
        body = title_p._p.getparent()
        body.remove(title_p._p)
        body.remove(term_p._p)


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


def _emit_term(anchor: Paragraph, doc, spans):
    """术语条目：编号(3.1)单独成行，下一行黑体中文术语 + 英文对应词。

    自动编号接入 numId=2 ilvl=2（术语为章 3 下的一级条）。
    """
    text = "".join(s.text for s in spans).strip()
    m = _TERM_SPLIT_RE.match(text)
    if m:
        cn, en = m.group(1).strip(), m.group(2).strip()
    else:
        cn, en = text, ""
    para = _new_paragraph_before(anchor, doc, S.S_TERM_1)
    _set_numbering(para, S.NUM_BODY, 2)
    para.add_run().add_break()          # 让自动编号 3.1 单独占一行
    r = para.add_run(cn)
    r.bold = True
    if en:
        para.add_run("　" + en)
    return para


def _emit_body_block(anchor: Paragraph, doc, blk, in_terms=False, in_normrefs=False,
                     appendix_letter=None):
    """把一个正文块插入到 anchor 之前。"""
    if isinstance(blk, model.Heading):
        if in_terms and blk.level == 2:
            _emit_term(anchor, doc, blk.spans)
            return
        style = S.HEADING_STYLE_BY_LEVEL.get(blk.level, S.S_PARA)
        _new_paragraph_before(anchor, doc, style, spans=blk.spans)
    elif isinstance(blk, model.Paragraph):
        if in_normrefs:
            _emit_normative_ref(anchor, doc, blk.spans)
        else:
            _new_paragraph_before(anchor, doc, S.S_PARA, spans=blk.spans)
    elif isinstance(blk, model.UntitledClause):
        # 无标题条：套用"X级无标题"样式 + 接入正文多级列表(numId=2)自动编号，
        # 不写字面编号；ilvl = 编号段数（"4.2.1" -> 3 段 -> ilvl=3 -> 渲染"4.2.1"）。
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
    elif isinstance(blk, model.Source):
        _emit_source(anchor, doc, blk)
    elif isinstance(blk, model.ListBlock):
        if blk.ordered:
            style = S.ORDERED_LIST_STYLE_BY_LEVEL.get(blk.level, S.S_LIST_NUMBER)
        else:
            style = S.S_LIST_DASH
        for it in blk.items:
            _new_paragraph_before(anchor, doc, style, spans=it.spans)
    elif isinstance(blk, model.TableModel):
        _emit_table(anchor, doc, blk, appendix_letter=appendix_letter)
    elif isinstance(blk, model.Figure):
        _emit_figure(anchor, doc, blk, appendix_letter=appendix_letter)
    elif isinstance(blk, model.Formula):
        fstyle = S.S_FORMULA_APPENDIX if appendix_letter else S.S_FORMULA
        _emit_formula(anchor, doc, blk, style=fstyle, appendix_letter=appendix_letter)


def _emit_formula(anchor: Paragraph, doc, formula, style=None, appendix_letter=None):
    """公式段：用制表位——[Tab]公式[Tab]（序号）。

    模板"标准文件_正文公式"已配好：居中制表位让公式居中，右制表位(带点引导)推序号。
    序号括在"（""）"内，括号为纯文本；书签仅围住不带括号的编号文字。
    SEQ 域使用中文前缀"公式"，使公式出现在 Word 交叉引用"公式"域。
    """
    from . import mathconv
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


# 表/图编号计数器（模块级在 build() 内重置）
class _Counter:
    def __init__(self):
        self.table = 0
        self.figure = 0
        self.bm = 1000
        self.seq_scope_counts = {}


_COUNTER = _Counter()


def _needs_seq_reset(ref_type: str, appendix_letter=None) -> bool:
    scope = appendix_letter or "body"
    key = (ref_type, scope)
    count = _COUNTER.seq_scope_counts.get(key, 0)
    _COUNTER.seq_scope_counts[key] = count + 1
    return count == 0


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


def _emit_table(anchor: Paragraph, doc, tbl: model.TableModel, appendix_letter=None):
    """表格：标题显式输出 `表N　标题`，模板样式只负责排版。"""
    style = S.S_APPENDIX_TABLE_CAPTION if appendix_letter else S.S_TABLE_CAPTION
    cap = _new_paragraph_before(anchor, doc, style)
    _set_numbering(cap, 0, 0)
    _emit_visible_caption(
        cap, "tbl", SEQ_TABLE, "表", tbl.anchor_id, tbl.caption or "",
        appendix_letter=appendix_letter,
    )

    # 表格本体
    ncols = len(tbl.header) if tbl.header else (len(tbl.rows[0]) if tbl.rows else 1)
    table = doc.add_table(rows=0, cols=ncols)
    try:
        table.style = doc.styles[S.S_TABLE_GRID]
    except KeyError:
        pass
    all_rows = ([tbl.header] if tbl.header else []) + tbl.rows
    for r_i, row in enumerate(all_rows):
        cells = table.add_row().cells
        for c_i in range(ncols):
            txt = row[c_i] if c_i < len(row) else ""
            cp = cells[c_i].paragraphs[0]
            try:
                cp.style = doc.styles[S.S_TABLE_CELL]
            except KeyError:
                pass
            cp.add_run(txt)
    # 把表格元素移动到 anchor 之前
    anchor._p.addprevious(table._tbl)


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


def _build_body(doc, sdoc: model.StandardDoc):
    """正文：定位正文节（引言/前言 sectPr 之后到 body-sectPr），清空样例后插入。"""
    ref_idx = _find_para(doc, style_name=S.S_REF_TITLE, text="参考文献")
    if ref_idx is None:
        return
    # 正文节终止段 = 参考文献标题前最近的承载 sectPr 的段落
    term_idx = None
    for i in range(ref_idx - 1, -1, -1):
        if _carries_sectpr(doc.paragraphs[i]):
            term_idx = i
            break
    if term_idx is None:
        return
    term_p = doc.paragraphs[term_idx]

    # 正文节起点：term 之前最近的另一个 sectPr 段（即引言/前言节终止段）
    start_idx = None
    for i in range(term_idx - 1, -1, -1):
        if _carries_sectpr(doc.paragraphs[i]):
            start_idx = i
            break
    start_p = doc.paragraphs[start_idx] if start_idx is not None else None

    # 清除正文节内既有样例（起点之后到终止段之间）
    if start_p is not None:
        _remove_between(doc, start_p, term_p)
    else:
        # 没有上游 sectPr，则清除文档开头到 term 之间——不应发生，保守跳过
        pass

    # 注入正文块；对特殊空章自动补导语
    blocks = _inject_chapter_boilerplate(sdoc.body)
    in_terms = False
    in_normrefs = False
    for blk in blocks:
        if isinstance(blk, model.Heading) and blk.level == 1:
            t = blk.text.strip()
            in_terms = "术语" in t and "定义" in t
            in_normrefs = "规范性引用文件" in t
        _emit_body_block(term_p, doc, blk, in_terms=in_terms, in_normrefs=in_normrefs)


def _inject_chapter_boilerplate(body: List[object]) -> List[object]:
    """为 规范性引用文件 / 术语和定义 / 符号和缩略语 空章补默认导语。"""
    out: List[object] = []
    n = len(body)
    for i, blk in enumerate(body):
        out.append(blk)
        if isinstance(blk, model.Heading) and blk.level == 1:
            title = blk.text.strip()
            # 判断该章后是否紧跟另一个一级章（即本章无正文内容）
            has_content = False
            for j in range(i + 1, n):
                nb = body[j]
                if isinstance(nb, model.Heading) and nb.level == 1:
                    break
                has_content = True
                break
            if not has_content:
                if title == S.CH_NORMATIVE_REF:
                    out.append(model.Paragraph(spans=[model.Span(bp.NORMATIVE_REF_NONE)]))
                elif title == S.CH_TERMS:
                    out.append(model.Paragraph(spans=[model.Span(bp.TERMS_NONE)]))
    return out


def _build_appendices(doc, sdoc: model.StandardDoc):
    if not sdoc.appendices:
        return
    ref_idx = _find_para(doc, style_name=S.S_REF_TITLE, text="参考文献")
    if ref_idx is None:
        return
    # 附录应位于正文节内、参考文献分节符之前：取参考文献标题前最近的承载 sectPr 的段落
    term_idx = None
    for i in range(ref_idx - 1, -1, -1):
        if _carries_sectpr(doc.paragraphs[i]):
            term_idx = i
            break
    anchor = doc.paragraphs[term_idx] if term_idx is not None else doc.paragraphs[ref_idx]
    for ai, appx in enumerate(sdoc.appendices):
        letter = chr(ord("A") + ai)
        # 附录标题块：同一段三行，避免把附录标识、性质、标题拆成段落符。
        nature = bp.APPENDIX_NORMATIVE if appx.nature == "normative" else bp.APPENDIX_INFORMATIVE
        head = _new_paragraph_before(anchor, doc, S.S_APPENDIX_MARK)  # 自动"附录A"
        head.paragraph_format.page_break_before = True
        head.add_run().add_break()
        nature_run = head.add_run(nature)
        _apply_style_run_properties(doc, nature_run, S.S_APPENDIX_NATURE)
        head.add_run().add_break()
        _add_styled_runs(head, doc, S.S_APPENDIX_TITLE, appx.title_spans)
        for blk in appx.blocks:
            if isinstance(blk, model.Heading):
                style = S.APPENDIX_CLAUSE_STYLE_BY_LEVEL.get(blk.level, S.S_PARA)
                _new_paragraph_before(anchor, doc, style, spans=blk.spans)
            else:
                _emit_body_block(anchor, doc, blk, appendix_letter=letter)


def _move_index_end_line_before_section_break(title_p: Paragraph):
    """把模板索引节末尾的居中线图移到上一节末尾，避免单独占尾页。"""
    body = title_p._p.getparent()
    section_break = title_p._p.getprevious()
    while section_break is not None:
        if section_break.tag == qn("w:p"):
            ppr = section_break.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                break
        section_break = section_break.getprevious()
    if section_break is None:
        return

    el = title_p._p.getnext()
    while el is not None:
        nxt = el.getnext()
        if el.tag == qn("w:p"):
            ppr = el.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                break
            txt = "".join(t.text or "" for t in el.findall(".//" + qn("w:t")))
            has_draw = el.findall(".//" + qn("w:drawing"))
            if has_draw and not txt.strip():
                section_break.addprevious(el)
        el = nxt
    body.remove(section_break)


def _find_index_end_line(title_p: Paragraph):
    """查找索引节末尾的标准文档结束线段。"""
    el = title_p._p.getnext()
    while el is not None:
        if el.tag == qn("w:p"):
            txt = "".join(t.text or "" for t in el.findall(".//" + qn("w:t")))
            has_draw = el.findall(".//" + qn("w:drawing"))
            if has_draw and not txt.strip():
                return el
        el = el.getnext()
    return None


def _clear_index_placeholders(title_p: Paragraph, end_line_el):
    """删除索引标题与结束线之间的模板占位段。"""
    body = title_p._p.getparent()
    el = title_p._p.getnext()
    while el is not None and el is not end_line_el:
        nxt = el.getnext()
        if el.tag == qn("w:p"):
            body.remove(el)
        el = nxt


def _emit_index_item(anchor: Paragraph, doc, item: model.IndexItem):
    para = _new_paragraph_before(anchor, doc, S.S_INDEX_ITEM)
    para.add_run(item.term)
    para.add_run("\t")
    para.add_run(item.targets)


def _build_references(doc, sdoc: model.StandardDoc):
    idx = _find_para(doc, style_name=S.S_REF_TITLE, text="参考文献")
    if idx is None:
        return
    title_p = doc.paragraphs[idx]
    term_idx = _next_sectpr_para(doc, idx + 1)
    if term_idx is None:
        return
    term_p = doc.paragraphs[term_idx]
    _remove_between(doc, title_p, term_p)
    if not sdoc.references:
        body = title_p._p.getparent()
        prev = title_p._p.getprevious()
        if prev is not None and prev.tag == qn("w:p"):
            ppr = prev.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                body.remove(prev)
        body.remove(title_p._p)
        return
    for item in sdoc.references:
        _new_paragraph_before(term_p, doc, S.S_REF_ITEM, text=item)


def _build_index(doc, sdoc: model.StandardDoc):
    """生成或删除索引节。"""
    idx = _find_para(doc, style_name=S.S_INDEX_TITLE, text="索引")
    if idx is None:
        return
    title_p = doc.paragraphs[idx]
    if sdoc.index_groups:
        end_line_el = _find_index_end_line(title_p)
        if end_line_el is None:
            return
        _clear_index_placeholders(title_p, end_line_el)
        anchor = Paragraph(end_line_el, title_p._parent)
        for group in sdoc.index_groups:
            _new_paragraph_before(anchor, doc, S.S_INDEX_LETTER, text=group.letter)
            for item in group.items:
                _emit_index_item(anchor, doc, item)
        return

    _move_index_end_line_before_section_break(title_p)
    body = title_p._p.getparent()
    # 删除标题之后、直到下一个承载 sectPr 的段或文末的空段
    el = title_p._p.getnext()
    while el is not None:
        nxt = el.getnext()
        if el.tag == qn("w:p"):
            ppr = el.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                break
            txt = "".join(t.text or "" for t in el.findall(".//" + qn("w:t")))
            has_draw = el.findall(".//" + qn("w:drawing"))
            if txt.strip() or has_draw:
                # 保留有图/有文字的段（如背景图）
                el = nxt
                continue
            body.remove(el)
        el = nxt
    if el is not None and el.tag == qn("w:p"):
        ppr = el.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            body.remove(el)
    body.remove(title_p._p)


# --------------------------------------------------------------------------- #
# 入口
# --------------------------------------------------------------------------- #
def build(sdoc: model.StandardDoc, template_path: str, output_path: str):
    _COUNTER.table = 0
    _COUNTER.figure = 0
    _COUNTER.bm = 1000
    _COUNTER.seq_scope_counts = {}

    # 复制模板再打开，避免改动原模板
    tmp = output_path
    shutil.copyfile(template_path, tmp)
    doc = Document(tmp)

    # 收紧列项/参考文献缩进（绝对磅值；正文段首行缩进=21pt≈2字）：
    #   破折号列项：破折号置于 2 字处(21pt)，文字悬挂到 42pt。
    #   破折号列项（二级）：再缩进一层。
    #   参考文献条目：编号[N]顶格，文字悬挂到 21pt。
    _fix_style_indent(doc, S.S_LIST_DASH, left_pt=42, hanging_pt=21)
    _fix_style_indent(doc, "标准文件_破折号列项（二级）", left_pt=63, hanging_pt=21)
    _fix_style_indent(doc, S.S_REF_ITEM, left_pt=21, hanging_pt=21)

    _build_cover(doc, sdoc.meta)
    _build_foreword(doc, sdoc.meta)
    _build_introduction(doc, sdoc.meta)
    _build_body(doc, sdoc)
    _build_appendices(doc, sdoc)
    _build_references(doc, sdoc)
    _build_index(doc, sdoc)
    _build_toc(doc)
    _enable_update_fields(doc)

    doc.save(output_path)
    return output_path
