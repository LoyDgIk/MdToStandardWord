# -*- coding: utf-8 -*-
"""Low-level WordprocessingML, numbering, section, field, and style helpers."""

from __future__ import annotations

import copy
import hashlib
import re
from typing import List, Optional

from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .. import model
from .. import styles as S
from .state import _COUNTER

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


def _style_numbering(doc, style_name: str):
    """读取模板样式自带的 numPr，返回 (numId, ilvl)。"""
    try:
        style = doc.styles[style_name]
    except KeyError:
        return None
    ppr = style.element.find(qn("w:pPr"))
    if ppr is None:
        return None
    numpr = ppr.find(qn("w:numPr"))
    if numpr is None:
        return None
    numid = numpr.find(qn("w:numId"))
    if numid is None or numid.get(qn("w:val")) is None:
        return None
    ilvl = numpr.find(qn("w:ilvl"))
    return int(numid.get(qn("w:val"))), int(ilvl.get(qn("w:val")) if ilvl is not None and ilvl.get(qn("w:val")) is not None else 0)


def _abstract_num_id_for_num(doc, num_id: int) -> Optional[int]:
    """返回 numId 对应的 abstractNumId。"""
    numbering = doc.part.numbering_part.element
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) != str(num_id):
            continue
        abstract = num.find(qn("w:abstractNumId"))
        if abstract is None or abstract.get(qn("w:val")) is None:
            return None
        return int(abstract.get(qn("w:val")))
    return None


def _next_numbering_id(doc) -> int:
    numbering = doc.part.numbering_part.element
    ids = []
    for num in numbering.findall(qn("w:num")):
        val = num.get(qn("w:numId"))
        if val is not None and val.isdigit():
            ids.append(int(val))
    return (max(ids) if ids else 0) + 1


def _new_numbering_instance_from_style(doc, style_name: str) -> Optional[int]:
    """基于模板样式的 abstractNum 新建一个 numId，使独立列表重新编号。"""
    style_numbering = _style_numbering(doc, style_name)
    if style_numbering is None:
        return None
    base_num_id, _ = style_numbering
    abstract_num_id = _abstract_num_id_for_num(doc, base_num_id)
    if abstract_num_id is None:
        return None

    new_num_id = _next_numbering_id(doc)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(new_num_id))
    abstract = OxmlElement("w:abstractNumId")
    abstract.set(qn("w:val"), str(abstract_num_id))
    num.append(abstract)
    for ilvl in range(9):
        lvl_override = OxmlElement("w:lvlOverride")
        lvl_override.set(qn("w:ilvl"), str(ilvl))
        start_override = OxmlElement("w:startOverride")
        start_override.set(qn("w:val"), "1")
        lvl_override.append(start_override)
        num.append(lvl_override)
    doc.part.numbering_part.element.append(num)
    return new_num_id


def _set_numbering_from_style(para: Paragraph, doc, style_name: str,
                              num_id_override: Optional[int] = None):
    """显式套用模板样式的编号定义，避免 Word 忽略样式级列表缩进。"""
    numbering = _style_numbering(doc, style_name)
    if numbering is None:
        return
    num_id, ilvl = numbering
    if num_id_override is not None:
        num_id = num_id_override
    _set_numbering(para, num_id, ilvl)


def _first_existing_path(candidates: List[str], label: str) -> str:
    for path in candidates:
        if path:
            import os
            if os.path.isfile(path):
                return path
    raise FileNotFoundError("找不到%s：%s" % (label, "；".join(candidates)))


def _configure_standard_styles(doc):
    # 模板自带的破折号列项/参考文献缩进偏大。这里仍使用模板样式和编号，
    # 只把编号级别缩进收敛到“自然段首行两字”附近。
    _clear_style_indent(doc, S.S_LIST_DASH)
    _clear_style_indent(doc, "标准文件_破折号列项（二级）")
    _fix_numbering_style_indent(doc, S.S_LIST_DASH, left_twips=600, hanging_twips=200)
    _fix_numbering_style_indent(doc, "标准文件_破折号列项（二级）", left_twips=920, hanging_twips=200)
    _fix_numbering_style_indent(doc, S.S_REF_ITEM, left_twips=620, hanging_twips=420)



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


def _add_ref_bookmark(para: Paragraph, bookmark_name: str, display_text: str = "?",
                      charformat: bool = False):
    """插入 REF 域，引用指定书签。display_text 是域更新前的占位结果。"""
    r = para.add_run()
    fb = OxmlElement("w:fldChar"); fb.set(qn("w:fldCharType"), "begin"); r._r.append(fb)
    r = para.add_run()
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    fmt = " \\* CHARFORMAT" if charformat else ""
    it.text = " REF %s \\h%s " % (bookmark_name, fmt)
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
                                  suffix: str = "", display_text: str = "?",
                                  charformat: bool = False):
    """插入一个外层整体可点击、内部编号可更新的 REF。"""
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), hyperlink_anchor)
    link.set(qn("w:history"), "1")
    if prefix:
        link.append(_hyperlink_run_text(prefix))
    link.append(_hyperlink_run_fld_char("begin"))
    fmt = " \\* CHARFORMAT" if charformat else ""
    link.append(_hyperlink_run_instr(" REF %s \\h%s " % (ref_bookmark, fmt)))
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
            charformat=True,
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


def _set_ind_twips(ind, left_twips=None, hanging_twips=None, first_line_twips=None):
    """用 twips 重设 w:ind，供模板 numbering 级别微调用。"""
    for a in ("firstLine", "firstLineChars", "hanging", "hangingChars",
              "left", "leftChars", "start", "startChars", "end", "endChars"):
        k = qn("w:" + a)
        if ind.get(k) is not None:
            del ind.attrib[k]
    if left_twips is not None:
        ind.set(qn("w:left"), str(left_twips))
    if hanging_twips is not None:
        ind.set(qn("w:hanging"), str(hanging_twips))
    if first_line_twips is not None:
        ind.set(qn("w:firstLine"), str(first_line_twips))


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


def _clear_style_indent(doc, style_name):
    """移除段落样式自身缩进，避免与 numbering 缩进叠加。"""
    try:
        st = doc.styles[style_name]
    except KeyError:
        return
    ppr = st.element.find(qn("w:pPr"))
    if ppr is None:
        return
    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        ppr.remove(ind)


def _fix_numbering_style_indent(doc, style_name, left_twips, hanging_twips):
    """按样式关联的 numbering level 修正列表/参考文献缩进。"""
    try:
        style = doc.styles[style_name]
    except KeyError:
        return
    style_id = style.element.get(qn("w:styleId"))
    if not style_id:
        return
    try:
        numbering = doc.part.numbering_part.element
    except Exception:
        return
    for lvl in numbering.findall(".//" + qn("w:lvl")):
        pstyle = lvl.find(qn("w:pStyle"))
        if pstyle is None or pstyle.get(qn("w:val")) != style_id:
            continue
        ppr = lvl.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            lvl.append(ppr)
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            ppr.append(ind)
        _set_ind_twips(ind, left_twips=left_twips, hanging_twips=hanging_twips)


def _set_runs(paragraph: Paragraph, spans: List[model.Span]):
    """按 spans 给段落添加 run（加粗/斜体）；RefSpan 转换为 REF 域。"""
    from .. import mathconv
    if not spans:
        return
    for sp in spans:
        if isinstance(sp, model.RefSpan):
            _add_typed_ref(paragraph, sp)
            continue
        if isinstance(sp, model.FormulaSpan):
            omath = mathconv.latex_to_omml(sp.text)
            if omath is not None:
                paragraph._p.append(omath)
            else:
                paragraph.add_run(sp.text)
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
                if getattr(sp, "subscript", False):
                    r.font.subscript = True
                if getattr(sp, "superscript", False):
                    r.font.superscript = True


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
                _apply_style_run_properties(doc, run, style_name)
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
                if getattr(sp, "subscript", False):
                    run.font.subscript = True
                if getattr(sp, "superscript", False):
                    run.font.superscript = True


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


def _new_numbered_style_paragraph(anchor: Paragraph, doc, style_name: str,
                                  text: str = "", spans: Optional[List[model.Span]] = None,
                                  num_id_override: Optional[int] = None) -> Paragraph:
    para = _new_paragraph_before(anchor, doc, style_name, text=text, spans=spans)
    _set_numbering_from_style(para, doc, style_name, num_id_override=num_id_override)
    return para


def _set_direct_paragraph_spacing(para: Paragraph, before: Optional[int] = None,
                                  after: Optional[int] = None):
    """设置模板标题段落保留的直接段距。"""
    if before is None and after is None:
        return
    ppr = para._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    if before is not None:
        spacing.set(qn("w:before"), str(before))
    if after is not None:
        spacing.set(qn("w:after"), str(after))


def _set_run_char_spacing(run, value: int, east_asia_hint: bool = False):
    rpr = run._r.get_or_add_rPr()
    if east_asia_hint:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:hint"), "eastAsia")
    spacing = rpr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        rpr.append(spacing)
    spacing.set(qn("w:val"), str(value))


def _set_run_east_asia_hint(run):
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:hint"), "eastAsia")


def _section_title_after(title: str, kind: str) -> Optional[int]:
    front_titles = {"目次", "前言", "引言"}
    body_titles = {"参考文献", "索引"}
    if kind == "national":
        if title in front_titles:
            return 468
        if title in body_titles:
            return 156
    else:
        if title in front_titles:
            return 360
        if title in body_titles:
            return 120
    return None


def _section_title_before(title: str, kind: str) -> Optional[int]:
    # 国家标准模板的前言首页标题带直接段前距，其他标题不额外设置。
    if kind == "national" and title == "前言":
        return 900
    return None


def _section_title_char_spacing(title: str) -> tuple[Optional[int], bool]:
    if title in {"目次", "前言", "引言"}:
        return 320, False
    if title == "参考文献":
        return 105, True
    if title == "索引":
        return 210, True
    return None, False


def _emit_section_title_before(anchor: Paragraph, doc, style_name: str, title: str,
                               kind: str = "group") -> Paragraph:
    """输出与模板标题段落一致的节标题。

    模板中的"目次/前言/引言/参考文献/索引"不是单个普通 run：
    标题前 N-1 个字带字符间距，最后一个字不带字符间距；段距也有直接格式。
    """
    para = _new_paragraph_before(anchor, doc, style_name)
    _set_direct_paragraph_spacing(
        para,
        before=_section_title_before(title, kind),
        after=_section_title_after(title, kind),
    )
    char_spacing, east_asia_hint = _section_title_char_spacing(title)
    if char_spacing is None or len(title) < 2:
        para.add_run(title)
        return para

    first = para.add_run(title[:-1])
    _set_run_char_spacing(first, char_spacing, east_asia_hint=east_asia_hint)
    last = para.add_run(title[-1])
    if east_asia_hint:
        _set_run_east_asia_hint(last)
    return para


def _emit_page_break(anchor: Paragraph, doc) -> Paragraph:
    para = _new_paragraph_before(anchor, doc, S.S_NORMAL)
    para.add_run().add_break(WD_BREAK.PAGE)
    return para


def _clear_section_header_footer_refs(sectpr):
    for tag in ("w:headerReference", "w:footerReference"):
        for el in list(sectpr.findall(qn(tag))):
            sectpr.remove(el)


def _clear_section_page_number(sectpr):
    old = sectpr.find(qn("w:pgNumType"))
    if old is not None:
        sectpr.remove(old)


def _set_section_page_number(sectpr, fmt: Optional[str] = None, start: Optional[int] = None):
    _clear_section_page_number(sectpr)
    if fmt is None and start is None:
        return
    pg = OxmlElement("w:pgNumType")
    if fmt:
        pg.set(qn("w:fmt"), fmt)
    if start is not None:
        pg.set(qn("w:start"), str(start))
    sectpr.append(pg)


def _add_section_ref(sectpr, tag: str, ref_type: str, rel_id: Optional[str]):
    if not rel_id:
        return
    ref = OxmlElement(tag)
    ref.set(qn("w:type"), ref_type)
    ref.set(qn("r:id"), rel_id)
    sectpr.insert(0, ref)


def _rel_target_number(rel) -> int:
    m = re.search(r"(\d+)\.xml$", rel.target_ref or "")
    return int(m.group(1)) if m else 0


def _rel_blob(rel) -> str:
    try:
        return rel.target_part.blob.decode("utf-8", "ignore")
    except Exception:
        return ""


def _doc_style_names_by_id(doc) -> dict:
    return {
        getattr(style, "style_id", ""): getattr(style, "name", "")
        for style in doc.styles
        if getattr(style, "style_id", "")
    }


def _part_paragraph_style_names(part, style_names_by_id: dict) -> List[str]:
    names = []
    element = getattr(part, "element", None)
    if element is None:
        return names
    for paragraph in element.iter(qn("w:p")):
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is None:
            continue
        pstyle = ppr.find(qn("w:pStyle"))
        if pstyle is None:
            continue
        style_id = pstyle.get(qn("w:val"))
        if style_id:
            names.append(style_names_by_id.get(style_id, style_id))
    return names


def _style_id_by_name(doc, style_name: str) -> Optional[str]:
    try:
        return doc.styles[style_name].style_id
    except KeyError:
        return None


def _set_part_paragraph_style(part, style_id: Optional[str]):
    if not style_id:
        return
    element = getattr(part, "element", None)
    if element is None:
        return
    for paragraph in element.iter(qn("w:p")):
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            paragraph.insert(0, ppr)
        pstyle = ppr.find(qn("w:pStyle"))
        if pstyle is None:
            pstyle = OxmlElement("w:pStyle")
            ppr.insert(0, pstyle)
        pstyle.set(qn("w:val"), style_id)
        # Let the applied standard header/footer style control alignment.
        # Some blueprint parts keep direct <w:jc w:val="right"/>, which
        # overrides the even-page header style until the style is reapplied in Word.
        for tag in ("w:jc",):
            old = ppr.find(qn(tag))
            if old is not None:
                ppr.remove(old)


def _choose_odd_even_rels(items, style_names_by_id: dict, odd_marker: str, even_marker: str):
    """Return (odd_rid, even_rid), preferring explicit standard odd/even styles."""
    if not items:
        return None, None
    odd = None
    even = None
    for item in items:
        names = _part_paragraph_style_names(item[1].target_part, style_names_by_id)
        if odd is None and any(odd_marker in name for name in names):
            odd = item
        if even is None and any(even_marker in name for name in names):
            even = item
    if odd is None:
        odd = items[0]
    if even is None:
        even = next((item for item in items if item is not odd), odd)
    return odd[0], even[0]


def _normalize_body_page_ref_styles(doc, refs: dict):
    style_map = {
        "header_default": _style_id_by_name(doc, S.S_HEADER_ODD),
        "header_even": _style_id_by_name(doc, S.S_HEADER_EVEN),
        "footer_default": _style_id_by_name(doc, S.S_FOOTER_ODD),
        "footer_even": _style_id_by_name(doc, S.S_FOOTER_EVEN),
    }
    applied_rids = set()
    for key, style_id in style_map.items():
        rid = refs.get(key)
        if not rid or rid not in doc.part.rels:
            continue
        if rid in applied_rids:
            continue
        _set_part_paragraph_style(doc.part.rels[rid].target_part, style_id)
        applied_rids.add(rid)


def _body_page_refs(doc):
    """查找封面蓝图中已打包的正文页眉/页脚关系。"""
    styleref_headers = []
    page_footers = []
    style_names_by_id = _doc_style_names_by_id(doc)
    for rid, rel in doc.part.rels.items():
        reltype = rel.reltype.rsplit("/", 1)[-1]
        blob = _rel_blob(rel)
        if reltype == "header" and "STYLEREF" in blob:
            styleref_headers.append((rid, rel))
        elif reltype == "footer" and "PAGE" in blob:
            page_footers.append((rid, rel))

    styleref_headers.sort(key=lambda item: _rel_target_number(item[1]))
    page_footers.sort(key=lambda item: _rel_target_number(item[1]))

    header_default, header_even = _choose_odd_even_rels(
        styleref_headers, style_names_by_id, "页眉奇数页", "页眉偶数页"
    )
    footer_default, footer_even = _choose_odd_even_rels(
        page_footers, style_names_by_id, "页脚奇数页", "页脚偶数页"
    )
    refs = {
        "header_default": header_default,
        "header_even": header_even,
        "footer_default": footer_default,
        "footer_even": footer_even,
    }
    _normalize_body_page_ref_styles(doc, refs)
    return refs


def _configure_section(sectpr, refs: Optional[dict] = None,
                       page_fmt: Optional[str] = None,
                       page_start: Optional[int] = None,
                       include_even: bool = False):
    type_el = sectpr.find(qn("w:type"))
    if type_el is None:
        type_el = OxmlElement("w:type")
        sectpr.insert(0, type_el)
    type_el.set(qn("w:val"), "nextPage")
    _clear_section_header_footer_refs(sectpr)
    if refs:
        _add_section_ref(sectpr, "w:headerReference", "default", refs.get("header_default"))
        _add_section_ref(sectpr, "w:footerReference", "default", refs.get("footer_default"))
        if include_even:
            _add_section_ref(sectpr, "w:headerReference", "even", refs.get("header_even"))
            _add_section_ref(sectpr, "w:footerReference", "even", refs.get("footer_even"))
    _set_section_page_number(sectpr, fmt=page_fmt, start=page_start)


def _new_section_break_before(anchor: Paragraph, doc, refs: Optional[dict] = None,
                              page_fmt: Optional[str] = None,
                              page_start: Optional[int] = None,
                              include_even: bool = False):
    """在 anchor 前插入下一页分节符，cover 后端按模板节结构切分各部分。"""
    new_p = OxmlElement("w:p")
    anchor._p.addprevious(new_p)
    para = Paragraph(new_p, anchor._parent)
    ppr = para._p.get_or_add_pPr()
    sectpr = copy.deepcopy(doc.sections[-1]._sectPr)
    _configure_section(
        sectpr,
        refs=refs,
        page_fmt=page_fmt,
        page_start=page_start,
        include_even=include_even,
    )
    ppr.append(sectpr)
    return para


def _configure_final_section(doc, refs: Optional[dict] = None,
                             page_start: Optional[int] = None,
                             include_even: bool = False):
    sectpr = doc.sections[-1]._sectPr
    _configure_section(
        sectpr,
        refs=refs,
        page_start=page_start,
        include_even=include_even,
    )


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


def _prev_sectpr_para(doc, before: int) -> Optional[int]:
    """从 before 之前向前找最近的承载 sectPr 的段落索引。"""
    for i in range(before - 1, -1, -1):
        if _carries_sectpr(doc.paragraphs[i]):
            return i
    return None


def _find_first_appendix_mark(doc, before: Optional[int] = None) -> Optional[int]:
    """查找模板中第一个附录标识段，用于分离正文样例和附录样例。"""
    end = before if before is not None else len(doc.paragraphs)
    for i in range(0, end):
        p = doc.paragraphs[i]
        if p.style is not None and p.style.name == S.S_APPENDIX_MARK:
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


def _set_even_and_odd_headers(doc, enabled: bool):
    settings = doc.settings.element
    old = settings.find(qn("w:evenAndOddHeaders"))
    if enabled:
        if old is None:
            settings.append(OxmlElement("w:evenAndOddHeaders"))
    elif old is not None:
        settings.remove(old)


def _normalize_cover_page_number_for_odd_even_export(doc, enabled: bool):
    if not enabled or not doc.sections:
        return
    # Word's PDF export inserts a blank page between an unnumbered cover and a
    # front-matter section that restarts at roman I when odd/even pages are on.
    # Starting the hidden cover section at 0 keeps the TOC on the next physical page.
    _set_section_page_number(doc.sections[0]._sectPr, start=0)


__all__ = [name for name in globals() if name.startswith("_") or name.startswith("SEQ_")]
