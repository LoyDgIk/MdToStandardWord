# -*- coding: utf-8 -*-
"""把 Markdown（YAML front matter + 正文）解析为 model.StandardDoc。

解析分两步：
1. 用 markdown-it-py 把正文 tokenize，归并成"原始块"列表（标题/段/列表/表/图）。
2. 路由：按 H1 把内容分流到 正文 / 附录 / 参考文献 / 索引，并识别 注、示例、图、表标题。
"""

from __future__ import annotations

import re
import warnings
from html.parser import HTMLParser
from typing import List, Tuple

import yaml
from markdown_it import MarkdownIt

from . import model

# --------------------------------------------------------------------------- #
# YAML front matter
# --------------------------------------------------------------------------- #
_FM_RE = re.compile(r"^﻿?---\s*\n(.*?)\n---\s*\n?", re.S)


def split_front_matter(text: str) -> Tuple[dict, str]:
    """返回 (元数据 dict, 正文 markdown)。无 front matter 时 dict 为空。"""
    m = _FM_RE.match(text)
    if not m:
        return {}, text
    data = yaml.safe_load(m.group(1)) or {}
    if not isinstance(data, dict):
        data = {}
    return data, text[m.end():]


# 块级公式：$$ latex $$ 可后跟 {#eq:id}。在 markdown 解析前抽取，避免 LaTeX 被 markdown 改写。
_FORMULA_LINE_RE = re.compile(
    r"^[ \t]*\$\$(.+?)\$\$[ \t]*(?:\{#([^}]+)\})?[ \t]*$", re.M)
_FORMULA_PLACEHOLDER_RE = re.compile(r"^\s*\[\[\[MD2STD-FORMULA-(\d+)\]\]\]\s*$")
_INLINE_FORMULA_RE = re.compile(r"\$\$(.+?)\$\$")
# 类型化锚点 {#tbl:id} / {#fig:id} / {#eq:id}
_ANCHOR_RE = re.compile(r"\{#([^}\s]+)\}")
_LEGACY_REF_RE = re.compile(r"\{@[^}]+\}")
_INLINE_REF_TOKEN_RE = re.compile(r"(\{\{[^{}]+\}\}|\{@[^}]+\})")
_NEW_REF_RE = re.compile(r"^\{\{([^{}]+)\}\}$")
_ANCHOR_ID_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_.-]*$")
_ALLOWED_REF_TYPES = {"tbl", "fig", "eq", "std"}
_ALLOWED_REF_MODES = {"num", "label", "full"}
_CAPTION_NUMBER_RE = {
    "tbl": re.compile(r"^\s*表\s*(?:[A-ZＡ-Ｚ]\s*[.．]\s*)?\d+(?:[.．]\d+)?\s*[\t 　]+"),
    "fig": re.compile(r"^\s*图\s*(?:[A-ZＡ-Ｚ]\s*[.．]\s*)?\d+(?:[.．]\d+)?\s*[\t 　]+"),
}
_FORMULA_TAG_RE = re.compile(r"\\tag\s*\{?")
# 显式分页控制符：推荐写法为独立一行 `<!-- pagebreak -->`。
_PAGE_BREAK_RE = re.compile(
    r"^\s*(?:<!--\s*(?:md2std:)?pagebreak\s*-->|\\pagebreak|\\newpage|\[pagebreak\])\s*$",
    re.I,
)
# 规范性引用文件条目："标准号  标准名称"。标准号允许不带年份。
_NORMREF_RE = re.compile(r"^\s*([A-Z][A-Z/]*\s+\d[\w.\-—–]*)(?:\s{2,}|　+)(.+)$")


def _pop_anchor(text):
    """从文本中取出首个 `{#type:id}`，返回 (去锚点文本, id)。"""
    m = _ANCHOR_RE.search(text)
    if not m:
        return text.strip(), ""
    anchor = m.group(1)
    cleaned = (text[:m.start()] + text[m.end():]).strip()
    return cleaned, anchor


def _parse_typed_anchor(anchor: str, expected_type: str, context: str) -> str:
    """校验 `{#tbl:id}` / `{#fig:id}` / `{#eq:id}`，返回本地 id。"""
    if not anchor:
        return ""
    if ":" not in anchor:
        raise ValueError("%s 的锚点必须写成 #%s:id，例如 #%s:demo。" %
                         (context, expected_type, expected_type))
    ref_type, local_id = anchor.split(":", 1)
    ref_type = ref_type.strip()
    local_id = local_id.strip()
    if ref_type != expected_type:
        raise ValueError("%s 的锚点类型应为 #%s:id，实际为 #%s。" %
                         (context, expected_type, anchor))
    if not _ANCHOR_ID_RE.match(local_id):
        raise ValueError("%s 的锚点 id 只能包含字母、数字、下划线、点和连字符，且需以字母开头：%s。" %
                         (context, local_id))
    return local_id


def _assert_clean_caption(ref_type: str, title: str, context: str):
    if ref_type in _CAPTION_NUMBER_RE and _CAPTION_NUMBER_RE[ref_type].match(title or ""):
        raise ValueError("%s 标题只写纯标题，不要手写编号：%s。" % (context, title))


def _assert_clean_formula(latex: str, context: str):
    if _FORMULA_TAG_RE.search(latex or ""):
        raise ValueError("%s 不要使用 LaTeX \\tag 手写编号，编号由 Word SEQ 自动生成。" % context)


def _is_page_break_marker(text: str) -> bool:
    return _PAGE_BREAK_RE.match(text or "") is not None


def _parse_ref_token(token: str, bold: bool = False, italic: bool = False) -> model.RefSpan:
    if _LEGACY_REF_RE.match(token):
        raise ValueError("旧交叉引用语法 %s 已废弃，请改用 {{tbl:id}} / {{fig:id}} / {{eq:id}} / {{std:标准号}}。" %
                         token)
    m = _NEW_REF_RE.match(token)
    if not m:
        raise ValueError("无效交叉引用语法：%s。" % token)
    raw = m.group(1).strip()
    parts = [x.strip() for x in raw.split(":")]
    ref_type = parts[0] if parts else ""
    if ref_type not in _ALLOWED_REF_TYPES:
        raise ValueError("未知交叉引用类型：%s。支持 tbl、fig、eq、std。" % ref_type)
    if ref_type == "std":
        if len(parts) != 2 or not parts[1]:
            raise ValueError("规范性引用文件引用应写成 {{std:标准号}}。")
        return model.RefSpan(
            text=token, bold=bold, italic=italic,
            ref_type=ref_type, target=parts[1], mode="num",
        )
    if len(parts) not in (2, 3) or not parts[1]:
        raise ValueError("图表公式引用应写成 {{%s:id}} 或 {{%s:id:label/full}}。" %
                         (ref_type, ref_type))
    target = parts[1]
    mode = parts[2] if len(parts) == 3 else "num"
    if mode not in _ALLOWED_REF_MODES:
        raise ValueError("交叉引用修饰符只支持 num、label、full：%s。" % token)
    if not _ANCHOR_ID_RE.match(target):
        raise ValueError("交叉引用 id 只能包含字母、数字、下划线、点和连字符，且需以字母开头：%s。" %
                         target)
    return model.RefSpan(
        text=token, bold=bold, italic=italic,
        ref_type=ref_type, target=target, mode=mode,
    )


def _split_reference_spans(spans: List[model.Span]) -> List[model.Span]:
    out: List[model.Span] = []
    for sp in spans:
        if isinstance(sp, model.FormulaSpan):
            out.append(sp)
            continue
        pos = 0
        for m in _INLINE_REF_TOKEN_RE.finditer(sp.text):
            if m.start() > pos:
                chunk = sp.text[pos:m.start()]
                if "{{" in chunk or "}}" in chunk or "{@" in chunk:
                    raise ValueError("无效交叉引用语法：%s。" % chunk)
                out.append(model.Span(chunk, sp.bold, sp.italic, sp.subscript, sp.superscript))
            out.append(_parse_ref_token(m.group(0), sp.bold, sp.italic))
            pos = m.end()
        if pos < len(sp.text):
            tail = sp.text[pos:]
            if "{{" in tail or "}}" in tail or "{@" in tail:
                raise ValueError("无效交叉引用语法：%s。" % tail)
            out.append(model.Span(tail, sp.bold, sp.italic, sp.subscript, sp.superscript))
    return out


def _append_inline_text_spans(
    spans: List[model.Span],
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    subscript: bool = False,
    superscript: bool = False,
):
    """Split inline formulas and Typora/Obsidian-style ^sup^/~sub~ markers into spans."""
    if not text:
        return
    pos = 0
    for m in _INLINE_FORMULA_RE.finditer(text):
        if m.start() > pos:
            _append_script_text_spans(
                spans,
                text[pos:m.start()],
                bold=bold,
                italic=italic,
                subscript=subscript,
                superscript=superscript,
            )
        latex = (m.group(1) or "").strip()
        if latex:
            _assert_clean_formula(latex, "行内公式")
            spans.append(model.FormulaSpan(latex, bold=bold, italic=italic))
        pos = m.end()
    if pos < len(text):
        _append_script_text_spans(
            spans,
            text[pos:],
            bold=bold,
            italic=italic,
            subscript=subscript,
            superscript=superscript,
        )


def _append_script_text_spans(
    spans: List[model.Span],
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    subscript: bool = False,
    superscript: bool = False,
):
    if not text:
        return
    pos = 0
    i = 0
    while i < len(text):
        marker = text[i]
        if marker not in ("^", "~"):
            i += 1
            continue
        j = text.find(marker, i + 1)
        if j <= i + 1:
            i += 1
            continue
        if i > pos:
            spans.append(model.Span(text[pos:i], bold, italic, subscript, superscript))
        marked = text[i + 1:j]
        spans.append(model.Span(
            marked,
            bold,
            italic,
            subscript or marker == "~",
            superscript or marker == "^",
        ))
        i = j + 1
        pos = i
    if pos < len(text):
        spans.append(model.Span(text[pos:], bold, italic, subscript, superscript))


def extract_formulas(md: str):
    """把块级 $$...$$ 公式行替换为占位段，返回 (新markdown, [Formula,...])。"""
    formulas = []

    def repl(m):
        i = len(formulas)
        latex = m.group(1).strip()
        _assert_clean_formula(latex, "公式")
        anchor = _parse_typed_anchor((m.group(2) or "").strip(), "eq", "公式")
        formulas.append(model.Formula(latex=latex, anchor_id=anchor))
        return "\n[[[MD2STD-FORMULA-%d]]]\n" % i

    return _FORMULA_LINE_RE.sub(repl, md), formulas


def _normalize_extra_note(value):
    if isinstance(value, (list, tuple)):
        return [str(x) for x in value if str(x).strip()]
    return str(value)


def _as_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    text = str(value).strip().lower()
    return text in ("1", "true", "yes", "y", "on", "是", "启用", "开启")


def build_meta(data: dict) -> model.Meta:
    fw_raw = data.get("foreword") or {}
    fw = model.Foreword(
        multipart_note=str(fw_raw.get("multipart_note", "") or ""),
        replace_changes=[str(x) for x in (fw_raw.get("replace_changes") or [])],
        patent_note=bool(fw_raw.get("patent_note", True)),
        proposer=str(fw_raw.get("proposer", "") or ""),
        owner=str(fw_raw.get("owner", "") or ""),
        draft_orgs=[str(x) for x in (fw_raw.get("draft_orgs") or [])],
        drafters=[str(x) for x in (fw_raw.get("drafters") or [])],
        history=str(fw_raw.get("history", "") or ""),
        extra_notes=[_normalize_extra_note(x) for x in (fw_raw.get("extra_notes") or [])],
    )

    def s(key):
        v = data.get(key, "")
        return "" if v is None else str(v)

    return model.Meta(
        standard_type=s("standard_type") or "团体标准",
        number=s("number"),
        replaces=s("replaces"),
        title=s("title"),
        title_en=s("title_en"),
        consistency_degree=s("consistency_degree") or s("consistency"),
        draft_version=s("draft_version") or s("draft_stage") or s("草案版次"),
        ics=s("ics"),
        ccs=s("ccs"),
        record_number=s("record_number") or s("filing_number") or s("备案号"),
        publish_date=s("publish_date"),
        implement_date=s("implement_date"),
        publisher=s("publisher"),
        foreword=fw,
        introduction=s("introduction"),
        important_notice=s("important_notice") or s("重要提示"),
        symbols_lead=s("symbols_lead"),
        odd_even_pages=_as_bool(data.get("odd_even_pages", False)),
        cover_form_protection=_as_bool(data.get("cover_form_protection", False)),
    )


# --------------------------------------------------------------------------- #
# 行内 -> spans
# --------------------------------------------------------------------------- #
def _inline_to_spans(inline_token) -> List[model.Span]:
    spans: List[model.Span] = []
    bold = 0
    italic = 0
    subscript = 0
    superscript = 0
    for ch in (inline_token.children or []):
        t = ch.type
        if t == "strong_open":
            bold += 1
        elif t == "strong_close":
            bold = max(0, bold - 1)
        elif t == "em_open":
            italic += 1
        elif t == "em_close":
            italic = max(0, italic - 1)
        elif t == "html_inline":
            content = (ch.content or "").strip().lower()
            if content.startswith("<sub"):
                subscript += 1
            elif content.startswith("</sub"):
                subscript = max(0, subscript - 1)
            elif content.startswith("<sup"):
                superscript += 1
            elif content.startswith("</sup"):
                superscript = max(0, superscript - 1)
        elif t in ("text", "code_inline"):
            if ch.content:
                _append_inline_text_spans(
                    spans,
                    ch.content,
                    bold=bold > 0,
                    italic=italic > 0,
                    subscript=subscript > 0,
                    superscript=superscript > 0,
                )
        elif t == "softbreak":
            # 段内软换行：中文文本直接连接
            pass
        elif t == "hardbreak":
            spans.append(model.Span(
                "\n",
                bold=bold > 0,
                italic=italic > 0,
                subscript=subscript > 0,
                superscript=superscript > 0,
            ))
        # image / link 文本由其它路径处理
    # 合并相邻同格式片段
    merged: List[model.Span] = []
    for sp in spans:
        if isinstance(sp, model.FormulaSpan):
            merged.append(sp)
            continue
        if (
            merged
            and not isinstance(merged[-1], model.FormulaSpan)
            and merged[-1].bold == sp.bold
            and merged[-1].italic == sp.italic
            and merged[-1].subscript == sp.subscript
            and merged[-1].superscript == sp.superscript
        ):
            merged[-1] = model.Span(
                merged[-1].text + sp.text,
                sp.bold,
                sp.italic,
                sp.subscript,
                sp.superscript,
            )
        else:
            merged.append(sp)
    return _split_reference_spans(merged)


def _inline_image(inline_token):
    """若 inline 仅含一张图片，返回 (alt, src)，否则 None。"""
    kids = [c for c in (inline_token.children or []) if c.type not in ("softbreak",)]
    imgs = [c for c in kids if c.type == "image"]
    texts = [c for c in kids if c.type == "text" and c.content.strip()]
    if len(imgs) == 1 and not texts:
        img = imgs[0]
        alt = "".join(g.content for g in (img.children or []) if g.type == "text")
        return alt, img.attrs.get("src", "")
    return None


# --------------------------------------------------------------------------- #
# tokens -> 原始块
# --------------------------------------------------------------------------- #
# 原始块用元组表示：
#   ('heading', level, spans)
#   ('para',  spans)
#   ('image', alt, src)
#   ('list',  ordered:bool, items:List[List[Span]], level:int)
#   ('table', header, rows, header_colspans, row_colspans, header_parts, row_parts)
#   ('pagebreak',)
def _tokens_to_blocks(tokens, level=1):
    blocks = []
    i = 0
    n = len(tokens)
    while i < n:
        tok = tokens[i]
        tt = tok.type
        if tt == "heading_open":
            lvl = int(tok.tag[1])
            inline = tokens[i + 1]
            blocks.append(("heading", lvl, _inline_to_spans(inline)))
            i += 3  # open, inline, close
        elif tt == "paragraph_open":
            inline = tokens[i + 1]
            img = _inline_image(inline)
            if img is not None:
                blocks.append(("image", img[0], img[1]))
            else:
                blocks.append(("para", _inline_to_spans(inline)))
            i += 3
        elif tt in ("bullet_list_open", "ordered_list_open"):
            ordered = tt == "ordered_list_open"
            j, items, nested_blocks = _parse_list(tokens, i, level)
            blocks.append(("list", ordered, items, level))
            blocks.extend(nested_blocks)
            i = j
        elif tt == "table_open":
            j, header, rows, header_colspans, row_colspans, header_parts, row_parts = _parse_table(tokens, i)
            blocks.append(("table", header, rows, header_colspans, row_colspans, header_parts, row_parts))
            i = j
        elif tt == "html_block":
            parsed = _parse_html_table_block(tok.content)
            if parsed is not None:
                header, rows, header_colspans, row_colspans, header_parts, row_parts = parsed
                blocks.append(("table", header, rows, header_colspans, row_colspans, header_parts, row_parts))
            elif _is_page_break_marker(tok.content):
                blocks.append(("pagebreak",))
            i += 1
        elif tt in ("fence", "code_block"):
            # 代码块当作普通段落原样输出
            for line in tok.content.rstrip("\n").split("\n"):
                blocks.append(("para", [model.Span(line)]))
            i += 1
        elif tt == "blockquote_open":
            # 引用块：取内部段落，原样作为段
            depth = 1
            i += 1
            while i < n and depth > 0:
                if tokens[i].type == "blockquote_open":
                    depth += 1
                elif tokens[i].type == "blockquote_close":
                    depth -= 1
                elif tokens[i].type == "inline":
                    blocks.append(("para", _inline_to_spans(tokens[i])))
                i += 1
        else:
            i += 1
    return blocks


def _parse_list(tokens, start, level):
    """解析 list，返回 (下一个索引, items, nested_blocks)。items 为 List[List[Span]]。

    简化处理：列项内的首段作为列项文本；列项内嵌套列表作为更深一级（level+1）
    追加为后续 ('list', ...) 由调用者处理。
    """
    items: List[List[model.Span]] = []
    nested_blocks = []
    i = start + 1  # skip *_list_open
    n = len(tokens)
    while i < n:
        tt = tokens[i].type
        if tt in ("bullet_list_close", "ordered_list_close"):
            return i + 1, items, nested_blocks
        if tt == "list_item_open":
            # 收集该列项内首个 inline 段
            spans: List[model.Span] = []
            k = i + 1
            iddepth = 1
            captured = False
            while k < n and iddepth > 0:
                ttype = tokens[k].type
                if ttype in ("bullet_list_open", "ordered_list_open") and iddepth == 1:
                    ordered = ttype == "ordered_list_open"
                    j, sub_items, sub_nested = _parse_list(tokens, k, level + 1)
                    if sub_items:
                        nested_blocks.append(("list", ordered, sub_items, level + 1))
                    nested_blocks.extend(sub_nested)
                    k = j
                    continue
                if ttype == "list_item_open":
                    iddepth += 1
                elif ttype == "list_item_close":
                    iddepth -= 1
                elif ttype == "inline" and not captured and iddepth == 1:
                    spans = _inline_to_spans(tokens[k])
                    captured = True
                k += 1
            items.append(spans)
            i = k
            continue
        i += 1
    return i, items, nested_blocks


def _parse_table(tokens, start):
    header: List[str] = []
    rows: List[List[str]] = []
    header_parts: List[List[model.TableCellPart]] = []
    row_parts: List[List[List[model.TableCellPart]]] = []
    i = start + 1
    n = len(tokens)
    cur: List[str] = []
    cur_parts: List[List[model.TableCellPart]] = []
    in_head = False
    while i < n:
        tt = tokens[i].type
        if tt == "table_close":
            i += 1
            break
        if tt == "thead_open":
            in_head = True
        elif tt == "thead_close":
            in_head = False
        elif tt == "tr_open":
            cur = []
            cur_parts = []
        elif tt == "tr_close":
            if in_head:
                header = cur
                header_parts = cur_parts
            else:
                rows.append(cur)
                row_parts.append(cur_parts)
        elif tt in ("th_open", "td_open"):
            inline = tokens[i + 1] if i + 1 < n and tokens[i + 1].type == "inline" else None
            parts = _parse_table_cell_parts(inline.content if inline is not None else "")
            txt = _table_parts_text(parts)
            cur.append(txt)
            cur_parts.append(parts)
        i += 1
    return (
        i,
        header,
        rows,
        [1 for _ in header],
        [[1 for _ in row] for row in rows],
        header_parts,
        row_parts,
    )


_TABLE_CELL_PART_RE = re.compile(r"(<eq>(.*?)</eq>|\$([^$]+)\$)", re.S | re.I)


def _parse_table_cell_parts(text: str) -> List[model.TableCellPart]:
    parts: List[model.TableCellPart] = []
    pos = 0
    for m in _TABLE_CELL_PART_RE.finditer(text or ""):
        if m.start() > pos:
            _append_text_part(parts, _normalize_html_cell_text(text[pos:m.start()]))
        formula = (m.group(2) if m.group(2) is not None else m.group(3) or "").strip()
        if formula:
            parts.append(model.TableCellPart("formula", formula))
        pos = m.end()
    if pos < len(text or ""):
        _append_text_part(parts, _normalize_html_cell_text((text or "")[pos:]))
    if not parts:
        parts.append(model.TableCellPart("text", ""))
    return parts


def _append_text_part(parts: List[model.TableCellPart], text: str):
    if not text:
        return
    pos = 0
    for m in _INLINE_REF_TOKEN_RE.finditer(text):
        if m.start() > pos:
            chunk = text[pos:m.start()]
            if "{{" in chunk or "}}" in chunk or "{@" in chunk:
                raise ValueError("无效交叉引用语法：%s。" % chunk)
            _append_plain_text_part(parts, chunk)
        ref = _parse_ref_token(m.group(0))
        parts.append(model.TableCellPart(
            "ref",
            ref.text,
            ref_type=ref.ref_type,
            target=ref.target,
            mode=ref.mode,
        ))
        pos = m.end()
    if pos < len(text):
        tail = text[pos:]
        if "{{" in tail or "}}" in tail or "{@" in tail:
            raise ValueError("无效交叉引用语法：%s。" % tail)
        _append_plain_text_part(parts, tail)


def _append_plain_text_part(parts: List[model.TableCellPart], text: str):
    if not text:
        return
    if parts and parts[-1].kind == "text":
        parts[-1].text += text
    else:
        parts.append(model.TableCellPart("text", text))


def _formula_display_text(latex: str) -> str:
    text = latex or ""
    text = text.replace(r"v_{\text{max}}", "vmax")
    text = text.replace(r"v_{\mathrm{max}}", "vmax")
    text = text.replace(r"\leqslant", "≤")
    text = text.replace(r"\leq", "≤")
    text = text.replace(r"\geqslant", "≥")
    text = text.replace(r"\geq", "≥")
    text = text.replace(r"\text{max}", "max")
    return re.sub(r"\s+", " ", text).strip()


def _table_parts_text(parts: List[model.TableCellPart]) -> str:
    out = []
    for part in parts:
        if part.kind == "formula":
            out.append(_formula_display_text(part.text))
        elif part.kind == "ref":
            if part.ref_type == "std":
                out.append(part.target)
            elif part.ref_type == "eq" and part.mode in ("label", "full"):
                out.append("式（?）")
            elif part.ref_type == "tbl":
                out.append("表?")
            elif part.ref_type == "fig":
                out.append("图?")
            else:
                out.append("?")
        else:
            out.append(part.text)
    return "".join(out).strip()


class _HtmlTableParser(HTMLParser):
    """提取简单 HTML 表格，支持 th/td、colspan 和 br。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows = []
        self._row = None
        self._cell = None
        self._cell_parts = None
        self._capture = False
        self._cell_is_header = False
        self._cell_colspan = 1
        self._inside_eq = False
        self._eq_buf = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            attr_map = {k.lower(): v for k, v in attrs}
            try:
                colspan = int(attr_map.get("colspan") or "1")
            except ValueError:
                colspan = 1
            self._cell = []
            self._cell_parts = []
            self._capture = True
            self._cell_is_header = tag == "th"
            self._cell_colspan = max(1, colspan)
        elif tag == "eq" and self._capture and self._cell_parts is not None:
            self._inside_eq = True
            self._eq_buf = []
        elif tag == "br" and self._capture and self._cell is not None:
            self._cell.append("\n")
            _append_text_part(self._cell_parts, "\n")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("td", "th") and self._row is not None and self._cell is not None:
            text = _table_parts_text(self._cell_parts or [])
            self._row.append({
                "text": text,
                "parts": self._cell_parts or [model.TableCellPart("text", text)],
                "colspan": self._cell_colspan,
                "header": self._cell_is_header,
            })
            self._cell = None
            self._cell_parts = None
            self._capture = False
            self._cell_is_header = False
            self._cell_colspan = 1
        elif tag == "eq" and self._capture and self._cell_parts is not None:
            formula = "".join(self._eq_buf).strip()
            if formula:
                self._cell_parts.append(model.TableCellPart("formula", formula))
            self._inside_eq = False
            self._eq_buf = []
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._capture and self._cell is not None:
            self._cell.append(data)
            if self._inside_eq:
                self._eq_buf.append(data)
            else:
                _append_text_part(self._cell_parts, _normalize_html_cell_text(data))


def _normalize_html_cell_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    text = text.replace(r"v_{\text{max}}", "vmax")
    text = text.replace(r"v_{\mathrm{max}}", "vmax")
    return text


def _parse_html_table_block(content: str):
    if "<table" not in content.lower():
        return None
    parser = _HtmlTableParser()
    parser.feed(content)
    rows = parser.rows
    if not rows:
        return None
    first = rows[0]
    use_first_as_header = any(cell["header"] for cell in first) or len(rows) > 1
    if use_first_as_header:
        header = [cell["text"] for cell in first]
        header_parts = [cell["parts"] for cell in first]
        header_colspans = [cell["colspan"] for cell in first]
        data_rows = rows[1:]
    else:
        header = []
        header_parts = []
        header_colspans = []
        data_rows = rows
    body_rows = [[cell["text"] for cell in row] for row in data_rows]
    row_parts = [[cell["parts"] for cell in row] for row in data_rows]
    row_colspans = [[cell["colspan"] for cell in row] for row in data_rows]
    return header, body_rows, header_colspans, row_colspans, header_parts, row_parts


# --------------------------------------------------------------------------- #
# 原始块 -> StandardDoc（路由）
# --------------------------------------------------------------------------- #
_NOTE_RE = re.compile(r"^\s*注\s*(\d+)?\s*[:：]")
_EXAMPLE_RE = re.compile(r"^\s*示例\s*(\d+)?\s*[:：]")
_SOURCE_RE = re.compile(r"^\s*\[?\s*来源\s*[:：]")
_TABLE_NOTE_RE = re.compile(r"^\s*\{\s*表注\s*(?P<index>\d+)?\s*[:：]\s*(?P<content>.*?)\s*\}\s*$")
_FIGURE_NOTE_RE = re.compile(r"^\s*\{\s*图注\s*(?P<index>\d+)?\s*[:：]\s*(?P<content>.*?)\s*\}\s*$")
_TABLE_SOURCE_RE = re.compile(r"^\s*\{\s*表来源\s*[:：]\s*(?P<content>.*?)\s*\}\s*$")
_FIGURE_SOURCE_RE = re.compile(r"^\s*\{\s*图来源\s*[:：]\s*(?P<content>.*?)\s*\}\s*$")
_TERM_MARKER_RE = re.compile(r"^\s*\{\s*术语\s*[:：]\s*(.+?)\s*\}\s*$")
_TERM_SPLIT_RE = re.compile(r"^(.*?)(?:\s*[|｜]\s*|[\s　]{2,})(.+)$")
_TABLE_CAP_RE = re.compile(r"^\s*\{\s*表\s*[:：]\s*#([^}\s]+)\s*\}\s+(.+?)\s*$")
_TABLE_CAP_MARKER_RE = re.compile(r"^\s*(?:\{\s*)?表\s*[:：]")
# 无标题条：显式声明编号层级，不手写具体编号。
_UNTITLED_RE = re.compile(r"^\s*\{\s*无标题条\s*[:：]\s*([2-6])\s*\}\s*(\S.*)$", re.S)
_LEGACY_UNTITLED_RE = re.compile(r"^\s*(\d+(?:\.\d+)+)\s+(\S.*)$", re.S)
_APPENDIX_RE = re.compile(r"^\s*附录\s*(?:[A-ZＡ-Ｚ]\s*)?[（(]?\s*(规范性|资料性)\s*[)）]?\s*(.*)$")
_INDEX_GROUP_RE = re.compile(r"^[A-Z]$")
_INDEX_ITEM_RE = re.compile(r"^\s*(.+?)\s*[:：]\s*(.+?)\s*$")


def _spans_text(spans):
    return "".join(s.text for s in spans)


def _example_has_inline_content(example: model.Example) -> bool:
    text = _spans_text(example.spans)
    return _EXAMPLE_RE.sub("", text, count=1).strip() != ""


def _marker_content_start(text: str, marker_end: int) -> int:
    pos = marker_end
    while pos < len(text) and text[pos].isspace():
        pos += 1
    return pos


def _clone_span_with_text(sp: model.Span, text: str) -> model.Span:
    if isinstance(sp, model.RefSpan):
        return model.RefSpan(
            text=text,
            bold=sp.bold,
            italic=sp.italic,
            subscript=sp.subscript,
            superscript=sp.superscript,
            ref_type=sp.ref_type,
            target=sp.target,
            mode=sp.mode,
        )
    if isinstance(sp, model.FormulaSpan):
        return model.FormulaSpan(
            text,
            bold=sp.bold,
            italic=sp.italic,
            subscript=sp.subscript,
            superscript=sp.superscript,
        )
    return model.Span(text, sp.bold, sp.italic, sp.subscript, sp.superscript)


def _slice_spans(spans: List[model.Span], start: int, end: int) -> List[model.Span]:
    """按 `_spans_text(spans)` 的字符位置截取，保留行内格式和引用片段。"""
    out: List[model.Span] = []
    pos = 0
    for sp in spans:
        sp_end = pos + len(sp.text)
        if sp_end <= start:
            pos = sp_end
            continue
        if pos >= end:
            break
        local_start = max(0, start - pos)
        local_end = min(len(sp.text), end - pos)
        piece = sp.text[local_start:local_end]
        if piece:
            if local_start == 0 and local_end == len(sp.text):
                out.append(sp)
            else:
                out.append(_clone_span_with_text(sp, piece))
        pos = sp_end
    return out


def _parse_figure_table_addon(spans: List[model.Span]):
    text = _spans_text(spans)
    for ref_type, regex in (("tbl", _TABLE_NOTE_RE), ("fig", _FIGURE_NOTE_RE)):
        m = regex.match(text)
        if not m:
            continue
        content_spans = _slice_spans(spans, m.start("content"), m.end("content"))
        return (
            ref_type,
            "note",
            model.Note(
                spans=content_spans,
                index=int(m.group("index")) if m.group("index") else None,
            ),
        )
    for ref_type, regex in (("tbl", _TABLE_SOURCE_RE), ("fig", _FIGURE_SOURCE_RE)):
        m = regex.match(text)
        if not m:
            continue
        content_spans = _slice_spans(spans, m.start("content"), m.end("content"))
        return ref_type, "source", model.FigureTableSource(spans=content_spans)
    return None


def _attach_figure_table_addon(target, addon):
    ref_type, kind, value = addon
    label = "表" if ref_type == "tbl" else "图"
    suffix = "注" if kind == "note" else "来源"
    expected_type = model.TableModel if ref_type == "tbl" else model.Figure
    if not isinstance(target, expected_type):
        raise ValueError("%s%s必须紧跟%s。" % (label, suffix, "表格" if ref_type == "tbl" else "图片"))
    if kind == "note":
        target.notes.append(value)
        return
    if target.source is not None:
        raise ValueError("%s来源只能写一次。" % label)
    target.source = value


def parse(text: str) -> model.StandardDoc:
    data, body_md = split_front_matter(text)
    doc = model.StandardDoc(meta=build_meta(data))

    body_md, formulas = extract_formulas(body_md)
    md = MarkdownIt("commonmark").enable("table")
    tokens = md.parse(body_md)
    blocks = _tokens_to_blocks(tokens)

    mode = "body"             # body | appendix | references | index
    cur_appendix = None
    cur_index_group = None
    table_caption = None      # None 或 (caption_text, verbatim_bool)
    expect_example_content = False
    last_addon_target = None

    idx = 0
    while idx < len(blocks):
        blk = blocks[idx]
        kind = blk[0]

        # --- 标题：可能切换 mode ---
        if kind == "heading":
            last_addon_target = None
            lvl, spans = blk[1], blk[2]
            htext = _spans_text(spans).strip()
            if lvl == 1:
                m_appx = _APPENDIX_RE.match(htext)
                if m_appx:
                    nature = "normative" if m_appx.group(1) == "规范性" else "informative"
                    title = m_appx.group(2).strip()
                    cur_appendix = model.Appendix(
                        nature=nature, title_spans=[model.Span(title)]
                    )
                    doc.appendices.append(cur_appendix)
                    mode = "appendix"
                    cur_index_group = None
                    idx += 1
                    continue
                if htext == "参考文献":
                    mode = "references"
                    cur_index_group = None
                    expect_example_content = False
                    idx += 1
                    continue
                if htext == "索引":
                    mode = "index"
                    cur_index_group = None
                    expect_example_content = False
                    idx += 1
                    continue
                mode = "body"
                cur_index_group = None
                expect_example_content = False
                doc.body.append(model.Heading(level=1, spans=spans))
                idx += 1
                continue
            else:  # lvl >= 2
                heading = model.Heading(level=lvl, spans=spans)
                expect_example_content = False
                if mode == "appendix" and cur_appendix is not None:
                    cur_appendix.blocks.append(heading)
                elif mode == "body":
                    doc.body.append(heading)
                elif mode == "index":
                    if lvl != 2 or not _INDEX_GROUP_RE.match(htext):
                        raise ValueError("索引分组必须写成 `## A` 到 `## Z`：%s。" % htext)
                    cur_index_group = model.IndexGroup(letter=htext)
                    doc.index_groups.append(cur_index_group)
                # references 下的子标题忽略
                idx += 1
                continue

        # --- 块级公式占位段 ---
        if kind == "para":
            ptext0 = _spans_text(blk[1]).strip()
            mf = _FORMULA_PLACEHOLDER_RE.match(ptext0)
            if mf:
                fi = int(mf.group(1))
                target_block = formulas[fi] if 0 <= fi < len(formulas) else None
                if target_block is not None:
                    if mode == "body":
                        doc.body.append(target_block)
                    elif mode == "appendix" and cur_appendix is not None:
                        cur_appendix.blocks.append(target_block)
                last_addon_target = None
                idx += 1
                continue

        # --- 表/图附加项：只绑定紧邻的上一张表或图 ---
        if kind == "para":
            addon = _parse_figure_table_addon(blk[1])
            if addon is not None:
                _attach_figure_table_addon(last_addon_target, addon)
                expect_example_content = False
                idx += 1
                continue

        # --- 表格标题（"{表：#tbl:id} 标题"，后接表格）---
        if kind == "para":
            ptext = _spans_text(blk[1]).strip()
            if idx + 1 < len(blocks) and blocks[idx + 1][0] == "table":
                m_cap = _TABLE_CAP_RE.match(ptext)
                if m_cap:
                    last_addon_target = None
                    anchor = _parse_typed_anchor(m_cap.group(1).strip(), "tbl", "表题")
                    title = m_cap.group(2).strip()
                    _assert_clean_caption("tbl", title, "表题")
                    table_caption = (title, anchor)
                    idx += 1
                    continue
                if _TABLE_CAP_MARKER_RE.match(ptext):
                    raise ValueError("表题必须写成 `{表：#tbl:id} 标题`，且标题不要手写编号：%s。" %
                                     ptext)

        # --- 构造目标块 ---
        target_block = _make_block(kind, blk, table_caption)
        table_caption = None
        if target_block is None:
            last_addon_target = None
            idx += 1
            continue

        if expect_example_content and isinstance(target_block, model.Paragraph):
            target_block = model.ExampleContent(spans=target_block.spans)
        expect_example_content = (
            isinstance(target_block, model.Example)
            and not _example_has_inline_content(target_block)
        )

        can_accept_addons = (
            isinstance(target_block, (model.TableModel, model.Figure))
            and mode in ("body", "appendix")
        )

        if mode == "body":
            doc.body.append(target_block)
        elif mode == "appendix" and cur_appendix is not None:
            cur_appendix.blocks.append(target_block)
        elif mode == "references":
            _route_reference(doc, kind, blk)
        elif mode == "index":
            _route_index(doc, cur_index_group, kind, blk)
        last_addon_target = target_block if can_accept_addons else None
        idx += 1

    _normalize_terms(doc)
    _validate_refs(doc)
    return doc


def _parse_term_title(text: str):
    value = (text or "").strip()
    m = _TERM_SPLIT_RE.match(value)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return value, ""


def _term_marker_text(blk) -> str:
    if not isinstance(blk, model.Paragraph):
        return ""
    m = _TERM_MARKER_RE.match(_spans_text(blk.spans).strip())
    return m.group(1).strip() if m else ""


def _term_from_start_block(blk) -> model.Term:
    if isinstance(blk, model.Heading) and blk.level == 2:
        term, term_en = _parse_term_title(blk.text)
        return model.Term(term=term, term_en=term_en)
    marker = _term_marker_text(blk)
    if marker:
        term, term_en = _parse_term_title(marker)
        return model.Term(term=term, term_en=term_en)
    return None


def _is_term_boundary(blk) -> bool:
    if isinstance(blk, model.Heading) and blk.level <= 2:
        return True
    return bool(_term_marker_text(blk))


def _consume_term_blocks(body: List[object], start: int):
    term = _term_from_start_block(body[start])
    if term is None:
        return None, start
    i = start + 1
    while i < len(body):
        blk = body[i]
        if _is_term_boundary(blk):
            break
        if isinstance(blk, model.Paragraph) and not term.definition:
            term.definition = blk.spans
            i += 1
            continue
        if isinstance(blk, model.Note):
            term.notes.append(blk)
            i += 1
            continue
        if isinstance(blk, model.Source) and term.source is None:
            term.source = blk
            i += 1
            continue
        break
    return term, i


def _normalize_terms(doc: model.StandardDoc):
    """把术语章内的术语标题/显式术语标记归一成 model.Term。"""
    out: List[object] = []
    in_terms = False
    i = 0
    while i < len(doc.body):
        blk = doc.body[i]
        if isinstance(blk, model.Heading) and blk.level == 1:
            in_terms = blk.text.strip() == "术语和定义"
            out.append(blk)
            i += 1
            continue
        if in_terms:
            term, next_i = _consume_term_blocks(doc.body, i)
            if term is not None:
                out.append(term)
                i = next_i
                continue
        out.append(blk)
        i += 1
    doc.body = out


def _make_block(kind, blk, table_caption):
    if kind == "pagebreak":
        return model.PageBreak()
    if kind == "para":
        spans = blk[1]
        ptext = _spans_text(spans)
        if _is_page_break_marker(ptext.strip()):
            return model.PageBreak()
        m = _NOTE_RE.match(ptext)
        if m:
            content_spans = _strip_leading(spans, _marker_content_start(ptext, m.end()))
            return model.Note(
                spans=content_spans,
                index=int(m.group(1)) if m.group(1) else None,
            )
        m = _EXAMPLE_RE.match(ptext)
        if m:
            content_spans = _strip_leading(spans, _marker_content_start(ptext, m.end()))
            return model.Example(
                spans=content_spans,
                index=int(m.group(1)) if m.group(1) else None,
            )
        if _SOURCE_RE.match(ptext):
            return model.Source(text=ptext.strip())
        m = _UNTITLED_RE.match(ptext)
        if m:
            level = int(m.group(1))
            rest_spans = _strip_leading(spans, len(ptext) - len(m.group(2)))
            return model.UntitledClause(level=level, spans=rest_spans)
        m_legacy = _LEGACY_UNTITLED_RE.match(ptext)
        if m_legacy:
            level = len([x for x in m_legacy.group(1).split(".") if x])
            warnings.warn(
                "检测到旧式无标题条手写编号 `%s`；将按普通段落保留原文。"
                "规范写法是 `{无标题条:%d} %s`。" %
                (m_legacy.group(1), level, m_legacy.group(2).strip()),
                UserWarning,
                stacklevel=2,
            )
        return model.Paragraph(spans=spans)
    if kind == "image":
        cap, raw_anchor = _pop_anchor(blk[1])
        _assert_clean_caption("fig", cap, "图题")
        anchor = _parse_typed_anchor(raw_anchor, "fig", "图题")
        return model.Figure(caption=cap, path=blk[2], anchor_id=anchor)
    if kind == "list":
        ordered, items, level = blk[1], blk[2], blk[3]
        return model.ListBlock(
            ordered=ordered,
            level=level,
            items=[model.ListItem(spans=it) for it in items],
        )
    if kind == "table":
        header, rows = blk[1], blk[2]
        header_colspans = blk[3] if len(blk) > 3 else []
        row_colspans = blk[4] if len(blk) > 4 else []
        header_parts = blk[5] if len(blk) > 5 else []
        row_parts = blk[6] if len(blk) > 6 else []
        if table_caption is None:
            cap, anchor = "", ""
        else:
            cap, anchor = table_caption
        return model.TableModel(
            caption=cap,
            anchor_id=anchor,
            header=header,
            rows=rows,
            header_colspans=header_colspans,
            row_colspans=row_colspans,
            header_parts=header_parts,
            row_parts=row_parts,
        )
    return None


def _strip_leading(spans, n):
    """从 spans 序列前部去掉 n 个字符（用于剥离无标题条编号前缀）。"""
    out = []
    remaining = n
    for sp in spans:
        if remaining <= 0:
            out.append(sp)
            continue
        if isinstance(sp, model.FormulaSpan):
            out.append(sp)
            continue
        if len(sp.text) <= remaining:
            remaining -= len(sp.text)
            continue
        out.append(model.Span(sp.text[remaining:], sp.bold, sp.italic, sp.subscript, sp.superscript))
        remaining = 0
    return out


def _iter_block_spans(blk):
    if isinstance(blk, (model.Heading, model.Paragraph, model.UntitledClause,
                        model.Note, model.Example, model.ExampleContent)):
        yield from blk.spans
    elif isinstance(blk, model.Term):
        yield from blk.definition
        for note in blk.notes:
            yield from note.spans
    elif isinstance(blk, model.Figure):
        for note in blk.notes:
            yield from note.spans
        if blk.source is not None:
            yield from blk.source.spans
    elif isinstance(blk, model.ListBlock):
        for item in blk.items:
            yield from item.spans
    elif isinstance(blk, model.TableModel):
        for note in blk.notes:
            yield from note.spans
        if blk.source is not None:
            yield from blk.source.spans
        part_rows = []
        if blk.header_parts:
            part_rows.append(blk.header_parts)
        part_rows.extend(blk.row_parts or [])
        for row in part_rows:
            for cell_parts in row:
                for part in cell_parts:
                    if part.kind == "ref":
                        yield model.RefSpan(
                            text=part.text,
                            ref_type=part.ref_type,
                            target=part.target,
                            mode=part.mode,
                        )


def _collect_normative_refs(doc: model.StandardDoc):
    refs = set()
    in_normrefs = False
    for blk in doc.body:
        if isinstance(blk, model.Heading) and blk.level == 1:
            in_normrefs = blk.text.strip() == "规范性引用文件"
            continue
        if not in_normrefs or not isinstance(blk, model.Paragraph):
            continue
        m = _NORMREF_RE.match(_spans_text(blk.spans).strip())
        if m:
            refs.add(m.group(1).strip())
    return refs


def _validate_refs(doc: model.StandardDoc):
    anchors = {}

    def add_anchor(ref_type, local_id, context):
        if not local_id:
            return
        key = (ref_type, local_id)
        if key in anchors:
            raise ValueError("重复的 %s 交叉引用 id：%s。" % (ref_type, local_id))
        anchors[key] = context

    all_blocks = list(doc.body)
    for appx in doc.appendices:
        all_blocks.extend(appx.blocks)

    for blk in all_blocks:
        if isinstance(blk, model.TableModel):
            add_anchor("tbl", blk.anchor_id, "表题")
        elif isinstance(blk, model.Figure):
            add_anchor("fig", blk.anchor_id, "图题")
        elif isinstance(blk, model.Formula):
            add_anchor("eq", blk.anchor_id, "公式")

    normrefs = _collect_normative_refs(doc)

    for blk in all_blocks:
        for sp in _iter_block_spans(blk):
            if not isinstance(sp, model.RefSpan):
                continue
            if sp.ref_type == "std":
                if sp.target not in normrefs:
                    raise ValueError("未知规范性引用文件：%s。请确认其已在“规范性引用文件”章中列出。" %
                                     sp.target)
                continue
            key = (sp.ref_type, sp.target)
            if key not in anchors:
                raise ValueError("未知交叉引用：{{%s:%s}}。请确认对应锚点存在。" %
                                 (sp.ref_type, sp.target))


def _route_reference(doc, kind, blk):
    if kind == "para":
        txt = _spans_text(blk[1]).strip()
        if txt:
            doc.references.append(txt)
    elif kind == "list":
        for it in blk[2]:
            txt = _spans_text(it).strip()
            if txt:
                doc.references.append(txt)


def _parse_index_item(text: str):
    m = _INDEX_ITEM_RE.match(text)
    if not m:
        raise ValueError("索引项必须写成 `术语：位置列表`：%s。" % text)
    return model.IndexItem(term=m.group(1).strip(), targets=m.group(2).strip())


def _route_index(doc, cur_group, kind, blk):
    if kind not in ("para", "list"):
        return
    if cur_group is None:
        raise ValueError("索引项必须放在 `## A` 到 `## Z` 分组下。")
    if kind == "para":
        txt = _spans_text(blk[1]).strip()
        if txt:
            cur_group.items.append(_parse_index_item(txt))
    elif kind == "list":
        for it in blk[2]:
            txt = _spans_text(it).strip()
            if txt:
                cur_group.items.append(_parse_index_item(txt))
