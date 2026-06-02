# -*- coding: utf-8 -*-
"""把 Markdown（YAML front matter + 正文）解析为 model.StandardDoc。

解析分两步：
1. 用 markdown-it-py 把正文 tokenize，归并成"原始块"列表（标题/段/列表/表/图）。
2. 路由：按 H1 把内容分流到 正文 / 附录 / 参考文献 / 索引，并识别 注、示例、图、表标题。
"""

from __future__ import annotations

import re
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
        pos = 0
        for m in _INLINE_REF_TOKEN_RE.finditer(sp.text):
            if m.start() > pos:
                chunk = sp.text[pos:m.start()]
                if "{{" in chunk or "}}" in chunk or "{@" in chunk:
                    raise ValueError("无效交叉引用语法：%s。" % chunk)
                out.append(model.Span(chunk, sp.bold, sp.italic))
            out.append(_parse_ref_token(m.group(0), sp.bold, sp.italic))
            pos = m.end()
        if pos < len(sp.text):
            tail = sp.text[pos:]
            if "{{" in tail or "}}" in tail or "{@" in tail:
                raise ValueError("无效交叉引用语法：%s。" % tail)
            out.append(model.Span(tail, sp.bold, sp.italic))
    return out


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
        extra_notes=[str(x) for x in (fw_raw.get("extra_notes") or [])],
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
        ics=s("ics"),
        ccs=s("ccs"),
        publish_date=s("publish_date"),
        implement_date=s("implement_date"),
        publisher=s("publisher"),
        foreword=fw,
        introduction=s("introduction"),
    )


# --------------------------------------------------------------------------- #
# 行内 -> spans
# --------------------------------------------------------------------------- #
def _inline_to_spans(inline_token) -> List[model.Span]:
    spans: List[model.Span] = []
    bold = 0
    italic = 0
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
        elif t in ("text", "code_inline"):
            if ch.content:
                spans.append(model.Span(ch.content, bold=bold > 0, italic=italic > 0))
        elif t == "softbreak":
            # 段内软换行：中文文本直接连接
            pass
        elif t == "hardbreak":
            spans.append(model.Span("\n", bold=bold > 0, italic=italic > 0))
        # image / link 文本由其它路径处理
    # 合并相邻同格式片段
    merged: List[model.Span] = []
    for sp in spans:
        if merged and merged[-1].bold == sp.bold and merged[-1].italic == sp.italic:
            merged[-1] = model.Span(merged[-1].text + sp.text, sp.bold, sp.italic)
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
#   ('table', header:List[str], rows:List[List[str]])
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
            j, items = _parse_list(tokens, i, level)
            blocks.append(("list", ordered, items, level))
            i = j
        elif tt == "table_open":
            j, header, rows = _parse_table(tokens, i)
            blocks.append(("table", header, rows))
            i = j
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
    """解析 list，返回 (下一个索引, items)。items 为 List[List[Span]]。

    简化处理：列项内的首段作为列项文本；列项内嵌套列表作为更深一级（level+1）
    追加为后续 ('list', ...) 由调用者处理——此处把嵌套列表展开为额外 items 携带 level。
    """
    items: List[List[model.Span]] = []
    i = start + 1  # skip *_list_open
    n = len(tokens)
    depth = 1
    nested_blocks = []
    while i < n and depth > 0:
        tt = tokens[i].type
        if tt in ("bullet_list_open", "ordered_list_open"):
            depth += 1
            i += 1
            continue
        if tt in ("bullet_list_close", "ordered_list_close"):
            depth -= 1
            i += 1
            continue
        if tt == "list_item_open" and depth == 1:
            # 收集该列项内首个 inline 段
            spans: List[model.Span] = []
            k = i + 1
            iddepth = 1
            captured = False
            while k < n and iddepth > 0:
                if tokens[k].type == "list_item_open":
                    iddepth += 1
                elif tokens[k].type == "list_item_close":
                    iddepth -= 1
                elif tokens[k].type == "inline" and not captured and iddepth == 1:
                    spans = _inline_to_spans(tokens[k])
                    captured = True
                k += 1
            items.append(spans)
            i = k
            continue
        i += 1
    return i, items


def _parse_table(tokens, start):
    header: List[str] = []
    rows: List[List[str]] = []
    i = start + 1
    n = len(tokens)
    cur: List[str] = []
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
        elif tt == "tr_close":
            if in_head:
                header = cur
            else:
                rows.append(cur)
        elif tt in ("th_open", "td_open"):
            inline = tokens[i + 1] if i + 1 < n and tokens[i + 1].type == "inline" else None
            txt = "".join(s.text for s in _inline_to_spans(inline)) if inline is not None else ""
            cur.append(txt)
        i += 1
    return i, header, rows


# --------------------------------------------------------------------------- #
# 原始块 -> StandardDoc（路由）
# --------------------------------------------------------------------------- #
_NOTE_RE = re.compile(r"^\s*注\s*(\d+)?\s*[:：]")
_EXAMPLE_RE = re.compile(r"^\s*示例\s*(\d+)?\s*[:：]")
_SOURCE_RE = re.compile(r"^\s*\[?\s*来源\s*[:：]")
_TABLE_CAP_RE = re.compile(r"^\s*\{\s*表\s*[:：]\s*#([^}\s]+)\s*\}\s+(.+?)\s*$")
_TABLE_CAP_MARKER_RE = re.compile(r"^\s*(?:\{\s*)?表\s*[:：]")
# 无标题条：以多段数字编号开头，后接内容，如 "4.2.1  正文……"
_UNTITLED_RE = re.compile(r"^\s*(\d+(?:\.\d+)+)\s+(\S.*)$", re.S)
_APPENDIX_RE = re.compile(r"^\s*附录\s*(?:[A-ZＡ-Ｚ]\s*)?[（(]?\s*(规范性|资料性)\s*[)）]?\s*(.*)$")
_INDEX_GROUP_RE = re.compile(r"^[A-Z]$")
_INDEX_ITEM_RE = re.compile(r"^\s*(.+?)\s*[:：]\s*(.+?)\s*$")


def _spans_text(spans):
    return "".join(s.text for s in spans)


def _example_has_inline_content(example: model.Example) -> bool:
    text = _spans_text(example.spans)
    return _EXAMPLE_RE.sub("", text, count=1).strip() != ""


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

    idx = 0
    while idx < len(blocks):
        blk = blocks[idx]
        kind = blk[0]

        # --- 标题：可能切换 mode ---
        if kind == "heading":
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
                idx += 1
                continue

        # --- 表格标题（"{表：#tbl:id} 标题"，后接表格）---
        if kind == "para":
            ptext = _spans_text(blk[1]).strip()
            if idx + 1 < len(blocks) and blocks[idx + 1][0] == "table":
                m_cap = _TABLE_CAP_RE.match(ptext)
                if m_cap:
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
            idx += 1
            continue

        if expect_example_content and isinstance(target_block, model.Paragraph):
            target_block = model.ExampleContent(spans=target_block.spans)
        expect_example_content = (
            isinstance(target_block, model.Example)
            and not _example_has_inline_content(target_block)
        )

        if mode == "body":
            doc.body.append(target_block)
        elif mode == "appendix" and cur_appendix is not None:
            cur_appendix.blocks.append(target_block)
        elif mode == "references":
            _route_reference(doc, kind, blk)
        elif mode == "index":
            _route_index(doc, cur_index_group, kind, blk)
        idx += 1

    _validate_refs(doc)
    return doc


def _make_block(kind, blk, table_caption):
    if kind == "para":
        spans = blk[1]
        ptext = _spans_text(spans)
        m = _NOTE_RE.match(ptext)
        if m:
            return model.Note(spans=spans, index=int(m.group(1)) if m.group(1) else None)
        m = _EXAMPLE_RE.match(ptext)
        if m:
            return model.Example(spans=spans, index=int(m.group(1)) if m.group(1) else None)
        if _SOURCE_RE.match(ptext):
            return model.Source(text=ptext.strip())
        m = _UNTITLED_RE.match(ptext)
        if m:
            number = m.group(1)
            rest_spans = _strip_leading(spans, len(ptext) - len(m.group(2)))
            return model.UntitledClause(number=number, spans=rest_spans)
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
        if table_caption is None:
            cap, anchor = "", ""
        else:
            cap, anchor = table_caption
        return model.TableModel(caption=cap, anchor_id=anchor, header=header, rows=rows)
    return None


def _strip_leading(spans, n):
    """从 spans 序列前部去掉 n 个字符（用于剥离无标题条编号前缀）。"""
    out = []
    remaining = n
    for sp in spans:
        if remaining <= 0:
            out.append(sp)
            continue
        if len(sp.text) <= remaining:
            remaining -= len(sp.text)
            continue
        out.append(model.Span(sp.text[remaining:], sp.bold, sp.italic))
        remaining = 0
    return out


def _iter_block_spans(blk):
    if isinstance(blk, (model.Heading, model.Paragraph, model.UntitledClause,
                        model.Note, model.Example, model.ExampleContent)):
        yield from blk.spans
    elif isinstance(blk, model.ListBlock):
        for item in blk.items:
            yield from item.spans


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
