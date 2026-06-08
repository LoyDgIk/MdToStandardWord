# -*- coding: utf-8 -*-
"""Cover blueprint selection, metadata filling, and cover protection helpers."""

from __future__ import annotations

import os
import re
import shutil
import zipfile
from typing import Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .. import model
from .. import resources
from .. import styles as S
from .oxml import (
    _carries_sectpr,
    _first_existing_path,
    _new_paragraph_before,
    _replace_text_keep_format,
)

def _resolve_kind(kind: str, meta: model.Meta) -> str:
    if kind not in ("auto", "group", "national"):
        raise ValueError("kind 只能是 auto、group 或 national。")
    if kind != "auto":
        return kind
    standard_type = (meta.standard_type or "").strip()
    number = (meta.number or "").strip().upper()
    if "国家" in standard_type or number.startswith("GB"):
        return "national"
    return "group"


def _default_cover_path(kind: str) -> str:
    if kind == "national":
        return _first_existing_path(
            resources.template_candidates("cover_national.docx"),
            "国家标准封面蓝图",
        )
    return _first_existing_path(
        resources.template_candidates("cover_group.docx"),
        "团体标准封面蓝图",
    )


def _copy_cover_base(cover_path: str, output_path: str):
    shutil.copyfile(cover_path, output_path)


def _read_cover_end_line_image(docx_path: str, kind: str) -> bytes:
    """读取封面蓝图包内自带的标准结束线图片。"""
    if kind == "national":
        candidates = [
            "word/media/image3.jpg",
            "word/media/image3.jpeg",
        ]
    else:
        candidates = [
            "word/media/image1.jpeg",
            "word/media/image1.jpg",
        ]
    with zipfile.ZipFile(docx_path, "r") as zf:
        for name in candidates:
            try:
                return zf.read(name)
            except KeyError:
                continue
    raise FileNotFoundError("封面蓝图缺少标准结束线图片：%s" % "；".join(candidates))
def _set_field(doc, style_name: str, text: str) -> bool:
    """把第一个套用 style_name 的段落文本替换为 text，保留首个 run 的格式。返回是否命中。"""
    for p in doc.paragraphs:
        if p.style is not None and p.style.name == style_name:
            _replace_text_keep_format(p, text)
            return True
    return False


def _set_field_or_placeholder(doc, style_name: str, text: str, patterns: List[str]) -> bool:
    if _set_field(doc, style_name, text):
        return True
    for p in doc.paragraphs:
        current = p.text.strip()
        if any(re.search(pattern, current) for pattern in patterns):
            _replace_text_keep_format(p, text)
            return True
    return False


def _format_parenthesized(text: str) -> str:
    value = (text or "").strip()
    if not value:
        return ""
    if value.startswith(("(", "（")) and value.endswith((")", "）")):
        return value
    return "（%s）" % value


def _set_record_number(doc, text: str) -> bool:
    value = (text or "").strip()
    if not value:
        return False
    if not value.startswith("备案号"):
        value = "备案号：%s" % value
    for p in doc.paragraphs:
        if "备案号" in (p.text or ""):
            _replace_text_keep_format(p, value)
            return True
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "备案号" in (p.text or ""):
                        _replace_text_keep_format(p, value)
                        return True
    return False


def _set_section_form_protection(sectpr, protected: bool):
    form_prot = sectpr.find(qn("w:formProt"))
    if protected:
        if form_prot is not None:
            sectpr.remove(form_prot)
        return
    if form_prot is None:
        form_prot = OxmlElement("w:formProt")
        sectpr.append(form_prot)
    form_prot.set(qn("w:val"), "0")


def _enable_cover_form_field_protection(doc):
    """Activate legacy form fields on the cover while keeping later sections editable."""
    settings = doc.settings.element
    protection = settings.find(qn("w:documentProtection"))
    if protection is None:
        protection = OxmlElement("w:documentProtection")
        settings.append(protection)
    for attr in list(protection.attrib):
        del protection.attrib[attr]
    protection.set(qn("w:edit"), "forms")
    protection.set(qn("w:enforcement"), "1")

    for idx, section in enumerate(doc.sections):
        _set_section_form_protection(section._sectPr, protected=(idx == 0))


def _disable_form_field_protection(doc):
    settings = doc.settings.element
    protection = settings.find(qn("w:documentProtection"))
    if protection is not None:
        for attr in list(protection.attrib):
            del protection.attrib[attr]
        protection.set(qn("w:edit"), "forms")
        protection.set(qn("w:enforcement"), "0")
    for section in doc.sections:
        _set_section_form_protection(section._sectPr, protected=False)


def _should_enable_cover_form_protection(
    meta: model.Meta,
    cover_form_protection: Optional[bool],
) -> bool:
    if cover_form_protection is not None:
        return bool(cover_form_protection)
    return bool(getattr(meta, "cover_form_protection", False))


def _legacy_dropdown_field_info(begin_run_el):
    fld = begin_run_el.find(qn("w:fldChar"))
    if fld is None:
        return None
    ffdata = fld.find(qn("w:ffData"))
    if ffdata is None:
        return None
    name_el = ffdata.find(qn("w:name"))
    ddlist = ffdata.find(qn("w:ddList"))
    if name_el is None or ddlist is None:
        return None
    return name_el.get(qn("w:val"), ""), ddlist


def _normalize_draft_version(value: str) -> str:
    text = (value or "").strip()
    aliases = {
        "工作组讨论稿": "（工作组讨论稿）",
        "征求意见稿": "（征求意见稿）",
        "送审讨论稿": "（送审讨论稿）",
        "送审稿": "（送审稿）",
        "报批稿": "（报批稿）",
    }
    return aliases.get(text, text)


def _set_legacy_dropdown_value(doc, field_name: str, value: str) -> bool:
    selected = _normalize_draft_version(value)
    if not selected:
        return False
    for run_el in doc.element.iter(qn("w:r")):
        fld = run_el.find(qn("w:fldChar"))
        if fld is None or fld.get(qn("w:fldCharType")) != "begin":
            continue
        info = _legacy_dropdown_field_info(run_el)
        if info is None:
            continue
        name, ddlist = info
        if name != field_name:
            continue

        entries = ddlist.findall(qn("w:listEntry"))
        values = [entry.get(qn("w:val"), "") for entry in entries]
        if selected not in values:
            entry = OxmlElement("w:listEntry")
            entry.set(qn("w:val"), selected)
            ddlist.append(entry)
            values.append(selected)
        idx = values.index(selected)
        result = ddlist.find(qn("w:result"))
        if result is None:
            result = OxmlElement("w:result")
            ddlist.insert(0, result)
        result.set(qn("w:val"), str(idx))
        return True
    return False


# --------------------------------------------------------------------------- #
# 各节构建
# --------------------------------------------------------------------------- #
def _apply_cover_fields(doc, meta: model.Meta, kind: Optional[str] = None):
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

    standard_type = meta.standard_type
    if kind == "national" and standard_type in ("", "国家标准", "中华人民共和国国家标准"):
        standard_type = "中华人民共和国国家标准"
    if standard_type:
        _set_field_or_placeholder(doc, S.S_COVER_TYPE, standard_type, [
            r"^团体标准$",
            r"^中华人民共和国国家标准$",
        ])
    if meta.number:
        _set_field_or_placeholder(doc, S.S_COVER_NUMBER, meta.number, [
            r"^T/XXX\s+XXXX",
            r"^GB/T\s+XXXXX",
            r"^GB\s+XXXXX",
        ])
    if meta.replaces:
        _set_field_or_placeholder(doc, S.S_COVER_REPLACES, "代替 %s" % meta.replaces, [
            r"^代替\s+",
        ])
    if meta.title:
        _set_field_or_placeholder(doc, S.S_COVER_NAME, meta.title, [
            r"^点击此处添加标准名称$",
        ])
    if meta.title_en:
        _set_field_or_placeholder(doc, S.S_COVER_NAME_EN, meta.title_en, [
            r"^点击此处添加标准名称的英文译名$",
        ])
    if meta.consistency_degree:
        _set_field_or_placeholder(
            doc,
            S.S_COVER_CONSISTENCY,
            _format_parenthesized(meta.consistency_degree),
            [
                r"一致性程度",
                r"一致程度",
                r"点击此处添加与国际标准一致性程度的标识",
            ],
        )
    if meta.draft_version:
        _set_legacy_dropdown_value(doc, "下拉1", meta.draft_version)
    if meta.record_number:
        _set_record_number(doc, meta.record_number)
    if meta.publish_date:
        _set_field_or_placeholder(doc, S.S_COVER_PUBLISH, "%s发布" % meta.publish_date, [
            r"^XXXX\s*-\s*XX\s*-\s*XX发布$",
        ])
    if meta.implement_date:
        _set_field_or_placeholder(doc, S.S_COVER_IMPLEMENT, "%s实施" % meta.implement_date, [
            r"^XXXX\s*-\s*XX\s*-\s*XX实施$",
        ])
    publisher_set = False
    if meta.publisher:
        publisher_set = _set_field_or_placeholder(doc, S.S_COVER_PUBLISHER, meta.publisher, [])
    return {"publisher": publisher_set}


_COVER_PLACEHOLDER_RE = re.compile(
    r"(\(?点击此处添加[^ \t\r\n<]*\)?)|"
    r"(点击此处添加[^ \t\r\n<]*)|"
    r"（本草案完成时间：）|"
    r"XXXX\s*-\s*XX\s*-\s*XX(?:发布|实施)?"
)


def _cleanup_placeholder_paragraph(p: Paragraph):
    text = p.text
    if not text or not _COVER_PLACEHOLDER_RE.search(text):
        return
    cleaned = _COVER_PLACEHOLDER_RE.sub("", text).strip()
    _replace_text_keep_format(p, cleaned)


def _cleanup_cover_placeholders(doc):
    for p in doc.paragraphs:
        _cleanup_placeholder_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _cleanup_placeholder_paragraph(p)


def _find_first_section_break_paragraph(doc) -> Optional[Paragraph]:
    for p in doc.paragraphs:
        if _carries_sectpr(p):
            return p
    return None


def _ensure_cover_publisher(doc, publisher: str, cover_info: Optional[dict], kind: str = "group"):
    if not publisher or (cover_info or {}).get("publisher"):
        return
    # 国家标准封面蓝图底部发布单位为打包图片；再插入文本段会与图片重叠。
    if kind == "national":
        return
    anchor = _find_first_section_break_paragraph(doc)
    if anchor is None:
        para = doc.add_paragraph()
    else:
        para = _new_paragraph_before(anchor, doc, S.S_COVER_PUBLISHER)
    if not para.runs:
        para.add_run(publisher)
    else:
        _replace_text_keep_format(para, publisher)


def _set_cell_text(cell, text):
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for extra in p.runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        p.add_run(text)


__all__ = [name for name in globals() if name.startswith("_")]
