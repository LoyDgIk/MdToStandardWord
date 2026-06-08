# -*- coding: utf-8 -*-
"""High-level DOCX build entrypoint."""

from __future__ import annotations

from typing import Optional

from docx import Document

from .. import model
from .content import _emit_cover_sections
from .cover import (
    _apply_cover_fields,
    _cleanup_cover_placeholders,
    _copy_cover_base,
    _default_cover_path,
    _disable_form_field_protection,
    _ensure_cover_publisher,
    _read_cover_end_line_image,
    _resolve_kind,
    _should_enable_cover_form_protection,
    _enable_cover_form_field_protection,
)
from .oxml import (
    _configure_standard_styles,
    _enable_update_fields,
    _normalize_cover_page_number_for_odd_even_export,
    _set_even_and_odd_headers,
)
from .state import _reset_counters


def build_cover(
    sdoc: model.StandardDoc,
    output_path: str,
    kind: str = "auto",
    cover_form_protection: Optional[bool] = None,
):
    """Build a standard DOCX from a cover blueprint and generated content."""
    _reset_counters()

    resolved_kind = _resolve_kind(kind, sdoc.meta)
    cover_path = _default_cover_path(resolved_kind)

    _copy_cover_base(cover_path, output_path)
    end_line_image = _read_cover_end_line_image(output_path, resolved_kind)

    doc = Document(output_path)
    _configure_standard_styles(doc)

    cover_info = _apply_cover_fields(doc, sdoc.meta, kind=resolved_kind)
    _ensure_cover_publisher(doc, sdoc.meta.publisher, cover_info, kind=resolved_kind)
    _cleanup_cover_placeholders(doc)

    _emit_cover_sections(doc, sdoc, end_line_image, resolved_kind)
    _enable_update_fields(doc)
    _set_even_and_odd_headers(doc, sdoc.meta.odd_even_pages)
    _normalize_cover_page_number_for_odd_even_export(doc, sdoc.meta.odd_even_pages)
    if _should_enable_cover_form_protection(sdoc.meta, cover_form_protection):
        _enable_cover_form_field_protection(doc)
    else:
        _disable_form_field_protection(doc)

    doc.save(output_path)
    return output_path
