# -*- coding: utf-8 -*-
"""Compatibility facade for DOCX generation.

The implementation lives in :mod:`md2std.docx`; this module keeps the historical
``md2std.docx_builder.build_cover`` import path stable.
"""

from __future__ import annotations

from .docx.builder import build_cover
from .docx.cover import _default_cover_path
from .docx.oxml import _native_ref_name

__all__ = ["build_cover"]
