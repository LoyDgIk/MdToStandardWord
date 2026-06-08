# -*- coding: utf-8 -*-
"""Build-scoped counters used by DOCX generation."""

from __future__ import annotations


class _Counter:
    def __init__(self):
        self.table = 0
        self.figure = 0
        self.bm = 1000
        self.seq_scope_counts = {}


_COUNTER = _Counter()


def _reset_counters():
    _COUNTER.table = 0
    _COUNTER.figure = 0
    _COUNTER.bm = 1000
    _COUNTER.seq_scope_counts = {}


def _next_bm_id():
    bid = _COUNTER.bm
    _COUNTER.bm += 1
    return bid


def _needs_seq_reset(ref_type: str, appendix_letter=None) -> bool:
    scope = appendix_letter or "body"
    key = (ref_type, scope)
    count = _COUNTER.seq_scope_counts.get(key, 0)
    _COUNTER.seq_scope_counts[key] = count + 1
    return count == 0
