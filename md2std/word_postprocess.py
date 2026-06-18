# -*- coding: utf-8 -*-
"""Microsoft Word COM 后处理。

该模块是可选依赖层：普通生成不导入 win32com；只有 CLI 显式启用
`--word-com-postprocess` 时才调用。后处理职责是让 Word 自身完成
域更新、重新分页，并按真实分页位置拆分跨页表格、插入续表题。
"""

from __future__ import annotations

import copy
import os
import re
import shutil
import subprocess
import tempfile
import threading
import zipfile
from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple

from lxml import etree as ET

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_NS = "http://www.w3.org/XML/1998/namespace"
_WD_ACTIVE_END_PAGE_NUMBER = 3
_CONTINUATION_STYLE_NAME = "标准文件_表格续"
_CONTINUATION_SUFFIX = "（续）"
_WORD_COM_TIMEOUT_SECONDS = 180
_MIN_BODY_ROWS_BEFORE_CONTINUATION = 1
_CAPTION_LOOKBACK_PARAGRAPHS = 6
_HORIZONTAL_SPLIT_REPEAT_COLUMNS = 1
_HORIZONTAL_SPLIT_MIN_COLUMNS = 4
_HORIZONTAL_MIN_READABLE_COLUMN_WIDTH = 1000
_CAPTION_SPACING_BEFORE_AFTER = 50
_CAPTION_LINE_SPACING = 400


@dataclass
class _TableContinuationPlan:
    """一个表格的真实跨页拆分计划。row_breaks 为 Word 1-based 行号。"""

    table_index: int
    row_breaks: List[int]
    header_count: int
    caption_text: str


@dataclass
class _HorizontalTableSplitPlan:
    """一个表格的横向拆分计划。列分组由 OOXML 实际宽度计算。"""

    table_index: int
    caption_text: str


def postprocess_with_word_com(path: str, visible: bool = False,
                              timeout_seconds: int = _WORD_COM_TIMEOUT_SECONDS) -> str:
    """使用本机 Microsoft Word COM 更新域、重新分页并保存 DOCX。"""
    abs_path = os.path.abspath(path)
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(abs_path)

    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "启用 --word-com-postprocess 需要安装 pywin32，并且只能在装有 Microsoft Word 的 Windows 环境使用。"
        ) from exc

    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = bool(visible)
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            pass
        _postprocess_with_word_instance(abs_path, word, timeout_seconds=timeout_seconds)
        word = None
    finally:
        if word is not None:
            _quit_word(word)
    return abs_path


def _postprocess_with_word_instance(abs_path: str, word, timeout_seconds: int):
    """在临时副本上运行 Word 后处理，成功后再替换原文件。"""
    work_path = _make_temp_docx_copy(abs_path)
    guard = _WordProcessGuard(_word_process_id(word), timeout_seconds)
    success = False
    try:
        guard.start()
        _postprocess_document(word, work_path)
        if guard.timed_out:
            raise RuntimeError("Word COM 后处理超时，已尝试清理本次创建的 Word 进程。")
        _quit_word(word)
        guard.cancel()
        os.replace(work_path, abs_path)
        success = True
    finally:
        guard.cancel()
        if not success and os.path.exists(work_path):
            try:
                os.remove(work_path)
            except OSError:
                pass


def _make_temp_docx_copy(abs_path: str) -> str:
    directory = os.path.dirname(abs_path) or os.getcwd()
    fd, work_path = tempfile.mkstemp(
        prefix=".md2std-wordcom-",
        suffix=".docx",
        dir=directory,
    )
    os.close(fd)
    shutil.copyfile(abs_path, work_path)
    return work_path


class _WordProcessGuard:
    """只监控并清理本次 DispatchEx 创建的 Word 进程。"""

    def __init__(self, pid, timeout_seconds: int):
        self.pid = pid
        self.timeout_seconds = timeout_seconds
        self.timed_out = False
        self._timer = None

    def start(self):
        if not self.pid or not self.timeout_seconds or self.timeout_seconds <= 0:
            return
        self._timer = threading.Timer(self.timeout_seconds, self._on_timeout)
        self._timer.daemon = True
        self._timer.start()

    def cancel(self):
        if self._timer is not None:
            self._timer.cancel()
            self._timer = None

    def _on_timeout(self):
        self.timed_out = True
        _kill_process_tree(self.pid)


def _word_process_id(word):
    try:
        import win32process  # type: ignore
        hwnd = int(word.Hwnd)
        if hwnd:
            return win32process.GetWindowThreadProcessId(hwnd)[1]
    except Exception:
        return None
    return None


def _kill_process_tree(pid):
    if not pid:
        return
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
    except Exception:
        pass


def _quit_word(word):
    try:
        word.Quit()
    except Exception:
        pass


def _postprocess_document(word, abs_path: str):
    """多轮后处理：测量真实分页 -> 拆续表 -> 重新分页，直到不再需要拆分。"""
    max_iterations = 20
    applied_signatures = set()
    for _ in range(max_iterations):
        plans, horizontal_plans = _measure_table_layout_plans(word, abs_path)
        if not plans:
            if horizontal_plans and _apply_horizontal_table_splits(abs_path, horizontal_plans):
                continue
            return
        plan = plans[0]
        signature = _table_continuation_plan_signature(plan)
        if signature in applied_signatures:
            return
        applied_signatures.add(signature)
        # 续表题与后续表格保持同页，拆表仍可能改变后续表的分页位置。
        # 因此每轮只处理文档中最早的一个计划，随后交给 Word 重新分页。
        if not _apply_table_continuations(abs_path, [plan]):
            return
    # 最后一轮保存一次 Word 分页结果；若仍有极端超高行，避免无限循环。
    _measure_table_layout_plans(word, abs_path)


def _measure_table_continuations(word, abs_path: str) -> List[_TableContinuationPlan]:
    """让 Word 完成分页后，收集真实跨页表格的断开行。"""
    plans, _horizontal_plans = _measure_table_layout_plans(word, abs_path)
    return plans


def _measure_table_layout_plans(word, abs_path: str) -> Tuple[List[_TableContinuationPlan], List[_HorizontalTableSplitPlan]]:
    """让 Word 完成分页后，收集跨页续表和横向拆表所需的真实表题。"""
    doc = None
    try:
        doc = _open_document(word, abs_path)
        _update_all_fields(doc)
        _disable_row_page_breaks(doc)
        try:
            doc.Repaginate()
        except Exception:
            pass
        if _allow_self_spanning_rows(doc):
            try:
                doc.Repaginate()
            except Exception:
                pass
        plans = _collect_table_continuation_plans(doc)
        horizontal_plans = _collect_horizontal_table_split_plans(doc)
        doc.Save()
        return plans, horizontal_plans
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass


def _open_document(word, abs_path: str):
    try:
        return word.Documents.Open(
            FileName=abs_path,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
            Revert=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
    except Exception:
        return word.Documents.Open(abs_path)


def _disable_row_page_breaks(doc):
    """尽量避免单行自身跨页，否则无法在行中插入续表题。"""
    try:
        count = doc.Tables.Count
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            doc.Tables(i).Rows.AllowBreakAcrossPages = False
        except Exception:
            pass


def _allow_self_spanning_rows(doc) -> bool:
    """超高行只能交给 Word 在行内自然分页，不能尝试插入续表题。"""
    changed = False
    try:
        table_count = doc.Tables.Count
    except Exception:
        return False
    for table_index in range(1, table_count + 1):
        try:
            table = doc.Tables(table_index)
            row_count = table.Rows.Count
        except Exception:
            continue
        row_spans = _table_row_page_spans_from_cells(doc, table)
        if not row_spans:
            continue
        for row_index in _self_spanning_row_indices(row_spans):
            if row_index < 1 or row_index > row_count:
                continue
            try:
                row = table.Rows(row_index)
                if not bool(row.AllowBreakAcrossPages):
                    row.AllowBreakAcrossPages = True
                    changed = True
            except Exception:
                continue
    return changed


def _self_spanning_row_indices(row_spans) -> List[int]:
    return [
        row_index
        for row_index, start_page, end_page in row_spans
        if isinstance(start_page, int) and isinstance(end_page, int) and end_page != start_page
    ]


def _collect_table_continuation_plans(doc) -> List[_TableContinuationPlan]:
    plans: List[_TableContinuationPlan] = []
    try:
        table_count = doc.Tables.Count
    except Exception:
        return plans
    for table_index in range(1, table_count + 1):
        try:
            table = doc.Tables(table_index)
            caption_text = _table_caption_text(doc, table)
        except Exception:
            continue
        if not _is_table_caption(caption_text):
            continue
        try:
            row_count = table.Rows.Count
        except Exception:
            continue
        header_count = _table_header_count(table, row_count)
        row_breaks = _table_row_page_breaks(doc, table, row_count, header_count)
        if not row_breaks:
            continue
        valid_breaks = _valid_continuation_breaks(row_breaks, header_count, row_count)
        if valid_breaks:
            plans.append(_TableContinuationPlan(
                table_index=table_index,
                row_breaks=valid_breaks,
                header_count=header_count,
                caption_text=_continuation_caption_text(caption_text),
            ))
    return _dedupe_table_continuation_plans(plans)


def _collect_horizontal_table_split_plans(doc) -> List[_HorizontalTableSplitPlan]:
    plans: List[_HorizontalTableSplitPlan] = []
    try:
        table_count = doc.Tables.Count
    except Exception:
        return plans
    for table_index in range(1, table_count + 1):
        try:
            table = doc.Tables(table_index)
            caption_text = _table_caption_text(doc, table)
        except Exception:
            continue
        if not _is_table_caption(caption_text):
            continue
        caption_text = _continuation_caption_text(caption_text)
        plans.append(_HorizontalTableSplitPlan(table_index=table_index, caption_text=caption_text))
    return plans


def _table_continuation_plan_signature(plan: _TableContinuationPlan):
    return (
        plan.table_index,
        tuple(plan.row_breaks),
        plan.header_count,
        plan.caption_text,
    )


def _dedupe_table_continuation_plans(plans: List[_TableContinuationPlan]) -> List[_TableContinuationPlan]:
    deduped: List[_TableContinuationPlan] = []
    seen = set()
    for plan in plans:
        signature = _table_continuation_plan_signature(plan)
        if signature in seen:
            continue
        seen.add(signature)
        deduped.append(plan)
    return deduped


def _table_caption_text(doc, table) -> str:
    before = doc.Range(0, table.Range.Start)
    try:
        para_count = before.Paragraphs.Count
    except Exception:
        para_count = 0
    fallback = ""
    for offset in range(0, min(_CAPTION_LOOKBACK_PARAGRAPHS, para_count)):
        try:
            para = before.Paragraphs.Item(para_count - offset)
            text = _clean_word_text(para.Range.Text)
            list_string = _clean_word_text(para.Range.ListFormat.ListString)
        except Exception:
            continue
        candidate = ((list_string + "　") if list_string else "") + text
        candidate = candidate.strip()
        if offset == 0:
            fallback = candidate
        if _is_table_caption(candidate):
            return candidate
    return fallback


def _clean_word_text(text: str) -> str:
    return (text or "").replace("\r", "").replace("\x07", "").strip()


def _is_table_caption(text: str) -> bool:
    # 仅处理由本项目生成的可见表题/续表题，不碰封面和普通排版表。
    return re.match(r"^表(?:[A-Z]\.)?\d+(?:\.\d+)?[\s　].+", text or "") is not None


def _continuation_caption_text(text: str) -> str:
    base = re.sub(r"(?:（续）|\(续\))\s*$", "", (text or "").strip())
    return base + _CONTINUATION_SUFFIX


def _table_header_count(table, row_count: int) -> int:
    count = 0
    row_access_failed = False
    for row_index in range(1, row_count + 1):
        try:
            is_header = bool(table.Rows(row_index).HeadingFormat)
        except Exception:
            row_access_failed = True
            break
        if not is_header:
            break
        count += 1
    if row_access_failed and count == 0:
        return _table_header_count_from_openxml(table)
    return count


def _table_row_page_breaks(doc, table, row_count: int, header_count: int) -> List[int]:
    row_spans = _table_row_page_spans_from_cells(doc, table)
    if row_spans:
        return _row_page_breaks(row_spans, header_count)
    row_spans = []
    for row_index in range(1, row_count + 1):
        try:
            row = table.Rows(row_index)
            start_page = doc.Range(row.Range.Start, row.Range.Start).Information(_WD_ACTIVE_END_PAGE_NUMBER)
            end_page = doc.Range(row.Range.End, row.Range.End).Information(_WD_ACTIVE_END_PAGE_NUMBER)
        except Exception:
            row_spans = _table_row_page_spans_from_cells(doc, table)
            break
        if not isinstance(start_page, int) or not isinstance(end_page, int):
            continue
        row_spans.append((row_index, start_page, end_page))
    return _row_page_breaks(row_spans, header_count)


def _row_page_breaks(row_spans, header_count: int) -> List[int]:
    breaks: List[int] = []
    previous_page = None
    for row_index, start_page, end_page in row_spans:
        if previous_page is not None and (start_page != previous_page or end_page != start_page):
            # Word 有时把整行视觉上推到下一页，但 row.Range.Start 仍报告上一页；
            # 此时用 End 页码把断点前移到真实换页行。
            if row_index > header_count:
                breaks.append(row_index)
        previous_page = end_page
    return breaks


def _table_row_page_spans_from_cells(doc, table):
    """Fallback for tables with vertical merges where Word disallows Rows(i)."""
    row_ranges = {}
    try:
        cells = table.Range.Cells
        count = cells.Count
    except Exception:
        return []
    for cell_index in range(1, count + 1):
        try:
            cell = cells.Item(cell_index)
            row_index = int(cell.RowIndex)
            start = int(cell.Range.Start)
            end = int(cell.Range.End)
        except Exception:
            continue
        if row_index not in row_ranges:
            row_ranges[row_index] = [start, end]
        else:
            row_ranges[row_index][0] = min(row_ranges[row_index][0], start)
            row_ranges[row_index][1] = max(row_ranges[row_index][1], end)

    row_spans = []
    for row_index in sorted(row_ranges):
        start, end = row_ranges[row_index]
        try:
            start_page = doc.Range(start, start).Information(_WD_ACTIVE_END_PAGE_NUMBER)
            end_page = doc.Range(end, end).Information(_WD_ACTIVE_END_PAGE_NUMBER)
        except Exception:
            continue
        if isinstance(start_page, int) and isinstance(end_page, int):
            row_spans.append((row_index, start_page, end_page))
    return row_spans


def _table_header_count_from_openxml(table) -> int:
    try:
        xml = table.Range.WordOpenXML
    except Exception:
        return 0
    if isinstance(xml, str):
        xml = xml.encode("utf-8")
    try:
        root = ET.fromstring(xml)
    except Exception:
        return 0
    table_el = root if root.tag == _w("tbl") else root.find(".//" + _w("tbl"))
    if table_el is None:
        return 0
    count = 0
    for row in table_el.findall(_w("tr")):
        header = row.find("./" + _w("trPr") + "/" + _w("tblHeader"))
        if header is None:
            break
        value = header.get(_w("val"))
        if value in ("0", "false", "False"):
            break
        count += 1
    return count


def _has_enough_body_rows_before_break(row_index: int, header_count: int,
                                       previous_break: Optional[int] = None) -> bool:
    body_start = max(header_count + 1, 1)
    segment_start = previous_break or body_start
    return row_index - segment_start >= _MIN_BODY_ROWS_BEFORE_CONTINUATION


def _apply_table_continuations(path: str, plans: List[_TableContinuationPlan]) -> bool:
    if not plans:
        return False
    with zipfile.ZipFile(path, "r") as zin:
        document_xml = zin.read("word/document.xml")
        styles_xml = zin.read("word/styles.xml")
        entries = [(item, zin.read(item.filename)) for item in zin.infolist()]

    root = ET.fromstring(document_xml)
    body = root.find(_w("body"))
    if body is None:
        return False
    style_id = _style_id_by_name(styles_xml, _CONTINUATION_STYLE_NAME)
    changed = _split_tables_for_continuations(body, plans, style_id)
    if not changed:
        return False

    new_document_xml = ET.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )
    fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item, data in entries:
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_document_xml)
                else:
                    zout.writestr(item, data)
        shutil.move(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    return True


def _apply_horizontal_table_splits(path: str, plans: List[_HorizontalTableSplitPlan]) -> bool:
    if not plans:
        return False
    with zipfile.ZipFile(path, "r") as zin:
        document_xml = zin.read("word/document.xml")
        styles_xml = zin.read("word/styles.xml")
        entries = [(item, zin.read(item.filename)) for item in zin.infolist()]

    root = ET.fromstring(document_xml)
    body = root.find(_w("body"))
    if body is None:
        return False
    available_width = _body_available_width_twips(body)
    if available_width <= 0:
        return False
    style_id = _style_id_by_name(styles_xml, _CONTINUATION_STYLE_NAME)
    changed = _split_tables_horizontally(body, plans, style_id, available_width)
    if not changed:
        return False

    new_document_xml = ET.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )
    fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item, data in entries:
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_document_xml)
                else:
                    zout.writestr(item, data)
        shutil.move(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    return True


def _split_tables_for_continuations(body, plans: List[_TableContinuationPlan], style_id: str) -> bool:
    tables = [el for el in list(body) if el.tag == _w("tbl")]
    changed = False
    for plan in sorted(plans, key=lambda p: p.table_index, reverse=True):
        if plan.table_index < 1 or plan.table_index > len(tables):
            continue
        table = tables[plan.table_index - 1]
        replacement = _split_table_element(table, plan, style_id)
        if not replacement:
            continue
        pos = list(body).index(table)
        body.remove(table)
        for offset, element in enumerate(replacement):
            body.insert(pos + offset, element)
        changed = True
    return changed


def _split_tables_horizontally(body, plans: List[_HorizontalTableSplitPlan],
                               style_id: str, available_width: int) -> bool:
    tables = [el for el in list(body) if el.tag == _w("tbl")]
    plan_by_index = {plan.table_index: plan for plan in plans}
    changed = False
    for table_index in range(len(tables), 0, -1):
        plan = plan_by_index.get(table_index)
        if plan is None:
            continue
        table = tables[table_index - 1]
        replacement = _split_table_horizontally(table, plan, style_id, available_width)
        if not replacement:
            continue
        pos = list(body).index(table)
        body.remove(table)
        for offset, element in enumerate(replacement):
            body.insert(pos + offset, element)
        changed = True
    return changed


def _split_table_horizontally(table, plan: _HorizontalTableSplitPlan,
                              style_id: str, available_width: int) -> List[ET.Element]:
    widths = _table_grid_widths(table)
    if len(widths) < _HORIZONTAL_SPLIT_MIN_COLUMNS:
        return []
    grouping_widths = _horizontal_grouping_widths(widths, available_width)
    if sum(widths) <= available_width and grouping_widths == widths:
        return []
    if _table_has_unsupported_horizontal_spans(table, len(widths)):
        return []
    groups = _horizontal_column_groups(grouping_widths, available_width)
    if len(groups) <= 1:
        return []

    trailing_full_span_start = _trailing_full_span_row_start(table, len(widths))
    elements: List[ET.Element] = []
    for segment_index, columns in enumerate(groups):
        is_last_segment = segment_index == len(groups) - 1
        if segment_index > 0:
            elements.append(_make_continuation_caption(table, plan.caption_text, style_id))
        elements.append(_make_table_with_columns(
            table,
            columns,
            [grouping_widths[i] for i in columns],
            trailing_full_span_start=trailing_full_span_start,
            include_trailing_full_span=is_last_segment,
        ))
    return elements


def _horizontal_grouping_widths(widths: Sequence[int], available_width: int) -> List[int]:
    widths = list(widths)
    if not widths or available_width <= 0:
        return widths
    if sum(widths) > available_width:
        return widths
    readable_widths = [max(width, _HORIZONTAL_MIN_READABLE_COLUMN_WIDTH) for width in widths]
    if sum(readable_widths) <= available_width:
        return widths
    return readable_widths


def _horizontal_column_groups(widths: Sequence[int], available_width: int) -> List[List[int]]:
    if not widths or available_width <= 0:
        return []
    groups: List[List[int]] = []
    repeat = min(_HORIZONTAL_SPLIT_REPEAT_COLUMNS, max(0, len(widths) - 1))
    start = 0
    while start < len(widths):
        columns: List[int] = []
        width = 0
        if groups and repeat:
            columns.extend(range(repeat))
            width = sum(widths[:repeat])
        added = False
        while start < len(widths):
            if start in columns:
                start += 1
                continue
            next_width = widths[start]
            if added and width + next_width > available_width:
                break
            columns.append(start)
            width += next_width
            start += 1
            added = True
            if width >= available_width:
                break
        if not added:
            return []
        groups.append(columns)
    return groups


def _split_table_element(table, plan: _TableContinuationPlan, style_id: str) -> List[ET.Element]:
    children = list(table)
    rows = [child for child in children if child.tag == _w("tr")]
    if len(rows) < 2:
        return []

    valid_breaks = _valid_continuation_breaks(
        _safe_vmerge_breaks(rows, plan.row_breaks),
        plan.header_count,
        len(rows),
    )
    if not valid_breaks:
        return []
    # 只应用当前分页测量得到的第一个断点。插入续表题和拆表后，
    # 旧的后续断点可能已经失效，需交给下一轮 Word 重新分页。
    valid_breaks = valid_breaks[:1]

    start_rows = [1] + valid_breaks
    end_rows = valid_breaks + [len(rows) + 1]
    header_count = min(max(plan.header_count, 0), len(rows))
    header_rows = rows[:header_count]

    elements: List[ET.Element] = []
    for segment_index, (start, end) in enumerate(zip(start_rows, end_rows)):
        segment_rows = rows[start - 1:end - 1]
        if not segment_rows:
            continue
        if segment_index > 0 and header_rows:
            segment_rows = header_rows + segment_rows
        if segment_index > 0:
            elements.append(_make_continuation_caption(table, plan.caption_text, style_id))
        elements.append(_make_table_with_rows(table, segment_rows))
    return elements if len(elements) > 1 else []


def _body_available_width_twips(body) -> int:
    sect_pr = body.find(_w("sectPr"))
    if sect_pr is None:
        return 0
    pg_sz = sect_pr.find(_w("pgSz"))
    pg_mar = sect_pr.find(_w("pgMar"))
    if pg_sz is None or pg_mar is None:
        return 0
    page_width = _int_attr(pg_sz, _w("w"), 0)
    left = _int_attr(pg_mar, _w("left"), 0)
    right = _int_attr(pg_mar, _w("right"), 0)
    gutter = _int_attr(pg_mar, _w("gutter"), 0)
    return page_width - left - right - gutter


def _table_grid_widths(table) -> List[int]:
    grid = table.find(_w("tblGrid"))
    if grid is None:
        return []
    widths: List[int] = []
    for col in grid.findall(_w("gridCol")):
        width = _int_attr(col, _w("w"), 0)
        if width <= 0:
            return []
        widths.append(width)
    return widths


def _table_has_unsupported_horizontal_spans(table, column_count: int) -> bool:
    for row in table.findall(_w("tr")):
        cells = row.findall(_w("tc"))
        if _row_is_full_width_span(row, column_count):
            continue
        col_pos = 0
        for cell in cells:
            if cell.find("./" + _w("tcPr") + "/" + _w("vMerge")) is not None:
                return True
            if cell.find("./" + _w("tcPr") + "/" + _w("hMerge")) is not None:
                return True
            span = _grid_span(cell)
            if span != 1:
                return True
            col_pos += span
        if col_pos != column_count:
            return True
    return False


def _trailing_full_span_row_start(table, column_count: int) -> int:
    rows = table.findall(_w("tr"))
    index = len(rows)
    while index > 0 and _row_is_full_width_span(rows[index - 1], column_count):
        index -= 1
    return index


def _row_is_full_width_span(row, column_count: int) -> bool:
    cells = row.findall(_w("tc"))
    return len(cells) == 1 and _grid_span(cells[0]) == column_count


def _grid_span(cell) -> int:
    span = cell.find("./" + _w("tcPr") + "/" + _w("gridSpan"))
    if span is None:
        return 1
    return _int_attr(span, _w("val"), 1)


def _make_table_with_columns(source_table, columns: Sequence[int], widths: Sequence[int],
                             trailing_full_span_start: int,
                             include_trailing_full_span: bool):
    new_table = ET.Element(source_table.tag, source_table.attrib)
    table_width = sum(widths)
    for child in list(source_table):
        if child.tag == _w("tr"):
            continue
        if child.tag == _w("tblGrid"):
            new_table.append(_make_table_grid(widths))
            continue
        copied = copy.deepcopy(child)
        if copied.tag == _w("tblPr"):
            _set_table_property_width(copied, table_width)
        new_table.append(copied)
    for row_index, row in enumerate(source_table.findall(_w("tr"))):
        full_span = _row_is_full_width_span(row, len(_table_grid_widths(source_table)))
        if full_span and row_index >= trailing_full_span_start and not include_trailing_full_span:
            continue
        new_table.append(_make_row_with_columns(row, columns, widths, table_width, full_span))
    return new_table


def _make_table_grid(widths: Sequence[int]):
    grid = ET.Element(_w("tblGrid"))
    for width in widths:
        col = ET.SubElement(grid, _w("gridCol"))
        col.set(_w("w"), str(width))
    return grid


def _set_table_property_width(tbl_pr, width: int):
    tbl_w = tbl_pr.find(_w("tblW"))
    if tbl_w is None:
        tbl_w = ET.Element(_w("tblW"))
        tbl_pr.insert(0, tbl_w)
    tbl_w.set(_w("type"), "dxa")
    tbl_w.set(_w("w"), str(width))


def _make_row_with_columns(row, columns: Sequence[int], widths: Sequence[int],
                           table_width: int, full_span: bool):
    new_row = ET.Element(row.tag, row.attrib)
    for child in list(row):
        if child.tag == _w("tc"):
            continue
        new_row.append(copy.deepcopy(child))
    cells = row.findall(_w("tc"))
    if full_span and cells:
        cell = copy.deepcopy(cells[0])
        _set_cell_grid_span(cell, len(columns))
        _set_cell_width_twips(cell, table_width)
        new_row.append(cell)
        return new_row
    for group_index, column_index in enumerate(columns):
        if column_index >= len(cells):
            continue
        cell = copy.deepcopy(cells[column_index])
        _set_cell_width_twips(cell, widths[group_index])
        new_row.append(cell)
    return new_row


def _set_cell_grid_span(cell, span: int):
    tc_pr = cell.find(_w("tcPr"))
    if tc_pr is None:
        tc_pr = ET.Element(_w("tcPr"))
        cell.insert(0, tc_pr)
    grid_span = tc_pr.find(_w("gridSpan"))
    if span <= 1:
        if grid_span is not None:
            tc_pr.remove(grid_span)
        return
    if grid_span is None:
        grid_span = ET.Element(_w("gridSpan"))
        tc_pr.append(grid_span)
    grid_span.set(_w("val"), str(span))


def _set_cell_width_twips(cell, width: int):
    tc_pr = cell.find(_w("tcPr"))
    if tc_pr is None:
        tc_pr = ET.Element(_w("tcPr"))
        cell.insert(0, tc_pr)
    tc_w = tc_pr.find(_w("tcW"))
    if tc_w is None:
        tc_w = ET.Element(_w("tcW"))
        tc_pr.insert(0, tc_w)
    tc_w.set(_w("type"), "dxa")
    tc_w.set(_w("w"), str(width))


def _valid_continuation_breaks(row_breaks: List[int], header_count: int,
                               row_count: int) -> List[int]:
    valid_breaks: List[int] = []
    previous_break: Optional[int] = None
    for br in sorted(set(row_breaks)):
        if header_count >= br or br > row_count:
            continue
        if not _has_enough_body_rows_before_break(br, header_count, previous_break):
            continue
        valid_breaks.append(br)
        previous_break = br
    return valid_breaks


def _safe_vmerge_breaks(rows: List[ET.Element], row_breaks: List[int]) -> List[int]:
    safe_breaks: List[int] = []
    row_count = len(rows)
    for br in row_breaks:
        safe = br
        while safe <= row_count and _row_has_vmerge_continue(rows[safe - 1]):
            safe += 1
        if safe <= row_count:
            safe_breaks.append(safe)
    return safe_breaks


def _row_has_vmerge_continue(row: ET.Element) -> bool:
    for cell in row.findall(_w("tc")):
        vmerge = cell.find("./" + _w("tcPr") + "/" + _w("vMerge"))
        if vmerge is None:
            continue
        value = vmerge.get(_w("val"))
        if value in (None, "", "continue"):
            return True
    return False


def _make_table_with_rows(source_table, rows: List[ET.Element]):
    new_table = ET.Element(source_table.tag, source_table.attrib)
    for child in list(source_table):
        if child.tag != _w("tr"):
            new_table.append(copy.deepcopy(child))
    for row in rows:
        new_table.append(copy.deepcopy(row))
    return new_table


def _make_continuation_caption(source_table, text: str, style_id: str):
    caption = ET.Element(_w("p"))
    ppr = ET.SubElement(caption, _w("pPr"))
    if style_id:
        pstyle = ET.SubElement(ppr, _w("pStyle"))
        pstyle.set(_w("val"), style_id)
    else:
        prev = source_table.getprevious()
        if prev is not None and prev.tag == _w("p"):
            prev_ppr = prev.find(_w("pPr"))
            if prev_ppr is not None:
                caption.remove(ppr)
                ppr = copy.deepcopy(prev_ppr)
                caption.append(ppr)
    if ppr.find(_w("keepNext")) is None:
        ET.SubElement(ppr, _w("keepNext"))
    _set_continuation_caption_spacing(ppr)
    run = ET.SubElement(caption, _w("r"))
    t = ET.SubElement(run, _w("t"))
    t.set("{%s}space" % _XML_NS, "preserve")
    t.text = text
    return caption


def _set_continuation_caption_spacing(ppr):
    spacing = ppr.find(_w("spacing"))
    if spacing is None:
        spacing = ET.SubElement(ppr, _w("spacing"))
    spacing.set(_w("before"), str(_CAPTION_SPACING_BEFORE_AFTER))
    spacing.set(_w("after"), str(_CAPTION_SPACING_BEFORE_AFTER))
    spacing.set(_w("line"), str(_CAPTION_LINE_SPACING))
    spacing.set(_w("lineRule"), "exact")
    for attr in (_w("beforeLines"), _w("afterLines")):
        if attr in spacing.attrib:
            del spacing.attrib[attr]


def _style_id_by_name(styles_xml: bytes, style_name: str) -> str:
    root = ET.fromstring(styles_xml)
    for style in root.findall(_w("style")):
        name = style.find(_w("name"))
        if name is not None and name.get(_w("val")) == style_name:
            return style.get(_w("styleId"), "")
    return ""


def _w(tag: str) -> str:
    return "{%s}%s" % (_W_NS, tag)


def _int_attr(element, name: str, default: int = 0) -> int:
    try:
        return int(element.get(name, default))
    except (TypeError, ValueError):
        return default


def _update_all_fields(doc):
    """更新正文、页眉页脚、脚注等 story range 中的域。"""
    _safe_update(doc.Fields)
    _safe_update_collection(getattr(doc, "TablesOfContents", None))
    _safe_update_collection(getattr(doc, "TablesOfFigures", None))

    try:
        story_ranges = doc.StoryRanges
        count = story_ranges.Count
    except Exception:
        return

    for i in range(1, count + 1):
        try:
            rng = story_ranges.Item(i)
        except Exception:
            continue
        while rng is not None:
            _safe_update(getattr(rng, "Fields", None))
            try:
                rng = rng.NextStoryRange
            except Exception:
                rng = None


def _safe_update(fields):
    if fields is None:
        return
    try:
        fields.Update()
    except Exception:
        pass


def _safe_update_collection(collection):
    if collection is None:
        return
    try:
        count = collection.Count
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            collection.Item(i).Update()
        except Exception:
            pass
