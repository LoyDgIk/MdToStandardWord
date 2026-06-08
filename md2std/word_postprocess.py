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
from typing import List

from lxml import etree as ET

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_NS = "http://www.w3.org/XML/1998/namespace"
_WD_ACTIVE_END_PAGE_NUMBER = 3
_CONTINUATION_STYLE_NAME = "标准文件_表格续"
_CONTINUATION_SUFFIX = "（续）"
_WORD_COM_TIMEOUT_SECONDS = 180


@dataclass
class _TableContinuationPlan:
    """一个表格的真实跨页拆分计划。row_breaks 为 Word 1-based 行号。"""

    table_index: int
    row_breaks: List[int]
    header_count: int
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
    max_iterations = 4
    for _ in range(max_iterations):
        plans = _measure_table_continuations(word, abs_path)
        if not plans:
            return
        if not _apply_table_continuations(abs_path, plans):
            return
    # 最后一轮保存一次 Word 分页结果；若仍有极端超高行，避免无限循环。
    _measure_table_continuations(word, abs_path)


def _measure_table_continuations(word, abs_path: str) -> List[_TableContinuationPlan]:
    """让 Word 完成分页后，收集真实跨页表格的断开行。"""
    doc = None
    try:
        doc = _open_document(word, abs_path)
        _update_all_fields(doc)
        _disable_row_page_breaks(doc)
        try:
            doc.Repaginate()
        except Exception:
            pass
        plans = _collect_table_continuation_plans(doc)
        doc.Save()
        return plans
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
        if row_breaks:
            plans.append(_TableContinuationPlan(
                table_index=table_index,
                row_breaks=row_breaks,
                header_count=header_count,
                caption_text=_continuation_caption_text(caption_text),
            ))
    return plans


def _table_caption_text(doc, table) -> str:
    before = doc.Range(0, table.Range.Start)
    para = before.Paragraphs.Last
    text = _clean_word_text(para.Range.Text)
    try:
        list_string = _clean_word_text(para.Range.ListFormat.ListString)
    except Exception:
        list_string = ""
    if list_string:
        return (list_string + "　" + text).strip()
    return text


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
    for row_index in range(1, row_count + 1):
        try:
            is_header = bool(table.Rows(row_index).HeadingFormat)
        except Exception:
            break
        if not is_header:
            break
        count += 1
    return count


def _table_row_page_breaks(doc, table, row_count: int, header_count: int) -> List[int]:
    breaks: List[int] = []
    previous_page = None
    for row_index in range(1, row_count + 1):
        try:
            row = table.Rows(row_index)
            page = doc.Range(row.Range.Start, row.Range.Start).Information(_WD_ACTIVE_END_PAGE_NUMBER)
        except Exception:
            continue
        if not isinstance(page, int):
            continue
        if previous_page is not None and page != previous_page:
            # 避免把“只有表头留在上一页”的表硬拆成续表；这种情况应另行整表前置。
            if row_index > header_count + 1:
                breaks.append(row_index)
        previous_page = page
    return breaks


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


def _split_table_element(table, plan: _TableContinuationPlan, style_id: str) -> List[ET.Element]:
    children = list(table)
    rows = [child for child in children if child.tag == _w("tr")]
    if len(rows) < 2:
        return []

    valid_breaks = sorted(set(
        br for br in plan.row_breaks
        if plan.header_count < br <= len(rows)
    ))
    if not valid_breaks:
        return []

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
    if ppr.find(_w("pageBreakBefore")) is None:
        ET.SubElement(ppr, _w("pageBreakBefore"))
    run = ET.SubElement(caption, _w("r"))
    t = ET.SubElement(run, _w("t"))
    t.set("{%s}space" % _XML_NS, "preserve")
    t.text = text
    return caption


def _style_id_by_name(styles_xml: bytes, style_name: str) -> str:
    root = ET.fromstring(styles_xml)
    for style in root.findall(_w("style")):
        name = style.find(_w("name"))
        if name is not None and name.get(_w("val")) == style_name:
            return style.get(_w("styleId"), "")
    return ""


def _w(tag: str) -> str:
    return "{%s}%s" % (_W_NS, tag)


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
