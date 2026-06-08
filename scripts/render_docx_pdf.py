# -*- coding: utf-8 -*-
"""Export DOCX to PDF with Microsoft Word COM and optionally render PNG pages.

This script is intentionally Windows/Word based and does not call LibreOffice.
PNG rendering uses the optional ``pypdfium2`` package.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
import threading
from pathlib import Path
from typing import Iterable, List, Optional

_WD_EXPORT_FORMAT_PDF = 17
_WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
_WD_EXPORT_ALL_DOCUMENT = 0
_WD_EXPORT_DOCUMENT_CONTENT = 0
_WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1


class WordProcessGuard:
    """Kill only the Word process created for this export if it times out."""

    def __init__(self, pid: Optional[int], timeout_seconds: int):
        self.pid = pid
        self.timeout_seconds = timeout_seconds
        self.timed_out = False
        self._timer: Optional[threading.Timer] = None

    def start(self) -> None:
        if not self.pid or self.timeout_seconds <= 0:
            return
        self._timer = threading.Timer(self.timeout_seconds, self._on_timeout)
        self._timer.daemon = True
        self._timer.start()

    def cancel(self) -> None:
        if self._timer is not None:
            self._timer.cancel()
            self._timer = None

    def _on_timeout(self) -> None:
        self.timed_out = True
        _kill_process_tree(self.pid)


def _kill_process_tree(pid: Optional[int]) -> None:
    if not pid:
        return
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/T", "/F"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )


def _word_process_id(word) -> Optional[int]:
    try:
        import win32process  # type: ignore

        hwnd = int(word.Hwnd)
        if hwnd:
            return int(win32process.GetWindowThreadProcessId(hwnd)[1])
    except Exception:
        return None
    return None


def _quit_word(word) -> None:
    try:
        word.Quit()
    except Exception:
        pass


def _open_document(word, path: Path):
    try:
        return word.Documents.Open(
            FileName=str(path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Revert=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
    except Exception:
        return word.Documents.Open(str(path))


def _safe_update(fields) -> None:
    if fields is None:
        return
    try:
        fields.Update()
    except Exception:
        pass


def _safe_update_collection(collection) -> None:
    if collection is None:
        return
    try:
        count = collection.Count
    except Exception:
        return
    for index in range(1, count + 1):
        try:
            collection.Item(index).Update()
        except Exception:
            pass


def _update_all_fields(doc) -> None:
    _safe_update(getattr(doc, "Fields", None))
    _safe_update_collection(getattr(doc, "TablesOfContents", None))
    _safe_update_collection(getattr(doc, "TablesOfFigures", None))

    try:
        story_ranges = doc.StoryRanges
        count = story_ranges.Count
    except Exception:
        return

    for index in range(1, count + 1):
        try:
            rng = story_ranges.Item(index)
        except Exception:
            continue
        while rng is not None:
            _safe_update(getattr(rng, "Fields", None))
            try:
                rng = rng.NextStoryRange
            except Exception:
                rng = None


def _temp_pdf_path(output_pdf: Path) -> Path:
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix=".md2std-word-pdf-",
        suffix=".pdf",
        dir=str(output_pdf.parent),
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    tmp_path.unlink(missing_ok=True)
    return tmp_path


def export_pdf_with_word_com(
    input_docx: Path,
    output_pdf: Path,
    *,
    visible: bool = False,
    update_fields: bool = True,
    timeout_seconds: int = 180,
) -> Path:
    """Export ``input_docx`` to ``output_pdf`` through Microsoft Word COM."""
    input_docx = input_docx.resolve()
    output_pdf = output_pdf.resolve()
    if not input_docx.is_file():
        raise FileNotFoundError(input_docx)
    if output_pdf.suffix.lower() != ".pdf":
        output_pdf = output_pdf.with_suffix(".pdf")

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "需要安装 pywin32，并且只能在装有 Microsoft Word 的 Windows 环境使用："
            " python -m pip install pywin32"
        ) from exc

    tmp_pdf = _temp_pdf_path(output_pdf)
    pythoncom.CoInitialize()
    word = None
    doc = None
    guard = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = bool(visible)
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3
        except Exception:
            pass
        guard = WordProcessGuard(_word_process_id(word), timeout_seconds)
        guard.start()
        doc = _open_document(word, input_docx)
        if update_fields:
            _update_all_fields(doc)
            try:
                doc.Repaginate()
            except Exception:
                pass
        doc.ExportAsFixedFormat(
            OutputFileName=str(tmp_pdf),
            ExportFormat=_WD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
            OptimizeFor=_WD_EXPORT_OPTIMIZE_FOR_PRINT,
            Range=_WD_EXPORT_ALL_DOCUMENT,
            From=1,
            To=1,
            Item=_WD_EXPORT_DOCUMENT_CONTENT,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=_WD_EXPORT_CREATE_HEADING_BOOKMARKS,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
        if guard.timed_out:
            raise TimeoutError("Word COM 导出 PDF 超时。")
        if not tmp_pdf.is_file() or tmp_pdf.stat().st_size <= 0:
            raise RuntimeError("Word COM 未生成有效 PDF。")
        os.replace(tmp_pdf, output_pdf)
        return output_pdf
    finally:
        if guard is not None:
            guard.cancel()
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            _quit_word(word)
        pythoncom.CoUninitialize()
        if tmp_pdf.exists():
            tmp_pdf.unlink(missing_ok=True)


def render_pdf_pages(pdf_path: Path, output_dir: Path, *, scale: float = 2.0) -> List[Path]:
    """Render PDF pages to PNG files using optional pypdfium2."""
    try:
        import pypdfium2 as pdfium  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "PDF 已导出，但渲染 PNG 需要 pypdfium2：python -m pip install pypdfium2。"
            "如只需要 PDF，可加 --no-png。"
        ) from exc

    output_dir.mkdir(parents=True, exist_ok=True)
    doc = pdfium.PdfDocument(str(pdf_path))
    pages: List[Path] = []
    try:
        for index in range(len(doc)):
            page = doc[index]
            image = page.render(scale=scale).to_pil()
            out = output_dir / f"page-{index + 1}.png"
            image.save(out)
            pages.append(out)
    finally:
        close = getattr(doc, "close", None)
        if callable(close):
            close()
    return pages


def _default_pdf_path(input_docx: Path) -> Path:
    return input_docx.with_suffix(".pdf")


def _default_pages_dir(output_pdf: Path) -> Path:
    return output_pdf.with_name(output_pdf.stem + "_pages")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="使用 Microsoft Word COM 导出 DOCX 为 PDF，并可渲染 PDF 页面 PNG；不依赖 LibreOffice。"
    )
    parser.add_argument("input", help="输入 DOCX 文件。")
    parser.add_argument("-o", "--output", help="输出 PDF 路径；默认与 DOCX 同名。")
    parser.add_argument("--pages-dir", help="PNG 页面输出目录；默认 <pdf文件名>_pages。")
    parser.add_argument("--scale", type=float, default=2.0, help="PNG 渲染倍率，默认 2.0。")
    parser.add_argument("--no-png", action="store_true", help="只导出 PDF，不渲染页面 PNG。")
    parser.add_argument("--visible", action="store_true", help="显示 Word 窗口，便于调试。")
    parser.add_argument("--no-update-fields", action="store_true", help="导出前不更新域和重新分页。")
    parser.add_argument("--timeout", type=int, default=180, help="Word COM 超时时间，秒；默认 180。")
    return parser


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    input_docx = Path(args.input)
    output_pdf = Path(args.output) if args.output else _default_pdf_path(input_docx)
    output_pdf = export_pdf_with_word_com(
        input_docx,
        output_pdf,
        visible=args.visible,
        update_fields=not args.no_update_fields,
        timeout_seconds=args.timeout,
    )
    print("PDF -> %s" % output_pdf)

    if not args.no_png:
        pages_dir = Path(args.pages_dir) if args.pages_dir else _default_pages_dir(output_pdf)
        pages = render_pdf_pages(output_pdf, pages_dir, scale=args.scale)
        print("PNG -> %s (%d pages)" % (pages_dir, len(pages)))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        raise SystemExit(1)
