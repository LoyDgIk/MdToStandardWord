# -*- coding: utf-8 -*-
"""命令行入口： python -m md2std 输入.md -o 输出.docx"""

from __future__ import annotations

import argparse
import os
import sys

from . import md_parser
from . import docx_builder

# 默认模板：优先 templates/团体标准模板.docx，回退到项目根的原始模板
_PKG_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJ_DIR = os.path.dirname(_PKG_DIR)


def _resolve_default_template():
    candidates = [
        os.path.join(_PROJ_DIR, "templates", "团体标准模板.docx"),
        os.path.join(_PROJ_DIR, "2 团体标准——模板.docx"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return candidates[0]


_DEFAULT_TEMPLATE = _resolve_default_template()


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="md2std",
        description="把 Markdown 转换为团体标准（GB/T 1.1—2020）标准文本 Word。",
    )
    parser.add_argument("input", help="输入 Markdown 文件（含 YAML front matter）")
    parser.add_argument("-o", "--output", help="输出 .docx 路径（默认与输入同名）")
    parser.add_argument("-t", "--template", default=_DEFAULT_TEMPLATE,
                        help="模板 .docx 路径（默认使用内置团体标准模板）")
    args = parser.parse_args(argv)

    if not os.path.isfile(args.input):
        parser.error("找不到输入文件：%s" % args.input)
    if not os.path.isfile(args.template):
        parser.error("找不到模板文件：%s" % args.template)

    output = args.output
    if not output:
        base, _ = os.path.splitext(args.input)
        output = base + ".docx"

    with open(args.input, "r", encoding="utf-8") as f:
        text = f.read()

    sdoc = md_parser.parse(text)
    docx_builder.build(sdoc, args.template, output)
    # 避免中文控制台编码问题，使用 ascii 安全输出
    sys.stdout.write("OK -> %s\n" % output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
