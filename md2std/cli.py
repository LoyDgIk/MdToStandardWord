# -*- coding: utf-8 -*-
"""命令行入口： python -m md2std 输入.md -o 输出.docx"""

from __future__ import annotations

import argparse
import os
import sys

from . import md_parser
from . import docx_builder


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="md2std",
        description="把 Markdown 转换为符合 GB/T 1.1—2020 结构的标准文本 Word。",
    )
    parser.add_argument("input", help="输入 Markdown 文件（含 YAML front matter）")
    parser.add_argument("-o", "--output", help="输出 .docx 路径（默认与输入同名）")
    parser.add_argument("--kind", choices=("auto", "group", "national"), default="auto",
                        help="标准类型：auto 根据 standard_type/编号判断，group=团体标准，national=国家标准")
    parser.add_argument("--word-com-postprocess", "--word-com", action="store_true",
                        help="生成后调用本机 Microsoft Word COM 更新域、重新分页并保存（默认不启用）")
    parser.add_argument("--cover-form-protection", dest="cover_form_protection",
                        action="store_true", default=None,
                        help="启用封面旧式 FORMDROPDOWN 表单域保护，仅保护封面节，正文节保持可编辑")
    parser.add_argument("--no-cover-form-protection", dest="cover_form_protection",
                        action="store_false",
                        help="关闭封面表单域保护；用于覆盖 YAML 中的 cover_form_protection: true")
    args = parser.parse_args(argv)

    if not os.path.isfile(args.input):
        parser.error("找不到输入文件：%s" % args.input)

    output = args.output
    if not output:
        base, _ = os.path.splitext(args.input)
        output = base + ".docx"

    with open(args.input, "r", encoding="utf-8") as f:
        text = f.read()

    try:
        sdoc = md_parser.parse(text, source_path=args.input)
    except ValueError as exc:
        parser.error(str(exc))
    docx_builder.build_cover(
        sdoc,
        output,
        kind=args.kind,
        cover_form_protection=args.cover_form_protection,
    )
    if args.word_com_postprocess:
        from . import word_postprocess
        try:
            word_postprocess.postprocess_with_word_com(output)
        except Exception as exc:
            parser.error(str(exc))
    # 避免中文控制台编码问题，使用 ascii 安全输出
    sys.stdout.write("OK -> %s\n" % output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
