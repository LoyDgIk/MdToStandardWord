# -*- coding: utf-8 -*-
"""命令行入口： python -m md2std 输入.md -o 输出.docx"""

from __future__ import annotations

import argparse
import os
import sys

from . import md_parser
from . import docx_builder
from . import resources

# 默认完整模板：template 后端按 kind 选择；_DEFAULT_TEMPLATE 保持团体模板用于旧测试兼容。


def _resolve_default_template(kind="group"):
    if kind == "national":
        candidates = resources.template_candidates("template_national.docx")
    else:
        candidates = resources.template_candidates("template_group.docx")
    for c in candidates:
        if os.path.isfile(c):
            return c
    raise FileNotFoundError("找不到默认模板：%s" % kind)


def _resolve_kind(kind, meta):
    if kind != "auto":
        return kind
    standard_type = (meta.standard_type or "").strip()
    number = (meta.number or "").strip().upper()
    if "国家" in standard_type or number.startswith("GB"):
        return "national"
    return "group"


_DEFAULT_TEMPLATE = _resolve_default_template("group")
_DEFAULT_NATIONAL_TEMPLATE = _resolve_default_template("national")


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="md2std",
        description="把 Markdown 转换为符合 GB/T 1.1—2020 结构的标准文本 Word。",
    )
    parser.add_argument("input", help="输入 Markdown 文件（含 YAML front matter）")
    parser.add_argument("-o", "--output", help="输出 .docx 路径（默认与输入同名）")
    parser.add_argument("--backend", choices=("cover", "template"), default="cover",
                        help="生成后端：cover=封面蓝图直接生成（默认），template=完整模板替换旧模式")
    parser.add_argument("--kind", choices=("auto", "group", "national"), default="auto",
                        help="标准类型：auto 根据 standard_type/编号判断，group=团体标准，national=国家标准")
    parser.add_argument("-t", "--template",
                        help="template 后端使用的完整模板 .docx 路径（默认按 kind 自动选择）")
    parser.add_argument("--word-com-postprocess", "--word-com", action="store_true",
                        help="生成后调用本机 Microsoft Word COM 更新域、重新分页并保存（默认不启用）")
    args = parser.parse_args(argv)

    if not os.path.isfile(args.input):
        parser.error("找不到输入文件：%s" % args.input)

    output = args.output
    if not output:
        base, _ = os.path.splitext(args.input)
        output = base + ".docx"

    with open(args.input, "r", encoding="utf-8") as f:
        text = f.read()

    sdoc = md_parser.parse(text)
    if args.backend == "template":
        resolved_kind = _resolve_kind(args.kind, sdoc.meta)
        template_path = args.template or _resolve_default_template(resolved_kind)
        if not os.path.isfile(template_path):
            parser.error("找不到模板文件：%s" % template_path)
        docx_builder.build(sdoc, template_path, output, kind=resolved_kind)
    else:
        docx_builder.build_cover(sdoc, output, kind=args.kind)
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
