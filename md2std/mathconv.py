# -*- coding: utf-8 -*-
"""LaTeX 数学公式 -> Word 原生公式(OMML) 元素。

流程：LaTeX --latex2mathml--> MathML --MML2OMML.XSL(XSLT)--> OMML(m:oMath)。
MML2OMML.XSL 随 Microsoft Office 安装提供，自动探测其路径；找不到时返回 None，
调用方应回退为纯文本，避免中断。
"""

from __future__ import annotations

import glob
import os
from typing import Optional

from lxml import etree

_OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_XSL_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office*\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office*\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office*\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office*\MML2OMML.XSL",
]

_transform = None
_init_done = False


def _find_xsl() -> Optional[str]:
    env = os.environ.get("MD2STD_MML2OMML_XSL")
    if env and os.path.isfile(env):
        return env
    for pat in _XSL_CANDIDATES:
        hits = sorted(glob.glob(pat))
        if hits:
            return hits[-1]
    return None


def _get_transform():
    global _transform, _init_done
    if _init_done:
        return _transform
    _init_done = True
    xsl = _find_xsl()
    if not xsl:
        return None
    try:
        _transform = etree.XSLT(etree.parse(xsl))
    except Exception:
        _transform = None
    return _transform


def available() -> bool:
    """LaTeX->OMML 工具链是否可用。"""
    try:
        import latex2mathml.converter  # noqa: F401
    except Exception:
        return False
    return _get_transform() is not None


def latex_to_omml(latex: str):
    """把 LaTeX 转为 OMML 的 m:oMath 元素（lxml Element）；失败返回 None。"""
    try:
        import latex2mathml.converter as conv
    except Exception:
        return None
    transform = _get_transform()
    if transform is None:
        return None
    try:
        mathml = conv.convert(latex)
        dom = etree.fromstring(mathml.encode("utf-8"))
        omml = transform(dom)
        root = omml.getroot()
    except Exception:
        return None
    if root is None:
        return None
    # 规整：取 m:oMath；若包了 m:oMathPara 则取内部首个 m:oMath
    tag = etree.QName(root.tag)
    if tag.localname == "oMath":
        return root
    if tag.localname == "oMathPara":
        for ch in root:
            if etree.QName(ch.tag).localname == "oMath":
                return ch
    # 兜底：查找任意 oMath
    for el in root.iter("{%s}oMath" % _OMML_NS):
        return el
    return None
