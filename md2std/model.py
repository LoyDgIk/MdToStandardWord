# -*- coding: utf-8 -*-
"""中间数据模型：把 Markdown 解析结果表示为与排版无关的结构，供 docx_builder 使用。

设计原则：解析层（md_parser）只关心"这是什么内容"，不关心样式；
构建层（docx_builder）负责把这些模型映射到模板的命名样式。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional


# --------------------------------------------------------------------------- #
# 行内文本片段（用于加粗/斜体）
# --------------------------------------------------------------------------- #
@dataclass
class Span:
    """一段行内文本及其格式。"""
    text: str
    bold: bool = False
    italic: bool = False
    subscript: bool = False
    superscript: bool = False


@dataclass
class FormulaSpan(Span):
    """行内公式片段，text 保存 LaTeX 源码。"""
    pass


@dataclass
class RefSpan(Span):
    """一个类型化交叉引用，如 {{tbl:classify:label}}。"""
    ref_type: str = ""
    target: str = ""
    mode: str = "num"


@dataclass
class TableCellPart:
    """表格单元格内片段。kind=text/formula/ref，formula 使用 LaTeX 源码。"""
    kind: str
    text: str = ""
    ref_type: str = ""
    target: str = ""
    mode: str = "num"


# --------------------------------------------------------------------------- #
# 块级元素
# --------------------------------------------------------------------------- #
@dataclass
class Heading:
    """章/条标题。level 从 1 开始：

    1 -> 章标题，2 -> 一级条标题，3 -> 二级条标题，……（正文）。
    标题文本不含编号，编号由模板样式自动生成。
    """
    level: int
    spans: List[Span] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.spans)


@dataclass
class Paragraph:
    """普通段落 -> 标准文件_段。"""
    spans: List[Span] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.spans)


@dataclass
class PageBreak:
    """显式分页控制符。"""
    pass


@dataclass
class UntitledClause:
    """无标题条：带编号但无标题的悬挂段。

    level   编号层级，如 3 表示 X.Y.Z。
    spans   条文内容（不含编号）。
    """
    level: int
    spans: List[Span] = field(default_factory=list)

    @property
    def segments(self) -> int:
        return self.level

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.spans)


@dataclass
class ListItem:
    """列项内容。"""
    spans: List[Span] = field(default_factory=list)


@dataclass
class ListBlock:
    """列表块。

    ordered=False -> 破折号列项（—）
    ordered=True  -> 字母编号列项（a) b) …）
    嵌套有序列表由 md_parser 拆成 level 区分，docx_builder 据 level 选样式。
    """
    ordered: bool
    items: List[ListItem] = field(default_factory=list)
    level: int = 1


@dataclass
class Note:
    """注：/ 注X：。

    index 为 None 表示"注："（单条），否则为"注1：注2："等。
    """
    spans: List[Span] = field(default_factory=list)
    index: Optional[int] = None


@dataclass
class Example:
    """示例：/ 示例X：。"""
    spans: List[Span] = field(default_factory=list)
    index: Optional[int] = None


@dataclass
class ExampleContent:
    """示例后续内容 -> 标准文件_示例内容。"""
    spans: List[Span] = field(default_factory=list)


@dataclass
class Source:
    """术语来源标注 -> [来源：...]，套用 标准文件_来源 段落。"""
    text: str = ""


@dataclass
class Term:
    """术语条目（术语和定义章内）。

    term     中文术语
    term_en  英文对应词（可空）
    definition 定义段落 spans
    notes    可选注
    source   可选来源
    """
    term: str
    term_en: str = ""
    definition: List[Span] = field(default_factory=list)
    notes: List[Note] = field(default_factory=list)
    source: Optional[Source] = None


@dataclass
class TableModel:
    """表格。header 为表头行，rows 为数据行；每个单元格是纯文本。

    caption 仅为表题文字（不含"表N"，编号由 SEQ 域自动生成）。
    anchor_id 为类型化交叉引用本地 id（来自 `{#tbl:id}`）。
    header_colspans/row_colspans 对应 HTML 表格的 colspan；GFM 表格默认为 1。
    header_parts/row_parts 保存单元格内文本/公式片段；header/rows 保留纯文本视图。
    """
    ref_type: str = "tbl"
    caption: str = ""
    anchor_id: str = ""
    header: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    header_colspans: List[int] = field(default_factory=list)
    row_colspans: List[List[int]] = field(default_factory=list)
    header_parts: List[List[TableCellPart]] = field(default_factory=list)
    row_parts: List[List[List[TableCellPart]]] = field(default_factory=list)


@dataclass
class Figure:
    """插图。caption 仅为图题文字（不含"图N"）。"""
    ref_type: str = "fig"
    caption: str = ""
    path: str = ""
    anchor_id: str = ""


@dataclass
class Formula:
    """公式：LaTeX 源码。序号由 SEQ 域自动生成（body:（N）/appendix:（A.N））。

    anchor_id 为类型化交叉引用本地 id（来自 `{#eq:id}`）。
    """
    latex: str
    ref_type: str = "eq"
    anchor_id: str = ""


# --------------------------------------------------------------------------- #
# 章节容器
# --------------------------------------------------------------------------- #
@dataclass
class Appendix:
    """附录块。

    nature: 'normative'(规范性) 或 'informative'(资料性)
    letter: A/B/C…（由解析顺序自动分配，仅用于展示判断；编号靠模板自动生成）
    blocks: 附录正文块序列（Heading.level 在附录内：1->附录章标题(标题本身另存 title)，
            实际子条 level>=2 用附录条标题样式）
    """
    nature: str
    title_spans: List[Span] = field(default_factory=list)
    blocks: List[object] = field(default_factory=list)

    @property
    def title(self) -> str:
        return "".join(s.text for s in self.title_spans)


@dataclass
class IndexItem:
    """索引项：左侧术语 + 右侧条文/图表位置列表。"""
    term: str
    targets: str


@dataclass
class IndexGroup:
    """索引字母分组，如 B / C。"""
    letter: str
    items: List[IndexItem] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# 前言信息
# --------------------------------------------------------------------------- #
@dataclass
class Foreword:
    multipart_note: str = ""
    replace_changes: List[str] = field(default_factory=list)
    patent_note: bool = True
    proposer: str = ""
    owner: str = ""
    draft_orgs: List[str] = field(default_factory=list)
    drafters: List[str] = field(default_factory=list)
    history: str = ""
    extra_notes: List[object] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# 元数据
# --------------------------------------------------------------------------- #
@dataclass
class Meta:
    standard_type: str = "团体标准"
    number: str = ""
    replaces: str = ""
    title: str = ""
    title_en: str = ""
    ics: str = ""
    ccs: str = ""
    publish_date: str = ""
    implement_date: str = ""
    publisher: str = ""
    foreword: Foreword = field(default_factory=Foreword)
    introduction: str = ""
    odd_even_pages: bool = False


# --------------------------------------------------------------------------- #
# 顶层文档模型
# --------------------------------------------------------------------------- #
@dataclass
class StandardDoc:
    meta: Meta = field(default_factory=Meta)
    # 正文块序列（章/条/段/列表/表/图/注/示例/术语）
    body: List[object] = field(default_factory=list)
    appendices: List[Appendix] = field(default_factory=list)
    # 参考文献条目（每条一段）
    references: List[str] = field(default_factory=list)
    # 索引分组（可选）
    index_groups: List[IndexGroup] = field(default_factory=list)
