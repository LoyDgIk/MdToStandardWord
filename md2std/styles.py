# -*- coding: utf-8 -*-
"""模板命名样式常量与层级映射。

模板（2 团体标准——模板.docx）采用 GB/T 1.1—2020 命名样式体系，
正文章条编号靠样式联动多级列表（numId=2）自动生成、附录靠 numId=4 自动生成，
因此构建时只需套用正确的样式名、不写编号文字。
"""

# --- 封面 / 元数据字段样式 ---------------------------------------------------
S_COVER_TYPE = "标准称谓"              # “团体标准”
S_COVER_NUMBER = "标准文件_文件编号"     # T/XXX XXXX—XXXX
S_COVER_REPLACES = "标准文件_替换文件编号"
S_COVER_NAME = "标准文件_文件名称"       # 封面标准中文名称
S_COVER_NAME_EN = "封面标准英文名称"
S_COVER_PUBLISH = "其他发布日期"         # XXXX-XX-XX发布
S_COVER_IMPLEMENT = "其他实施日期"       # XXXX-XX-XX实施
S_COVER_PUBLISHER = "其他发布部门"       # 发布部门

# --- 节标题 -----------------------------------------------------------------
S_TOC_TITLE = "标准文件_目录标题"        # 目次
S_PREFACE_TITLE = "标准文件_前言、引言标题"  # 前言 / 引言
S_REF_TITLE = "标准文件_参考文献标题"     # 参考文献
S_INDEX_TITLE = "标准文件_索引标题"       # 索引
S_INDEX_LETTER = "标准文件_索引字母"      # A/B/C...
S_INDEX_ITEM = "标准文件_索引项"          # 术语 + 点引导 + 位置列表
S_BODY_STANDARD_NAME = "标准文件_正文标准名称"  # 正文首页标准名称标题

# --- 正文章条标题（level 1..6） ----------------------------------------------
S_CHAPTER = "标准文件_章标题"            # 1  范围
S_CLAUSE_1 = "标准文件_一级条标题"        # 5.1
S_CLAUSE_2 = "标准文件_二级条标题"        # 5.1.1
S_CLAUSE_3 = "标准文件_三级条标题"        # 5.1.1.1
S_CLAUSE_4 = "标准文件_四级条标题"        # 5.1.1.1.1
S_CLAUSE_5 = "标准文件_五级条标题"        # 5.1.1.1.1.1

# Markdown 标题层级（#=1）-> 正文样式
HEADING_STYLE_BY_LEVEL = {
    1: S_CHAPTER,
    2: S_CLAUSE_1,
    3: S_CLAUSE_2,
    4: S_CLAUSE_3,
    5: S_CLAUSE_4,
    6: S_CLAUSE_5,
}

# --- 正文段落与列项 ----------------------------------------------------------
S_PARA = "标准文件_段"
S_LIST_DASH = "标准文件_破折号列项"        # — （无序）
S_LIST_LETTER = "标准文件_字母编号列项（一级）"  # a) b)（有序一级）
S_LIST_NUMBER = "标准文件_数字编号列项"     # 1) 2)（独立数字列项）
S_LIST_NUMBER_2 = "标准文件_数字编号列项（二级）"  # 嵌套 1) 2)
S_LIST_NUMBER_3 = "标准文件_编号列项（三级）"      # 嵌套 (1) (2)

# 有序列表层级 -> 样式
ORDERED_LIST_STYLE_BY_LEVEL = {
    1: S_LIST_LETTER,
    2: S_LIST_NUMBER_2,
    3: S_LIST_NUMBER_3,
}

# --- 无标题条（无标题、带编号的悬挂条） --------------------------------------
# 这些样式按对应"条标题"排版但不进 TOC；编号由正文多级列表自动生成。
S_UNTITLED_1 = "标准文件_一级无标题"   # X.Y
S_UNTITLED_2 = "标准文件_二级无标题"   # X.Y.Z
S_UNTITLED_3 = "标准文件_三级无标题"   # X.Y.Z.W
S_UNTITLED_4 = "标准文件_四级无标题"
S_UNTITLED_5 = "标准文件_五级无标题"

# 编号段数 -> 无标题条样式（2 段=X.Y -> 一级无标题，3 段=X.Y.Z -> 二级无标题 …）
UNTITLED_STYLE_BY_SEGMENTS = {
    2: S_UNTITLED_1,
    3: S_UNTITLED_2,
    4: S_UNTITLED_3,
    5: S_UNTITLED_4,
    6: S_UNTITLED_5,
}

# 正文章条的多级列表 numId（章标题/条标题/无标题条共用，保证编号连续）。
# ilvl 约定：章=1，一级条(X.Y)=2，二级条(X.Y.Z)=3 …… 即 ilvl = 编号段数。
NUM_BODY = 2

# --- 注 / 示例 / 来源 --------------------------------------------------------
S_NOTE = "标准文件_注："                  # 注：（单条）
S_NOTE_X = "标准文件_注×："               # 注1：注2：（多条）
S_EXAMPLE = "标准文件_示例："             # 示例：（单条）
S_EXAMPLE_CONTENT = "标准文件_示例内容"
S_SOURCE = "标准文件_来源"               # [来源：...]（字符样式）

# --- 术语条标题（按层级，可选精排；v1 用条标题样式即可） ----------------------
S_TERM_1 = "标准文件_术语条一"
S_TERM_2 = "标准文件_术语条二"

# --- 附录 -------------------------------------------------------------------
S_APPENDIX_MARK = "标准文件_附录标识"      # “附录A”（自动，文本留空）
S_APPENDIX_NATURE = "附录性质"            # （规范性）/（资料性）
S_APPENDIX_TITLE = "标准文件_附录章标题"    # 附录标题
S_APPENDIX_CLAUSE_1 = "标准文件_附录一级条标题"  # A.1
S_APPENDIX_CLAUSE_2 = "标准文件_附录二级条标题"  # A.1.1
S_APPENDIX_CLAUSE_3 = "标准文件_附录三级条标题"
S_APPENDIX_CLAUSE_4 = "标准文件_附录四级条标题"
S_APPENDIX_CLAUSE_5 = "标准文件_附录五级条标题"

# 附录内 Markdown 标题层级（##=2 为附录内第一级条）-> 样式
APPENDIX_CLAUSE_STYLE_BY_LEVEL = {
    2: S_APPENDIX_CLAUSE_1,
    3: S_APPENDIX_CLAUSE_2,
    4: S_APPENDIX_CLAUSE_3,
    5: S_APPENDIX_CLAUSE_4,
    6: S_APPENDIX_CLAUSE_5,
}

# --- 公式 -------------------------------------------------------------------
S_FORMULA = "标准文件_正文公式"          # 居中公式 + 右对齐序号（带点引导）
S_FORMULA_APPENDIX = "标准文件_附录公式"  # 附录公式

# --- 表 / 图 ----------------------------------------------------------------
S_TABLE_CAPTION = "标准文件_正文表标题"     # 表N 标题
S_FIGURE_CAPTION = "标准文件_正文图标题"     # 图N 标题
S_APPENDIX_TABLE_CAPTION = "标准文件_附录表标题"   # 表A.N 标题
S_APPENDIX_FIGURE_CAPTION = "标准文件_附录图标题"  # 图A.N 标题
S_TABLE_CELL = "标准文件_表格"             # 表格单元格内文字
S_REF_ITEM = "标准文件_参考文献条目"

# --- 内置 ------------------------------------------------------------------
S_NORMAL = "Normal"
S_TABLE_GRID = "Table Grid"

# 识别"标准章"的固定章名（用于判断规范性引用文件/术语等特殊处理）
CH_SCOPE = "范围"
CH_NORMATIVE_REF = "规范性引用文件"
CH_TERMS = "术语和定义"
CH_SYMBOLS = "符号和缩略语"
CH_REFERENCES = "参考文献"
