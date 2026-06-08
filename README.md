# MdToStandardWord

`md2std` 是一个将中文标准草稿 Markdown 转换为标准文本 Word `.docx` 的命令行工具。项目面向符合 GB/T 1.1-2020 起草规则的标准文件，支持团体标准和国家标准的封面、前言、引言、目次、正文、附录、参考文献、索引、图表公式题注和交叉引用。

本仓库只维护 CLI 工具本体和生成器必须解析的输入契约。配套 Codex Skill 已拆分到独立仓库：[LoyDgIk/md2std-standard-skill](https://github.com/LoyDgIk/md2std-standard-skill)。标准审校流程由 `standard-audit-workflow` 承担，确定性审校规则由 [LoyDgIk/md2std-standard-auditor](https://github.com/LoyDgIk/md2std-standard-auditor) 维护。

## 功能

- 根据 Markdown 和 YAML front matter 生成标准文本 `.docx`。
- 支持团体标准、国家标准封面字段填充。
- 支持封面一致性程度标识、备案号字段和正文首页重要提示。
- 支持规范性引用文件、术语和定义、符号和缩略语章节的常用导语自动补齐。
- 支持显式术语块，并兼容旧式二级标题术语写法。
- 支持标准章节结构、条款层级、无标题条、列项、注、示例。
- 支持 GFM 表格和简单 HTML 表格，HTML 表格可保留横向合并单元格。
- 支持表/图单位、来源、脚注等紧邻式图表附加项，并支持表格单元格内普通注。
- 支持正文公式和表格单元格内公式片段。
- 支持行内上标、下标，兼容 Typora/Obsidian 风格 `^...^`、`~...~` 扩展 Markdown。
- 支持行内公式片段 `$$...$$`，用于让正文中的数学变量保持与公式一致的 Word 原生公式字体。
- 支持表、图、公式、规范性引用文件的类型化交叉引用。
- 支持显式分页符和可选 Word COM 后处理；后处理可更新域、重新分页并按真实分页处理续表。
- 模板资产随 Python 包安装，安装后不依赖仓库根目录的 `templates/` 文件夹。

## 环境要求

- Python 3.9 或更高版本。
- Windows、macOS、Linux 均可执行基础转换。
- Word COM 后处理仅适用于安装 Microsoft Word 的 Windows 环境。
- 公式转换依赖 `latex2mathml`，并使用 Office 自带 `MML2OMML.XSL`；缺失时公式会回退为文本输出。

## 安装

开发安装：

```powershell
python -m pip install -e .
```

也可以从 GitHub tag 安装：

```powershell
python -m pip install "md2std @ git+https://github.com/LoyDgIk/MdToStandardWord.git@v0.1.0"
```

主要依赖：

- `python-docx`
- `markdown-it-py`
- `PyYAML`
- `latex2mathml`
- `lxml`

## 使用

模块入口：

```powershell
python -X utf8 -m md2std examples/地热温泉资源开发利用规范.md -o examples/地热温泉资源开发利用规范.docx
```

安装后也可直接使用 console script：

```powershell
md2std examples/地热温泉资源开发利用规范.md -o examples/地热温泉资源开发利用规范.docx
```

生成国家标准文档：

```powershell
md2std examples/汽车、摩托车用车速表.md -o examples/汽车、摩托车用车速表.docx --kind national
```

最终交付前启用 Word COM 后处理：

```powershell
md2std input.md -o output.docx --word-com-postprocess
```

常用参数：

| 参数 | 说明 |
| --- | --- |
| `input` | 输入 Markdown 文件路径。 |
| `-o, --output` | 输出 `.docx` 文件路径；未指定时输出到输入文件同名 `.docx`。 |
| `--kind {auto,group,national}` | 标准类型；默认 `auto`，根据 `standard_type` 或标准编号推断。 |
| `--word-com-postprocess`, `--word-com` | 调用本机 Microsoft Word COM 更新域、重新分页并保存。 |
| `--cover-form-protection` | 启用封面旧式 `FORMDROPDOWN` 表单域保护；仅保护封面节，正文节保持可编辑。 |
| `--no-cover-form-protection` | 关闭封面表单域保护；用于覆盖 YAML 中的 `cover_form_protection: true`。 |

生成后的文档首次在 Word 中打开时，如提示是否更新域，应选择更新。未启用 Word COM 后处理时，也可在 Word 中全选正文后按 `F9` 更新域。

### PDF 导出与页面渲染

项目提供 Windows 专用脚本 `scripts/render_docx_pdf.py`，使用 Microsoft Word COM 导出 PDF，不依赖 LibreOffice；如安装 `pypdfium2`，还会把 PDF 渲染为逐页 PNG 便于视觉检查。

安装可选依赖：

```powershell
python -m pip install -e ".[render]"
```

导出 PDF 并渲染 PNG：

```powershell
python scripts/render_docx_pdf.py examples/图表附加项颗粒度示例.docx -o temp/图表示例.pdf
```

只导出 PDF：

```powershell
python scripts/render_docx_pdf.py examples/图表附加项颗粒度示例.docx -o temp/图表示例.pdf --no-png
```

## Markdown 契约

本节只描述生成器能够稳定解析和渲染的输入格式，不作为标准内容审校清单。交付前的审校步骤应由 `standard-audit-workflow` 完成。

Markdown 文件应以 YAML front matter 开始，封面和前言信息都从 front matter 读取：

```yaml
---
standard_type: 团体标准
number: T/XXXX XXX-XXXX
title: 标准中文名称
title_en: English title
consistency_degree: MOD
draft_version: （征求意见稿）
ics: "27.010"
ccs: D10
record_number: XXXX-XXXX
publish_date: XXXX-XX-XX
implement_date: XXXX-XX-XX
publisher: 发布单位
important_notice: 涉及人身安全或健康的整体提示。
symbols_lead: 下列符号和缩略语适用于本文件。
odd_even_pages: false
cover_form_protection: false
foreword:
  patent_note: true
  proposer: 提出单位
  owner: 归口单位
  draft_orgs:
    - 起草单位一
  drafters:
    - 起草人一
introduction: |
  引言内容。
---
```

封面草案版次可通过 `draft_version` 固定到旧式下拉框“下拉1”，常用值包括 `草案版次选择`、`（工作组讨论稿）`、`（征求意见稿）`、`（送审讨论稿）`、`（送审稿）`、`（报批稿）`；也可写不带括号的简称，如 `征求意见稿`。`cover_form_protection: true` 与命令行 `--cover-form-protection` 等效；命令行显式传入 `--cover-form-protection` 或 `--no-cover-form-protection` 时优先于 YAML。

`consistency_degree` 会生成在封面英文译名下方；若只写 `MOD`、`IDT` 等内容，生成时自动加全角圆括号。`record_number` 用于存在“备案号”占位的封面。`important_notice` 生成在正文首页标准名称之后、`# 范围` 之前；若未写 `重要提示：`、`危险：`、`警告：` 或 `注意：` 前缀，生成时自动补 `重要提示：`。`symbols_lead` 用于覆盖“符号和缩略语”章的默认导语。

正文建议按标准章节顺序编写：

```md
# 范围

# 规范性引用文件

# 术语和定义

# 技术章标题

# 附录 资料性 附录标题

# 参考文献

# 索引
```

关键规则：

- 不要在 Markdown 中手写章条编号、表号、图号、公式号。
- 不要使用旧交叉引用语法 `{@id}`。
- 不要使用 LaTeX `\tag{...}`。
- 表、图、公式锚点分别使用 `#tbl:id`、`#fig:id`、`#eq:id`。
- 交叉引用使用 `{{tbl:id}}`、`{{fig:id:label}}`、`{{eq:id:label}}`、`{{std:GB/T 11615}}`。
- 无标题条使用 `{无标题条:2}`、`{无标题条:3}`、`{无标题条:4}`。
- 行内上标使用 `^...^`，如下 `2^10^`；行内下标使用 `~...~`，如下 `H~2~O`。该写法与 Typora、Obsidian 等扩展 Markdown 习惯一致。也兼容 `<sup>...</sup>`、`<sub>...</sub>`，但推荐使用 `^...^`、`~...~`。
- 正文中的数学变量需要和公式字体一致时，使用行内公式 `$$...$$`，如 `$$T_r$$`、`$$Q_e$$`。不要用行内公式写长句或手写公式编号。

术语示例：

```md
{术语：地热温泉 | geothermal hot spring}

由地球内部热源加热，自然出露于地表或经人工工程揭露的地下热水。

注：本文件中的“地热温泉”不等同于生活饮用水水源。

[来源：GB/T 11615—2010，3.1，有修改]
```

兼容旧式写法：

```md
## 地热温泉  geothermal hot spring

由地球内部热源加热，自然出露于地表或经人工工程揭露的地下热水。
```

表格示例：

```md
{表：#tbl:classify} 温泉利用分类

| 利用类别 | 适宜水温 |
| --- | --- |
| 医疗保健洗浴〔注：适宜水温为常用建议范围。〕〔注：具体水温应结合安全要求调节。〕 | 36-45 |
```

复杂 HTML 表格可使用 `rowspan`、`colspan`，也可用 HTML 属性控制边框：

```md
{表：#tbl:merge} 合并和边框示例

<table data-border-outer="thick" data-border-inner="thin">
  <tr>
    <th rowspan="2" data-border-right="thick">类别</th>
    <th colspan="2">指标</th>
  </tr>
  <tr>
    <th>值</th>
    <th data-border-bottom="none">备注</th>
  </tr>
  <tr>
    <td>一类</td>
    <td></td>
    <td>同上</td>
  </tr>
</table>
```

`data-border-outer`、`data-border-inner` 和单元格上的 `data-border-top/right/bottom/left` 支持 `none`、`thin`、`thick`。空单元格保持为空白；“同上”不做自动推断，按普通文本输出。

图表通用附加项应紧跟在目标表格或图片之后，目标由上一张表或图自动确定：

```md
{单位} 单位为毫米

{脚注} 脚注的内容。

{来源} 资料来自试验记录。
```

表格中的普通注写在相关单元格内，使用 `〔注：...〕`；单元格可以只有注，例如 `<td colspan="3">〔注：...〕</td>`。同一单元格连续多个 `〔注：...〕` 会自动生成多条编号注，单条注使用不编号的注样式。需要在表格单元格中标出脚注引用点时，使用 `〔脚注〕`；脚注说明使用块级 `{脚注} ...`，编号由 `标准文件_图表脚注` 样式自动生成，不在 Markdown 中手写 a、b、c。

图片示例：

```md
![分级流程图 {#fig:flow}](images/flow.png)
```

图片附加项也使用同一套通用标记：

```md
{单位} 单位为毫米

{脚注} 图脚注的内容。

{来源} 资料来自流程设计文件。
```

带分图、标引序号说明、图内段和图脚注的图片示例：

```md
{图：#fig:subparts} 分图示例

{单位} 关于单位的陈述

{分图组:2}

![第一张分图题](images/subfigure-a.png)

![第二张分图题](images/subfigure-b.png)

{图标引} 说明的内容

{图标引} 说明的内容

{图段} 段（可包含要求型条款）〔注：图中的注的内容〕

{脚注} 图脚注的内容
```

普通单图继续使用 `![图题 {#fig:id}](path)`；组合分图使用 `{图：#fig:id} 图题` 创建图题和锚点，再用 `{分图组:2}` 声明每行并排显示 2 张分图，后续连续 Markdown 图片块会归入该父图。分图编号 `a)`、`b)` 由生成器输出，不应预先画进图片。CLI 转换时，图片相对路径以输入 Markdown 文件所在目录为基准。

`{单位}`、`{脚注}`、`{来源}` 只能绑定到紧邻的上一张表或图；`{分图组:2}`、`{图标引}`、`{图段}` 只能绑定到紧邻的上一张图。生成 DOCX 时，表的单位位于表题与表格之间并右对齐，表脚注另起跨列行，表来源位于表脚注之后另起跨列行；图的单位位于图片之前并右对齐，分图组、图标引、图段、图内段注、图脚注和图来源位于图片与图题之间。附加项和注内容支持加粗、斜体、上下标、行内公式和类型化交叉引用。

公式示例：

```md
按{{eq:depth:label}}计算：

$$H = \frac{T_{r} - T_{0}}{G} + h$${#eq:depth}

式中：

$$T_r$$ ——热储温度，单位为摄氏度（℃）；

m^3^/d ——立方米每天。
```

多块示例使用 `{示例结束}` 显式结束：

```md
示例：

第一段示例内容。

{表：#tbl:example} 示例表

| 项目 | 值 |
| --- | --- |
| A | 1 |

第二段示例内容。

{示例结束}
```

没有 `{示例结束}` 时，旧式单段示例写法继续可用。

## 示例

| 文件 | 说明 |
| --- | --- |
| `examples/基础要素与术语块改进示例.md` | 国家标准示例，集中展示封面一致性程度标识、正文首页重要提示、标准章导语自动补齐、`symbols_lead` 和显式术语块。 |
| `examples/图表附加项颗粒度示例.md` | 团体标准示例，按 GB/T 1.1 示例展示表单位、表内普通注、自动编号脚注和分图编排效果。 |
| `examples/表格合并与边框颗粒度示例.md` | 团体标准示例，展示 `rowspan`、`colspan`、空单元格、“同上”、单元格边框和多块示例。 |
| `examples/图表引用样式编号示例.md` | 团体标准示例，展示正文和附录图表题使用样式自动编号后的交叉引用效果。 |
| `examples/地热温泉资源开发利用规范.md` | 团体标准示例，覆盖术语、无标题条、表格、附录、参考文献和索引。 |
| `examples/汽车、摩托车用车速表.md` | 国家标准示例，覆盖国家标准封面、公式、HTML 表格和交叉引用。 |

## 验证

运行测试：

```powershell
python -X utf8 -m pytest -q
```

构建 wheel 并确认包内模板资源：

```powershell
python -m pip wheel . -w dist --no-deps
```

转换示例：

```powershell
python -X utf8 -m md2std examples/基础要素与术语块改进示例.md -o examples/基础要素与术语块改进示例.docx --kind national
python -X utf8 -m md2std examples/图表附加项颗粒度示例.md -o examples/图表附加项颗粒度示例.docx
python -X utf8 -m md2std examples/表格合并与边框颗粒度示例.md -o examples/表格合并与边框颗粒度示例.docx
python -X utf8 -m md2std examples/图表引用样式编号示例.md -o examples/图表引用样式编号示例.docx --word-com-postprocess
python -X utf8 -m md2std examples/地热温泉资源开发利用规范.md -o examples/地热温泉资源开发利用规范.docx
md2std examples/汽车、摩托车用车速表.md -o examples/汽车、摩托车用车速表.docx --kind national
```

## 仓库结构

```text
md2std/
  cli.py              命令行入口
  md_parser.py        Markdown 解析
  model.py            中间数据模型
  docx_builder.py     DOCX 构建兼容入口
  docx/               DOCX 内部生成模块
  resources.py        包内资源路径解析
  styles.py           Word 样式映射
  mathconv.py         LaTeX/MathML/OMML 转换
  word_postprocess.py Word COM 后处理
  templates/          封面蓝图与版式资产
examples/             示例 Markdown 与生成结果
tests/                自动化测试
```

## 工具边界

- `md2std` / `MdToStandardWord`：只负责 Markdown 到 DOCX 的排版生成。
- `md2std-standard-skill`：只负责生成器输入说明和 `scripts/run_md2std.py` 调用包装。
- `standard-drafting-workflow`：负责标准编写流程，并将审校任务委托给审校 skill。
- `standard-audit-workflow`：负责机械校验和 AI 清单审查。
- `md2std-standard-auditor`：负责确定性 Markdown 审校规则。

Skill 仓库不复制 CLI 源码，而是在 `scripts/requirements.txt` 中固定引用本仓库的 release tag：

```text
md2std @ git+https://github.com/LoyDgIk/MdToStandardWord.git@v0.1.0
```

发布新 CLI 版本时，先在本仓库提交并打 tag，再到 Skill 仓库更新依赖 tag。
