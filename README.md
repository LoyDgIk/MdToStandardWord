# MdToStandardWord（md2std）

`md2std` 是中文标准 Markdown 到 Word `.docx` 的排版生成器。它负责解析约定格式的 Markdown，生成符合标准文本结构的 Word 文档，支持团体标准和国家标准的封面、前言、引言、目次、正文、附录、参考文献、索引、图表公式题注和交叉引用。

本仓库只维护生成器本体和输入格式契约，不负责标准内容起草和审校。相关仓库：

- 标准编写、审校和生成技能：[standard-workflow-skills](https://github.com/LoyDgIk/standard-workflow-skills)
- 机械审校器：[md2std-standard-auditor](https://github.com/LoyDgIk/md2std-standard-auditor)

## 功能范围

- 根据 Markdown 和 YAML front matter 生成标准文本 `.docx`。
- 填充团体标准、国家标准封面字段。
- 生成前言、引言、目次、正文、附录、参考文献和索引。
- 支持规范性引用文件、术语和定义、符号和缩略语章的常用导语。
- 支持显式术语块、无标题条、列项、注、示例和分页符。
- 支持 GFM 表格和常用 HTML 表格；HTML 表格可设置合并单元格和边框。
- 支持图表单位、来源、脚注、表内注、图内段和组合分图。
- 支持块级公式、行内公式、上标、下标和公式符号说明。
- 支持表、图、公式、规范性引用文件的类型化交叉引用。
- 可选调用 Word COM 后处理，用于更新域、重新分页和处理续表。

`md2std` 不判断技术内容是否正确，也不判断标准是否完整符合 GB/T 1.1。交付前应使用 `md2std-standard-auditor` 和标准审校工作流检查 Markdown。

## 环境要求

- Python 3.9 或更高版本。
- Windows、macOS、Linux 均可执行基础转换。
- Word COM 后处理仅适用于安装 Microsoft Word 的 Windows。
- 公式转换依赖 `latex2mathml` 和 Office 的 `MML2OMML.XSL`；缺失时公式会回退为文本输出。

## 安装

开发安装：

```powershell
python -m pip install -e .
```

从 GitHub tag 安装：

```powershell
python -m pip install "md2std @ git+https://github.com/LoyDgIk/MdToStandardWord.git@v0.1.2"
```

可选安装 PDF 渲染依赖：

```powershell
python -m pip install -e ".[render]"
```

主要依赖包括 `python-docx`、`markdown-it-py`、`PyYAML`、`latex2mathml` 和 `lxml`。

## 基本用法

模块入口：

```powershell
python -X utf8 -m md2std examples/地热温泉资源开发利用规范.md -o examples/地热温泉资源开发利用规范.docx
```

安装后也可使用命令行入口：

```powershell
md2std examples/地热温泉资源开发利用规范.md -o examples/地热温泉资源开发利用规范.docx
```

指定国家标准：

```powershell
md2std examples/汽车、摩托车用车速表.md -o examples/汽车、摩托车用车速表.docx --kind national
```

启用 Word COM 后处理：

```powershell
md2std input.md -o output.docx --word-com-postprocess
```

常用参数：

| 参数 | 说明 |
| --- | --- |
| `input` | 输入 Markdown 文件路径。 |
| `-o, --output` | 输出 `.docx` 文件路径；未指定时输出到输入文件同名 `.docx`。 |
| `--kind {auto,group,national}` | 标准类型；默认 `auto`，根据 `standard_type` 或标准编号推断。 |
| `--word-com-postprocess`, `--word-com` | 调用本机 Microsoft Word 更新域、重新分页并保存。 |
| `--cover-form-protection` | 启用封面旧式下拉框表单保护；只保护封面节。 |
| `--no-cover-form-protection` | 关闭封面表单保护，用于覆盖 YAML 配置。 |

未启用 Word COM 后处理时，首次在 Word 中打开文档如提示更新域，应选择更新；也可全选正文后按 `F9` 手动更新。

## PDF 导出和页面渲染

Windows 环境下可使用 `scripts/render_docx_pdf.py` 通过 Microsoft Word COM 导出 PDF。安装 `pypdfium2` 后，脚本还可把 PDF 渲染为逐页 PNG。

导出 PDF 并渲染 PNG：

```powershell
python scripts/render_docx_pdf.py examples/图表附加项颗粒度示例.docx -o temp/图表示例.pdf
```

只导出 PDF：

```powershell
python scripts/render_docx_pdf.py examples/图表附加项颗粒度示例.docx -o temp/图表示例.pdf --no-png
```

## Markdown 输入契约

本节说明生成器能稳定解析的 Markdown 写法。它不是审校清单。

### YAML front matter

Markdown 文件应以 YAML front matter 开始：

```yaml
---
standard_type: 团体标准
number: T/XXXX XXX—XXXX
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

说明：

- `draft_version` 用于设置封面草案版次，可写 `（征求意见稿）`、`（送审稿）` 等。
- `cover_form_protection: true` 等效于命令行 `--cover-form-protection`；命令行参数优先级更高。
- `consistency_degree` 用于一致性程度标识；写 `MOD`、`IDT` 等内容时，生成器会补全括号。
- `record_number` 用于封面备案号。
- `important_notice` 生成在正文首页标准名称之后、`# 范围` 之前。
- `symbols_lead` 用于覆盖“符号和缩略语”章默认导语。

### 章节结构

正文建议按以下顺序编写：

```md
# 范围

# 规范性引用文件

# 术语和定义

# 技术章标题

# 附录 资料性 附录标题

# 参考文献

# 索引
```

基本规则：

- 不手写章条编号、表号、图号、公式号。
- 不使用旧交叉引用语法 `{@id}`。
- 不使用 LaTeX `\tag{...}`。
- 表、图、公式锚点分别使用 `#tbl:id`、`#fig:id`、`#eq:id`。
- 交叉引用使用 `{{tbl:id}}`、`{{fig:id:label}}`、`{{eq:id:label}}`、`{{std:GB/T 11615}}`。
- `# 规范性引用文件` 清单推荐行首显式注册：`{{std:GB/T 11615}} GB/T 11615  地热资源地质勘查规范`；外文标准可写 `{{std:ISO 3160-2:2015}} ISO 3160-2:2015  ...`；旧版自动识别仅作为兼容回退并会输出警告。
- `# 范围`、`# 规范性引用文件`、`# 术语和定义`、`# 符号和缩略语` 章内不使用无标题条标记，直接写普通段落或术语块。
- 无标题条只用于“术语和定义”之后的技术章或附录，写成 `{无标题条:2}`、`{无标题条:3}`、`{无标题条:4}`。
- 标准编号中的年份连接号应写一字线 `—`，例如 `GB/T 1.1—2020`。
- 参考文献条目不手写 `[1]`、`[2]`；最终编号由 Word 样式生成。

### 术语

推荐使用显式术语块：

```md
{术语：地热温泉 | geothermal hot spring}

由地球内部热源加热，自然出露于地表或经人工工程揭露的地下热水。

注：本文件中的“地热温泉”不等同于生活饮用水水源。

[来源：GB/T 11615—2010，3.1，有修改]
```

普通英文对应词不要加斜体；拉丁学名、生物分类学名称等需要斜体时，在英文对应词中使用 Markdown 斜体：

```md
{术语：大肠埃希氏菌 | *Escherichia coli*}
```

兼容旧式二级标题写法：

```md
## 地热温泉  geothermal hot spring

由地球内部热源加热，自然出露于地表或经人工工程揭露的地下热水。
```

### 表格

表题应紧贴表格：

```md
{表：#tbl:classify} 温泉利用分类

| 利用类别 | 适宜水温 |
| --- | --- |
| 医疗保健洗浴〔注：适宜水温为常用建议范围。〕 | 36—45 |

{单位} 单位为摄氏度

{脚注} 表脚注内容。

{来源} 资料来自试验记录。
```

复杂表格可使用 HTML：

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
    <td data-align="left">同上</td>
  </tr>
</table>
```

`data-border-outer`、`data-border-inner` 和单元格级 `data-border-top/right/bottom/left` 支持 `none`、`thin`、`thick`。
表头默认居中，正文短值、数字和代码默认居中，较长说明性文字默认左对齐。HTML 单元格可用 `data-align="left|center|right|decimal"` 显式覆盖；`decimal` 作为数字列控制，当前按右对齐输出。

### 图片和分图

普通单图：

```md
![分级流程图 {#fig:flow}](images/flow.png)
```

组合分图：

```md
{图：#fig:subparts} 分图示例

{单位} 单位为毫米

{分图组:2}

![第一张分图题](images/subfigure-a.png)

![第二张分图题](images/subfigure-b.png)

{图标引} 说明的内容

{图段} 图内段落内容〔注：图中的注的内容〕

{脚注} 图脚注内容。

{来源} 资料来自流程设计文件。
```

`{单位}`、`{脚注}`、`{来源}` 绑定到紧邻的上一张表或图；`{分图组:n}`、`{图标引}`、`{图段}` 只绑定到紧邻的上一张图。

### 公式、上标和下标

块级公式：

```md
按{{eq:depth:label}}计算：

$$H = \frac{T_{r} - T_{0}}{G} + h$${#eq:depth}

式中：

$$T_r$$ ——热储温度，单位为摄氏度（℃）；

m^3^/d ——立方米每天。
```

行内上标使用 `^...^`，如下 `2^10^`；行内下标使用 `~...~`，如下 `H~2~O`。正文、表格单元格、图内段、来源和脚注中的数学变量如需保持公式字体，可写行内公式 `$$T_r$$`。

### 示例

多块示例使用 `{示例结束}`：

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

## 示例文件

| 文件 | 说明 |
| --- | --- |
| `examples/基础要素与术语块改进示例.md` | 国家标准示例，展示封面字段、正文首页重要提示、章导语和显式术语块。 |
| `examples/图表附加项颗粒度示例.md` | 团体标准示例，展示表单位、表内注、脚注和分图。 |
| `examples/表格合并与边框颗粒度示例.md` | 团体标准示例，展示 HTML 表格合并、边框和多块示例。 |
| `examples/图表引用样式编号示例.md` | 团体标准示例，展示正文和附录图表交叉引用。 |
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

## 项目边界

| 项目 | 职责 |
| --- | --- |
| `MdToStandardWord` / `md2std` | Markdown 到 DOCX 的排版生成。 |
| `md2std-standard-auditor` | Markdown 机械审校，输出审校报告。 |
| `standard-workflow-skills/md2std-standard` | 生成器输入说明和调用脚本。 |
| `standard-workflow-skills/standard-audit-workflow` | 组织机械审校和清单审查。 |
| `standard-workflow-skills/standard-drafting-workflow` | 组织资料检索、标准编写、审校修订和交付。 |

发布新版本时，先在本仓库提交并打 tag，再更新 `standard-workflow-skills/md2std-standard/scripts/requirements.txt` 中引用的 tag。
