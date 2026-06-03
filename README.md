# MdToStandardWord

MdToStandardWord（`md2std`）是用于将 Markdown 标准草稿转换为标准文本 Word 文档的命令行工具。项目面向符合 GB/T 1.1—2020 起草规则的中文标准文档，支持团体标准和国家标准的封面蓝图、前言、引言、目次、正文、附录、参考文献、索引、图表公式题注以及交叉引用。

默认生成流程使用封面蓝图资产生成封面与基础版式，正文、附录、参考文献和索引由程序按标准结构写入。图、表、公式编号采用 Word `SEQ` 域，交叉引用采用 Word `REF` 域，便于在 Word 中更新域后获得可跳转的引用。

## 功能范围

- 根据 Markdown 和 YAML front matter 生成标准文本 `.docx`。
- 支持团体标准、国家标准封面字段填充。
- 支持前言、引言、范围、规范性引用文件、术语和定义、正文、附录、参考文献、索引等标准结构。
- 支持有标题条、无标题条、破折号列项、有序列项、注、示例等常见写法。
- 支持 GFM 表格和简单 HTML 表格，HTML 表格可保留横向合并单元格。
- 支持表格单元格内公式片段、正文公式块，并转换为 Word 原生公式。
- 支持表、图、公式、规范性引用文件的类型化交叉引用。
- 支持显式分页符和可选 Word COM 后处理；后处理可更新域、重新分页并按真实分页处理续表。
- 提供项目内 Codex Skill，用于辅助编写标准 Markdown 和执行转换流程。

## 环境要求

- Python 3.9 或更高版本。
- Windows、macOS、Linux 均可执行基础转换。
- Word COM 后处理仅适用于安装 Microsoft Word 的 Windows 环境。
- 公式转换依赖 `latex2mathml`，并使用 Office 自带 `MML2OMML.XSL`；缺失时公式会回退为文本输出。

## 安装

```powershell
pip install -r requirements.txt
```

主要依赖如下：

- `python-docx`
- `markdown-it-py`
- `PyYAML`
- `latex2mathml`

## 基本用法

在项目根目录执行：

```powershell
python -X utf8 -m md2std examples/地热温泉资源开发利用规范.md -o examples/地热温泉资源开发利用规范.docx
```

生成国家标准文档时，可显式指定标准类型：

```powershell
python -X utf8 -m md2std examples/汽车、摩托车用车速表.md -o examples/汽车、摩托车用车速表.docx --kind national
```

常用参数：

| 参数 | 说明 |
| --- | --- |
| `input` | 输入 Markdown 文件路径。 |
| `-o, --output` | 输出 `.docx` 文件路径；未指定时输出到输入文件同名 `.docx`。 |
| `--kind {auto,group,national}` | 标准类型；默认 `auto`，根据 `standard_type` 或标准编号推断。 |
| `--word-com-postprocess`, `--word-com` | 调用本机 Microsoft Word COM 更新域、重新分页并保存。 |

生成后的文档首次在 Word 中打开时，如提示是否更新域，应选择更新。未启用 Word COM 后处理时，也可在 Word 中全选正文后按 `F9` 更新域。

## Word COM 后处理

启用后处理：

```powershell
python -X utf8 -m md2std input.md -o output.docx --word-com-postprocess
```

该步骤会创建独立 Word 实例，打开同目录临时副本，更新域、重新分页、保存，并在成功后替换目标文件。程序不复用用户已打开的 Word 实例；超时清理仅针对本次创建的 Word 进程。

建议仅在最终交付前启用该参数，尤其适用于需要按 Word 实际分页生成续表题的文档。

## Markdown 文件结构

Markdown 文件应以 YAML front matter 开始，其后按标准章节顺序编写正文。建议结构如下：

```md
---
standard_type: 团体标准
number: T/XXXX XXX—XXXX
title: 标准中文名称
title_en: English title
ics: "27.010"
ccs: D10
publish_date: XXXX-XX-XX
implement_date: XXXX-XX-XX
publisher: 发布单位
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

# 范围

# 规范性引用文件

# 术语和定义

# 技术章标题

# 附录 资料性 附录标题

# 参考文献

# 索引
```

`ics`、含前导零的编号、日期占位符等建议写为字符串，避免 YAML 按数字解析。

## Front Matter 字段

| 字段 | 是否必需 | 说明 |
| --- | --- | --- |
| `standard_type` | 是 | 标准类型，如 `团体标准`、`国家标准`。 |
| `number` | 是 | 标准编号。 |
| `replaces` | 否 | 被代替标准编号。 |
| `title` | 是 | 标准中文名称。 |
| `title_en` | 否 | 标准英文名称。 |
| `ics` | 否 | ICS 分类号。 |
| `ccs` | 否 | CCS 分类号。 |
| `publish_date` | 否 | 发布日期。 |
| `implement_date` | 否 | 实施日期。 |
| `publisher` | 否 | 发布单位。 |
| `odd_even_pages` | 否 | 是否启用奇偶页页眉页脚，默认 `false`。 |
| `foreword` | 否 | 前言自动生成信息。 |
| `introduction` | 否 | 引言正文；缺省时不生成引言。 |

`foreword.extra_notes` 支持嵌套破折号列项：

```yaml
extra_notes:
  - 本文件及其所代替文件的历次版本发布情况为：
  - - 1994年首次发布为GB 15082—1994；
    - 本次为第三次修订。
```

## 正文标题与条款

标题不得手写编号，编号由 Word 样式和程序逻辑生成。

| Markdown | 用途 | 输出编号示例 |
| --- | --- | --- |
| `# 范围` | 章标题 | `1` |
| `## 一般要求` | 一级有标题条 | `5.1` |
| `### 子要求` | 二级有标题条 | `5.1.1` |
| `#### 更细要求` | 三级有标题条 | `5.1.1.1` |
| `{无标题条:2} 正文` | 一级无标题条 | `X.Y` |
| `{无标题条:3} 正文` | 二级无标题条 | `X.Y.Z` |
| `{无标题条:4} 正文` | 三级无标题条 | `X.Y.Z.W` |

旧式手写无标题条编号，如 `4.2.1 正文`，会作为普通段落处理并给出警告。正式文档应使用 `{无标题条:n}`。

显式分页符应独占一行：

```md
<!-- pagebreak -->
```

也支持 `<!-- md2std:pagebreak -->`、`\pagebreak`、`\newpage`、`[pagebreak]`。

## 列项、注和示例

破折号列项使用普通 Markdown 无序列表：

```md
- 第一项；
- 第二项。
```

有序列项使用 Markdown 有序列表。一级有序列项渲染为 `a)`、`b)`，嵌套有序列项渲染为 `1)`、`2)`：

```md
1. 一级列项。
2. 一级列项：
   1. 二级列项；
   2. 二级列项。
```

注和示例直接使用标准前缀：

```md
注：这是注。
注1：这是编号注。
示例：这是示例。
示例1：这是编号示例。
```

## 术语和定义

在 `# 术语和定义` 章内，每条术语使用二级标题。中文术语和英文对应词之间使用两个空格分隔：

```md
## 地热温泉  geothermal hot spring

由地球内部热源加热，自然出露于地表或经人工钻井揭露的地下热水。
```

术语编号由程序自动生成，中文术语和英文对应词均按术语样式加粗。

## 规范性引用文件

每个标准条目单独成段：

```md
GB 5749  生活饮用水卫生标准

GB/T 11615  地热资源地质勘查规范
```

正文中引用标准号时使用：

```md
各阶段工作应符合{{std:GB/T 11615}}的规定。
```

`{{std:...}}` 中的标准号应与规范性引用文件章中的标准号完全一致。

无规范性引用文件时写作：

```md
# 规范性引用文件

本文件没有规范性引用文件。
```

## 表格

表题写在表格之前，使用类型化锚点：

```md
{表：#tbl:classify} 温泉利用分类

| 利用类别 | 适宜水温 |
| --- | --- |
| 医疗保健洗浴 | 36～45 |
```

表题只写标题，不写 `表1`、`表A.1` 等编号。正文表自动编号为 `表1`、`表2`；附录表自动编号为 `表A.1`、`表B.1`。

需要保留横向合并单元格时，可使用简单 HTML 表格：

```md
{表：#tbl:limit} 限值

<table>
  <tr><td>项目</td><td>限值</td></tr>
  <tr><td colspan="2">注：合并单元格说明。</td></tr>
</table>
```

表格单元格内可写公式片段：

```md
<eq>v_{\text{max}}</eq>
```

跨页续表不在生成阶段按行数预拆。启用 Word COM 后处理后，程序按 Word 实际分页位置处理续表题。

## 图

图使用 Markdown 图片语法，图题写在 alt 文本中，锚点使用 `#fig:id`：

```md
![分级流程图 {#fig:flow}](images/flow.png)
```

图片路径应相对于执行命令的工作目录，或使用绝对路径。图题不得手写 `图1`、`图A.1` 等编号。

## 公式

公式使用 LaTeX 块，锚点写在公式块之后：

```md
按{{eq:depth:label}}计算：

$$H = \frac{T_{r} - T_{0}}{G} + h$${#eq:depth}

式中：

H ——循环深度，单位为米（m）；
```

公式不得使用 `\tag{...}`，不得手写可见编号。正文公式自动编号为 `（1）`、`（2）`；附录公式自动编号为 `（A.1）`、`（B.1）`。

## 交叉引用

表、图、公式使用双花括号语法。默认输出仅编号：

```md
{{tbl:classify}}        # 1 或 A.1
{{tbl:classify:label}}  # 表1 或 表A.1
{{tbl:classify:full}}   # 表1　温泉利用分类
{{fig:flow:label}}      # 图1
{{fig:flow:full}}       # 图1　分级流程图
{{eq:depth}}            # 1
{{eq:depth:label}}      # 式（1）
{{std:GB/T 11615}}      # GB/T 11615
```

锚点和引用均应带类型前缀：

- 表：`#tbl:id` 与 `{{tbl:id}}`
- 图：`#fig:id` 与 `{{fig:id}}`
- 公式：`#eq:id` 与 `{{eq:id}}`
- 标准号：`{{std:标准号}}`

旧语法 `{@id}`、`{@id:a}`、`{@id:b}` 不再使用。

## 附录

附录使用一级标题，格式为：

```md
# 附录 规范性 分级判定流程

## 判定步骤

……

# 附录 资料性 数据示例
```

附录字母由程序分配。附录内条款、表、图、公式按所在附录自动编号。

## 参考文献

参考文献为可选节。每条文献单独成段：

```md
# 参考文献

GB/T 1.1—2020　标准化工作导则 第1部分：标准化文件的结构和起草规则
```

## 索引

索引为可选节。按汉语拼音首字母分组，索引项使用 `术语：位置列表`：

```md
# 索引

## B

- 必备要素：3.2.5，6.2.2.1，6.2.2.3
- 标准：3.1.2，4.1，4.2
```

不要在 Markdown 中手写点引导线。

## 示例文件

项目提供两个完整示例：

| 文件 | 说明 |
| --- | --- |
| `examples/地热温泉资源开发利用规范.md` | 团体标准示例，覆盖术语、无标题条、表格、附录、参考文献和索引。 |
| `examples/汽车、摩托车用车速表.md` | 国家标准示例，覆盖国家标准封面、公式、HTML 表格和交叉引用。 |

项目内还提供 Codex Skill：

| 路径 | 用途 |
| --- | --- |
| `skills/md2std-standard/SKILL.md` | 说明本项目 Markdown 写作约定与 DOCX 转换流程。 |
| `skills/md2std-standard/examples/智能井盖运行维护规范.md` | Skill 内部独立示例，不依赖项目根目录示例。 |

## 验证

运行测试：

```powershell
python -X utf8 -m pytest -q
```

转换示例：

```powershell
python -X utf8 -m md2std examples/地热温泉资源开发利用规范.md -o examples/地热温泉资源开发利用规范.docx
python -X utf8 -m md2std examples/汽车、摩托车用车速表.md -o examples/汽车、摩托车用车速表.docx --kind national
```

最终交付前，可在 Windows Word 环境中启用 COM 后处理：

```powershell
python -X utf8 -m md2std input.md -o output.docx --word-com-postprocess
```

## 项目结构

```text
md2std/
  cli.py              命令行入口
  md_parser.py        Markdown 解析
  model.py            中间数据模型
  docx_builder.py     DOCX 构建
  styles.py           Word 样式映射
  mathconv.py         LaTeX/MathML/OMML 转换
  word_postprocess.py Word COM 后处理
templates/            封面蓝图与版式资产
examples/             示例 Markdown 与生成结果
tests/                自动化测试
skills/               项目内 Codex Skill
```

## 注意事项

- 文档编号、图题、表题、公式编号和交叉引用应由程序生成，不应在 Markdown 中手写。
- Word 域需要更新后才会显示最终编号和页码。
- Word COM 后处理可能触发本机 Word 安全提示或恢复提示，宜在最终生成阶段使用。
- 可选章节缺省时不生成对应空节。
