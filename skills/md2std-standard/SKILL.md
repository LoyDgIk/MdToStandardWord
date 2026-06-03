---
name: md2std-standard
description: Write, normalize, validate, and convert Chinese standard document Markdown for the md2std project. Use when Codex needs to draft GB/T-style standards, group standards, national standards, or md2std Markdown; migrate PDF/MinerU Markdown into the project syntax; fix table/figure/formula captions and cross references; or generate DOCX with python -m md2std. Use only the current cover-blueprint DOCX generation path.
---

# md2std Standard

## Core Rule

Use the current md2std Markdown contract and the default cover-blueprint DOCX generation path. Do not suggest alternate generation backends in user-facing output.

Never hand-write generated numbering in Markdown:

- Do not write heading numbers such as `4.2.1` in headings.
- Do not write caption labels such as `表1`、`图A.1`、`式（1）` in titles.
- Do not write LaTeX `\tag{...}`.
- Do not use old cross-reference syntax `{@id}`.

## Workflow

1. Inspect the source brief, PDF-converted Markdown, or example document before writing.
2. Create or normalize a single Markdown file with YAML front matter at the top.
3. Use typed anchors for every referenced table, figure, and equation.
4. Use `{{...}}` cross references only after the target anchor exists.
5. Run md2std from the project root and fix parser warnings/errors before delivering the DOCX.
6. Use Word COM post-processing only when final Word pagination, field update, and continuation-table splitting are required and Windows Microsoft Word is available.

For a compact complete sample, read `examples/智能井盖运行维护规范.md` inside this skill.

## Front Matter

Use YAML front matter as the only source for cover and boilerplate metadata:

```yaml
---
standard_type: 团体标准
number: T/XXXX XXX—XXXX
replaces: T/XXXX XXX—2020
title: 标准中文名称
title_en: English title
ics: "27.010"
ccs: D10
publish_date: XXXX-XX-XX
implement_date: XXXX-XX-XX
publisher: 发布单位
odd_even_pages: false
foreword:
  patent_note: true
  proposer: 提出单位
  owner: 归口单位
  draft_orgs:
    - 起草单位一
  drafters:
    - 起草人一
  replace_changes:
    - 更改了……
  extra_notes:
    - 本文件为首次发布。
introduction: |
  引言第一段。
---
```

Use `standard_type: 国家标准` or a `GB...` standard number for national standards. Quote ICS values and other values that may be parsed as numbers.

`foreword.extra_notes` supports nested dash items:

```yaml
extra_notes:
  - 本文件及其所代替文件的历次版本发布情况为：
  - - 1994年首次发布为GB 15082—1994；
    - 本次为第三次修订。
```

## Body Structure

Write chapters in standard order when applicable:

```md
# 范围
# 规范性引用文件
# 术语和定义
# 符号和缩略语
# 技术章标题
# 附录 规范性 附录标题
# 参考文献
# 索引
```

Use Markdown heading depth for titled clauses and let md2std number them:

```md
# 范围
## 一般要求
### 子要求
#### 更细要求
```

For no-title clauses, use explicit level syntax instead of writing the visible number:

```md
{无标题条:2} 这是 X.Y 级无标题条正文。
{无标题条:3} 这是 X.Y.Z 级无标题条正文。
{无标题条:4} 这是 X.Y.Z.W 级无标题条正文。
```

If source text contains old manual no-title clauses such as `4.2.1 正文`, keep them as ordinary text only if preserving source evidence is necessary; otherwise convert them to `{无标题条:n}`.

Use a standalone page break marker when a standard requires an explicit new page:

```md
<!-- pagebreak -->
```

Also accepted: `<!-- md2std:pagebreak -->`、`\pagebreak`、`\newpage`、`[pagebreak]`.

## Lists, Notes, and Examples

Use plain Markdown list syntax. Let Word numbering styles render the final list labels:

```md
- 破折号列项；
- 另一个破折号列项。

1. 一级有序列项，渲染为 a)。
2. 另一个一级有序列项：
   1. 二级有序列项，渲染为 1)。
   2. 另一个二级有序列项。
```

Use standard note/example prefixes directly:

```md
注：这是注。
注1：这是编号注。
示例：这是示例。
示例1：这是编号示例。
```

## Terms

In `# 术语和定义`, write each term as a second-level heading. Separate Chinese and English terms with two spaces:

```md
## 地热温泉  geothermal hot spring

由地球内部热源加热……
```

Do not include the generated term number in Markdown. The renderer makes both the Chinese and English term bold.

## Normative References

In `# 规范性引用文件`, write one standard per paragraph:

```md
GB 5749  生活饮用水卫生标准
GB/T 11615  地热资源地质勘查规范
```

Reference standard numbers with exact text matching:

```md
各阶段工作应符合{{std:GB/T 11615}}的规定。
```

If there are no normative references, write:

```md
# 规范性引用文件

本文件没有规范性引用文件。
```

## Tables

Put the table caption marker immediately before the table. The caption title must be pure title text:

```md
单位为千米每小时

{表：#tbl:test-speed} 测试车速

| 最高设计车速 | 测试车速 |
| --- | --- |
| ≤45 | 80%最高设计车速 |
```

Rules:

- Use typed IDs: `#tbl:id`.
- Do not write `表1` or `表A.1` in the caption text.
- GFM tables and simple HTML `<table>` are supported.
- Use HTML tables when source tables need `colspan`.
- Use `<eq>...</eq>` or `$...$` for formulas inside table cells.
- Continuation captions such as `表1　原表题（续）` are generated only by Word COM post-processing after real pagination.

HTML table example:

```md
{表：#tbl:limit} 限值

<table>
  <tr><td>项目</td><td>限值</td></tr>
  <tr><td colspan="2">注：合并单元格说明。</td></tr>
</table>
```

## Figures

Use Markdown image syntax with a typed anchor in the alt text:

```md
![分级流程图 {#fig:flow}](images/flow.png)
```

Rules:

- Use typed IDs: `#fig:id`.
- Do not write `图1` or `图A.1` in the figure title.
- Resolve image paths relative to the Markdown file or project root when testing.

## Equations

Use LaTeX display math with a typed anchor after the closing `$$`:

```md
按{{eq:depth:label}}计算：

$$H = \frac{T_{r} - T_{0}}{G} + h$${#eq:depth}

式中：

H ——循环深度，单位为米（m）；
```

Rules:

- Use typed IDs: `#eq:id`.
- Do not use `\tag{1}`.
- Do not write visible equation numbers manually.
- Use `式中：` and following symbol explanations as ordinary paragraphs.

## Cross References

Use only double-brace references:

```md
{{tbl:classify}}        # 仅编号，如 1 或 A.1
{{tbl:classify:label}}  # 表1 或 表A.1
{{tbl:classify:full}}   # 表1　温泉利用分类
{{fig:flow:label}}      # 图1
{{fig:flow:full}}       # 图1　分级流程图
{{eq:depth}}            # 1
{{eq:depth:label}}      # 式（1）
{{std:GB/T 11615}}      # GB/T 11615
```

Rules:

- Target IDs must exist and must match the reference type.
- IDs must be unique within their type.
- Use `:label` when the sentence needs the type name.
- Use `:full` when the sentence needs the complete caption.

## Appendices

Write appendix headings as first-level headings. Let md2std assign A, B, C:

```md
# 附录 规范性 分级判定流程

## 判定步骤

……

# 附录 资料性 数据示例
```

Appendix table, figure, and equation numbers automatically use appendix scopes such as `表A.1`、`图B.1`、`式（A.1）`.

## References and Index

Reference section is optional:

```md
# 参考文献

GB/T 1.1—2020　标准化工作导则……
ISO 11783　Tractors……
```

Index section is optional. Use pinyin initial groups and `术语：位置列表` entries:

```md
# 索引

## B

- 必备要素：3.2.5，6.2.2.1，6.2.2.3
- 标准：3.1.2，4.1，4.2
```

Do not hand-write dot leaders; Word styles generate the index layout.

## Conversion Commands

Run from the project root:

```powershell
python -X utf8 -m md2std examples/地热温泉资源开发利用规范.md -o examples/地热温泉资源开发利用规范.docx
```

Override standard kind only when auto-detection is not enough:

```powershell
python -X utf8 -m md2std input.md -o output.docx --kind national
python -X utf8 -m md2std input.md -o output.docx --kind group
```

Use Word COM post-processing only for final DOCX pagination/field work:

```powershell
python -X utf8 -m md2std input.md -o output.docx --word-com-postprocess
```

The post-process step requires Windows Microsoft Word and `pywin32`. It updates fields, repaginates, saves, and handles continuation tables based on actual Word pagination.

## Validation Checklist

Before delivering Markdown or DOCX, check:

- No `{@...}` old references.
- No visible generated numbers in headings, captions, or formulas.
- No LaTeX `\tag{...}`.
- Every `{{tbl:...}}`、`{{fig:...}}`、`{{eq:...}}`、`{{std:...}}` has a matching target.
- Table, figure, and equation IDs use `tbl:`、`fig:`、`eq:` typed anchors.
- Optional sections such as appendix, reference, and index can be absent without placeholder content.
- DOCX generation succeeds with `python -X utf8 -m md2std ...`.
