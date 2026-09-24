# academic-word-skill

[中文](#中文) | [English](#english)

## 中文

面向中文学位论文、开题报告和学术稿件的Codex技能，配套可复用的DOCX处理脚本。适用于文字润色、参考文献自动编号、题注与交叉引用、MathType对象保留，以及通过Microsoft Word进行排版核验。

核心原则是：**以用户最新保存的文档为底稿，保留手动修改，不覆盖原稿，另存新版本。** 本项目不是独立的Word插件，也不承诺自动完成任意文档的全部排版。

GitHub仓库名为`academic-word-skill`，技能调用名仍为`$academic-word`，本地技能目录保持`academic-word`。仓库重命名不需要重装或改变现有调用方式。

### 主要功能

- **学术文字处理**：润色和适度扩写，检查研究内容与技术路线等章节的一致性，不虚构实验、数据或参考文献。
- **参考文献与正文引用**：将符合条件的手打`[n]`转换为Word原生编号列表，以书签和`REF`域建立可更新的交叉引用。
- **题注与图号引用**：处理符合条件的原生`SEQ`题注，正文引用仅包含标签和编号；移动引导句时保留引用域。
- **公式与图件保护**：保留已有MathType可编辑OLE对象、图片和有关域，以字节内容及数量比较检查意外变化，保留公式基线与尺寸。
- **中文学术排版**：按需采用中英文间距、题注字体、真实空行、连续公式编号等偏好；这些不是强制适用于所有学校的格式规则。
- **流程图修订**：检查节点间距、箭头末段、回路和画布边缘；在Word中检查图题同页及表格内分页，不靠缩小整图解决拥挤。
- **Word核验**：通过独立Word COM实例选择性更新域、另存DOCX和导出PDF，按影响范围检查公式、题注、表格和分页。

### 安装与调用

#### 安装技能

本仓库已公开，可直接浏览、下载或通过HTTPS克隆，无需登录GitHub。下面的PowerShell示例需要已安装Git。

根据[Codex技能官方文档](https://learn.chatgpt.com/zh-Hans/docs/build-skills)，用户级技能可放在`$HOME/.agents/skills`，项目级技能可放在项目的`.agents/skills`中。以下示例安装到用户级目录，并拒绝覆盖已有安装。

```powershell
$skillRoot = Join-Path $env:USERPROFILE '.agents\skills'
$skillPath = Join-Path $skillRoot 'academic-word'
if (Test-Path -LiteralPath $skillPath) {
    throw '目标目录已存在，请先检查已有技能，不要覆盖或重复安装。'
}
New-Item -ItemType Directory -Path $skillRoot -Force | Out-Null
git clone https://github.com/Jerry-behappy/academic-word-skill.git $skillPath
```

本技能也已在使用`.codex/skills`目录的本地环境中使用。已有安装应沿用当前环境实际识别的目录，不要在两处重复安装同名技能。安装后若未出现，可重启Codex并检查技能列表。

#### 在对话中使用

提供文档的实际路径和具体修改范围，例如：

```text
$academic-word
以“D:\论文\开题报告.docx”为底稿，润色研究方法与技术路线。
保留我的手动修改、MathType公式和原有交叉引用，另存新版本，
并核验编号和排版，不覆盖原稿。
```

如果只需要文字建议，可以明确说明：

```text
$academic-word
只润色下面这段文字，暂不修改Word，不新增实验内容或性能指标。
```

运行前请先保存需要作为底稿的文件。技能不会擅自保存或关闭用户正在使用的Word文档。

### 环境依赖

| 用途 | 依赖 |
| --- | --- |
| DOCX结构检查与修改 | Python、`lxml` |
| 离线回归测试 | 额外需要`python-docx` |
| PDF转页面图片 | `pdf2image`、Pillow，以及单独安装的Poppler |
| Word COM更新、另存与导出 | Windows、PowerShell 7（推荐）、已安装的Microsoft Word桌面版 |
| 编辑或转换MathType公式 | 可用的MathType安装及经过验证的插件或API流程，不由现有脚本自动提供 |

优先使用当前环境已有的运行时；缺少Python依赖时，可在自建虚拟环境中安装：

```powershell
python -m pip install lxml python-docx pdf2image Pillow
```

Poppler不是上述pip命令的一部分。渲染前需确保`pdfinfo`和`pdftoppm`可以从命令行调用。

### 目录与脚本

| 文件 | 说明 |
| --- | --- |
| [SKILL.md](SKILL.md) | 技能入口、处理原则和交付要求 |
| [agents/openai.yaml](agents/openai.yaml) | 技能显示名称及默认调用提示 |
| [references/word-mechanics.md](references/word-mechanics.md) | Word编号、交叉引用、脚本命令和支持边界 |
| [references/chinese-format-profile.md](references/chinese-format-profile.md) | 可按需采用的中文学术排版偏好 |
| [references/flowcharts.md](references/flowcharts.md) | 流程图箭头、间距、算法一致性和Word内分页检查 |
| [scripts/docx_ops.py](scripts/docx_ops.py) | 检查、编号转换、题注处理、格式规范化、保留域的句子移动、对象审计 |
| [scripts/format_figure_ref_phrases.py](scripts/format_figure_ref_phrases.py) | 定点设置“如图 x 所示”完整短语的字号，保留原生交叉引用域 |
| [scripts/word_finalize.ps1](scripts/word_finalize.ps1) | Word COM选择性更新域、另存文档、导出PDF及编号核验 |
| [scripts/render_pdf.py](scripts/render_pdf.py) | 将已有PDF输出为逐页PNG、总览图和页面清单 |
| [scripts/test_docx_ops.py](scripts/test_docx_ops.py) | 离线结构与安全边界回归测试 |
| [scripts/test_word_fields.ps1](scripts/test_word_fields.ps1) | 临时Word文档中的原生编号与交叉引用测试 |

### 手动运行示例

以下命令在仓库根目录运行。`source.docx`是占位文件名，需替换为实际底稿；先阅读[脚本使用说明](references/word-mechanics.md)，不要直接对不符合条件的文档执行转换。

#### 检查文档

```powershell
New-Item -ItemType Directory -Path qa -Force | Out-Null
python scripts/docx_ops.py inspect source.docx --report qa/source.json
```

#### 转换符合条件的手打参考文献

```powershell
python scripts/docx_ops.py bibliography source.docx --output qa/numbered.docx --report qa/conversion.json
pwsh -NoProfile -File scripts/word_finalize.ps1 -InputPath qa/numbered.docx -OutputPath final.docx -PdfPath qa/final.pdf -ReportPath qa/fields.json -UpdateBibliography
python scripts/docx_ops.py audit-assets source.docx --compare final.docx
python scripts/render_pdf.py qa/final.pdf qa/pages
```

流程为：结构转换 → Word更新并另存 → 图片与嵌入对象审计 → 页面渲染与人工检查。输出文件必须使用尚不存在的新路径；重复运行时应更换输出名称，`qa/pages`也必须是尚不存在的目录。

#### 只导出PDF，不修改文档或更新域

```powershell
pwsh -NoProfile -File scripts/word_finalize.ps1 -InputPath source.docx -PdfPath qa/source.pdf
```

#### 图题注与引用

当前保存的个人偏好为`图 1 题注文字`，四号（14 pt）、中文宋体、英文及数字Times New Roman。本用户的开题报告还要求正文完整短语`如图 1所示`为四号，而不只修改其中的图号；其他文档以其模板或用户要求为准。

```powershell
python scripts/docx_ops.py figures source.docx --output qa/captions.docx --caption-size 14
pwsh -NoProfile -File scripts/word_finalize.ps1 -InputPath qa/captions.docx -OutputPath captions-final.docx -UpdateFigures
```

上述转换仅适用于尚无题注书签的简单原生SEQ题注。已有有效书签时不要重建，采用定点修改。`--label-separator none`可按需取消标签与编号之间的空格；`-FigureReferenceSize 12`仅在适用正文统一为12 pt时设置引用字号，否则逐处匹配。

已核对交叉引用域后，可单独设置完整正文短语的字号；`--expected-count`须替换为当前文档的实际匹配数，输出使用新的文件名：

```powershell
python scripts/format_figure_ref_phrases.py source.docx --output figure-refs.docx --size-pt 14 --expected-count 23
```

保留域的句子移动及可选格式规范化，请参阅[详细说明](references/word-mechanics.md)，或运行`python scripts/docx_ops.py --help`。

### 测试

```powershell
python -m unittest discover -s scripts -p "test_*.py" -v
pwsh -NoProfile -File scripts/test_word_fields.ps1
```

第一项是离线回归测试；第二项需要Windows和Microsoft Word，使用新建的合成文档核对列表编号、任意命名题注书签的引用更新、正文引用字号和锁定公式编号域保护。测试文件存放在新建的临时目录，也可用`-OutputDirectory`指定尚不存在的目录。测试通过不等于具体论文已完成科学内容核验或排版检查，也不等于测试了MathType插件内部编辑。

### 使用边界与注意事项

- **保留MathType不等于自动转换MathType。** Word原生OMML公式也不是MathType对象；需要转换时，应先在副本上验证可用的MathType流程。
- `bibliography`要求一个匹配的参考文献标题，后面紧接连续手打`[1]`至`[N]`的条目。已有自动编号、相关书签、修订记录或文献管理软件的ADDIN域等情况会触发保护检查。它不能自动区分文献引用与数学区间中的方括号数字，也不会自动按首次引用顺序重排文献。
- `figures`只处理尚未建立题注书签的简单原生`SEQ 图`题注，不负责将所有手打题注自动转换，也不判断图文含义是否对应。
- 不通过全局F9、批量解锁或将域转成普通文字来强行更新。Word COM脚本只处理选定范围内支持的域，不更新或解锁MathType、Zotero、EndNote管理的域。
- 当前检查和更新主要面向正文；页眉、页脚、文本框及复杂表格可能需要专门处理。段内公式周围的空格、公式编号和局部字体也可能需要定点修改。
- 图件和嵌入对象内容一致只能证明这些载荷未变，不能证明位置和排版正确。最终仍需检查受影响页面的内容遮挡、跨页边框、公式清晰度和引用显示。
- 文献原方法与项目调整应分开描述。例如直接在dBm数值上计算STD是具体项目选择，不应冒充线性功率STD的等价复现，也不是技能的全局默认算法。
- 该仓库只存放技能和通用脚本。处理具体文档时，请将论文、测试数据、导出的页面及其他敏感材料留在自己的工作目录，不要误提交到仓库。

---

## English

A Codex skill for Chinese-language theses, doctoral proposals, and academic manuscripts, with reusable DOCX utilities. It supports text polishing, native bibliography numbering, captions and cross-references, preservation of MathType objects, and layout verification through Microsoft Word.

The core principle is to **work from the user's latest saved document, preserve manual edits, and save a new version without overwriting the original.** This is not a standalone Word add-in, nor does it promise fully automatic formatting for arbitrary documents.

The GitHub repository is named `academic-word-skill`. The skill is still invoked as `$academic-word`, and its local folder remains `academic-word`. Renaming the repository does not require reinstalling the skill or changing existing prompts.

### Features

- **Academic writing:** Polish and moderately expand text, checking consistency across research content, methods, and technical plans without inventing experiments, data, or references.
- **Bibliographies and citations:** Convert supported manually typed `[n]` entries into native Word numbered lists, with bookmarks and `REF` fields for updateable citations.
- **Captions and figure references:** Work with supported native `SEQ` captions. Body references contain only the label and number; fields are preserved when introductory sentences are moved.
- **Equation and figure preservation:** Retain existing editable MathType OLE objects, images, and related fields. Compare payload bytes and counts to detect unintended changes while preserving equation baselines and dimensions.
- **Chinese academic formatting:** Apply optional preferences for Chinese–Latin spacing, caption fonts, actual blank paragraphs, and continuous equation numbering. These are not universal requirements for every institution.
- **Flowchart revisions:** Check node spacing, arrowheads and final connector segments, return paths, and canvas boundaries. Verify figure–caption placement and pagination inside Word tables without relying on excessive scaling.
- **Word verification:** Use a separate Word COM instance to selectively update fields, save a new DOCX, and export a PDF. Inspect equations, captions, tables, and pagination according to the scope of the changes.

### Installation and usage

#### Install the skill

This repository is public. You can browse, download, or clone it over HTTPS without signing in to GitHub. The PowerShell example below requires Git.

The [Codex skills documentation](https://learn.chatgpt.com/zh-Hans/docs/build-skills) describes user-level skills under `$HOME/.agents/skills` and project-level skills under `.agents/skills`. This example uses the user-level location and refuses to overwrite an existing installation.

```powershell
$skillRoot = Join-Path $env:USERPROFILE '.agents\skills'
$skillPath = Join-Path $skillRoot 'academic-word'
if (Test-Path -LiteralPath $skillPath) {
    throw 'The destination already exists. Inspect the installed skill before making changes.'
}
New-Item -ItemType Directory -Path $skillRoot -Force | Out-Null
git clone https://github.com/Jerry-behappy/academic-word-skill.git $skillPath
```

This skill has also been used in a local environment that discovers skills under `.codex/skills`. Keep an existing installation in the location recognized by your environment; do not install duplicate copies under both paths. If the skill does not appear after installation, restart Codex and check the available skills.

#### Invoke it in a conversation

Provide the actual document path and a specific editing scope, for example:

```text
$academic-word
Use "D:\thesis\proposal.docx" as the source and polish the research methods
and technical plan. Preserve my manual edits, MathType equations, and
existing cross-references. Save a new version and verify numbering and
layout without overwriting the original.
```

For text suggestions only, state that explicitly:

```text
$academic-word
Polish the following paragraph only. Do not modify the Word document
or add experiments or performance claims.
```

Save the intended source document before starting. The skill does not save or close documents in your active Word session without authorization.

### Requirements

| Purpose | Requirements |
| --- | --- |
| DOCX inspection and editing | Python and `lxml` |
| Offline regression tests | Also requires `python-docx` |
| PDF-to-image rendering | `pdf2image`, Pillow, and a separate Poppler installation |
| Word COM updates, saving, and export | Windows, PowerShell 7 (recommended), and desktop Microsoft Word |
| Editing or converting MathType equations | A working MathType installation and a verified add-in or API workflow; not provided automatically by these scripts |

Prefer runtimes already available in your environment. If Python dependencies are missing, install them in your own virtual environment:

```powershell
python -m pip install lxml python-docx pdf2image Pillow
```

Poppler is not included in this pip command. Before rendering, ensure that `pdfinfo` and `pdftoppm` are available from the command line.

### Files and scripts

| File | Purpose |
| --- | --- |
| [SKILL.md](SKILL.md) | Skill entry point, working principles, and delivery requirements |
| [agents/openai.yaml](agents/openai.yaml) | Display name and default invocation prompt |
| [references/word-mechanics.md](references/word-mechanics.md) | Word numbering, cross-references, script commands, and supported cases |
| [references/chinese-format-profile.md](references/chinese-format-profile.md) | Optional Chinese academic formatting preferences |
| [references/flowcharts.md](references/flowcharts.md) | Flowchart arrows, spacing, algorithm consistency, and pagination in Word |
| [scripts/docx_ops.py](scripts/docx_ops.py) | Inspection, numbering conversion, caption handling, optional formatting, field-preserving sentence moves, and asset audits |
| [scripts/format_figure_ref_phrases.py](scripts/format_figure_ref_phrases.py) | Set the size of complete “如图 x 所示” phrases while preserving native cross-reference fields |
| [scripts/word_finalize.ps1](scripts/word_finalize.ps1) | Selective Word COM field updates, saving, PDF export, and numbering checks |
| [scripts/render_pdf.py](scripts/render_pdf.py) | Render a PDF to page PNGs, overview sheets, and a page manifest |
| [scripts/test_docx_ops.py](scripts/test_docx_ops.py) | Offline regression tests for document structures and safety boundaries |
| [scripts/test_word_fields.ps1](scripts/test_word_fields.ps1) | Native numbering and cross-reference tests using synthetic Word documents |

The skill instructions and supporting reference guides are currently written in Chinese.

### Manual command examples

Run the following commands from the repository root. Replace `source.docx` with your actual source document. Read the [script guide](references/word-mechanics.md) before converting a document; unsupported structures require targeted handling.

#### Inspect a document

```powershell
New-Item -ItemType Directory -Path qa -Force | Out-Null
python scripts/docx_ops.py inspect source.docx --report qa/source.json
```

#### Convert a supported manually numbered bibliography

```powershell
python scripts/docx_ops.py bibliography source.docx --output qa/numbered.docx --report qa/conversion.json
pwsh -NoProfile -File scripts/word_finalize.ps1 -InputPath qa/numbered.docx -OutputPath final.docx -PdfPath qa/final.pdf -ReportPath qa/fields.json -UpdateBibliography
python scripts/docx_ops.py audit-assets source.docx --compare final.docx
python scripts/render_pdf.py qa/final.pdf qa/pages
```

The sequence is: structural conversion → Word field updates and saving → image and embedded-object audit → rendering and visual inspection. Output files must use new paths. Choose different output names when repeating a run; `qa/pages` must also be a new directory.

The bibliography command expects the heading `参考文献` by default. For another heading, supply an exact match with `--heading "References"`, for example.

#### Export a PDF without updating fields or modifying the document

```powershell
pwsh -NoProfile -File scripts/word_finalize.ps1 -InputPath source.docx -PdfPath qa/source.pdf
```

#### Figure captions and cross-references

The saved personal preference is `图 1 题注文字`: a space between the label and number, and another before the caption text. Captions use 14 pt (Chinese “四号”), SimSun for Chinese, and Times New Roman for Latin letters and digits. In this user's proposal, the complete body phrase `如图 1所示` also uses 14 pt, not just its figure number. Other documents follow their own template or user instructions.

```powershell
python scripts/docx_ops.py figures source.docx --output qa/captions.docx --caption-size 14
pwsh -NoProfile -File scripts/word_finalize.ps1 -InputPath qa/captions.docx -OutputPath captions-final.docx -UpdateFigures
```

This conversion supports only simple native `SEQ` figure captions without existing caption bookmarks. Preserve valid existing bookmarks and use targeted edits instead of rebuilding them. Use `--label-separator none` only when no space is wanted between the label and number. Use `-FigureReferenceSize 12` only when the applicable body text is uniformly 12 pt; otherwise match each reference to its surrounding text.

After checking the native cross-reference fields, the complete body phrases can be formatted separately. Replace the example count with the actual number of matches and use a new output path:

```powershell
python scripts/format_figure_ref_phrases.py source.docx --output figure-refs.docx --size-pt 14 --expected-count 23
```

For field-preserving sentence moves and optional formatting, see the [detailed guide](references/word-mechanics.md) or run `python scripts/docx_ops.py --help`.

### Tests

```powershell
python -m unittest discover -s scripts -p "test_*.py" -v
pwsh -NoProfile -File scripts/test_word_fields.ps1
```

The first command runs offline regression tests. The second requires Windows and Microsoft Word. It creates synthetic documents to check list numbering, cross-references to arbitrarily named caption bookmarks, body-reference font sizes, and protection of locked equation-number fields. Fixtures are saved in a new temporary directory; use `-OutputDirectory` to specify another directory that does not yet exist.

Passing these tests does not validate a particular manuscript's scientific content or layout, and does not test editing inside the MathType add-in.

### Limitations and precautions

- **Preserving MathType is not the same as converting equations to MathType.** Native Word OMML equations are not MathType objects. Verify any conversion workflow on a copy first.
- `bibliography` requires an exactly matched heading followed by consecutive manually typed `[1]` through `[N]` entries. Existing automatic numbering, relevant bookmarks, tracked revisions, or citation-manager `ADDIN` fields trigger safety checks. It cannot distinguish all bibliography citations from mathematical intervals, and does not automatically reorder entries by first citation.
- `figures` supports only simple native `SEQ 图` captions without existing caption bookmarks. It does not convert arbitrary manually typed captions or judge whether a figure supports the accompanying text.
- Do not force updates using global F9, bulk unlocking, or converting fields to plain text. The Word COM helper updates only supported selected fields and does not update or unlock fields managed by MathType, Zotero, or EndNote.
- Current inspection and updates primarily target the document body. Headers, footers, text boxes, and complex tables may need dedicated handling. Inline-equation spacing, equation numbering, and local fonts may also need targeted edits.
- Identical image and embedded-object payloads do not prove that placement or layout is correct. Inspect affected pages for clipping, table borders across pages, equation legibility, and reference display.
- Distinguish published methods from project-specific adaptations. For example, calculating STD directly on dBm values is a project choice, not an equivalent reproduction of linear-power STD or a global algorithm default for this skill.
- Keep only the skill and reusable utilities in this repository. Do not commit manuscripts, experimental data, rendered pages, or other sensitive materials from your working directories.
