# Word编号、交叉引用与脚本用法

## 参考文献

可见的`[n]`应由Word原生编号列表生成。每个参考文献段落通过`w:numPr`关联编号定义，其中编号格式为`w:lvlText="[%1]"`。不要再重复添加手打的`[n]`、SEQ编号或项目符号。

为目标参考文献段落建立书签，正文引用使用以下域代码：

```text
REF AWBib1 \n \h
```

`\n`返回原生段落编号及其编号格式，例如`[1]`；`\h`使该引用可以点击跳转。这里引用的是“编号项的段落编号”，不是整条参考文献。Word中“编号项”的交叉引用选项名称，与题注中的“仅标签和编号”有所不同。

对于已有的区间引用`[1-2]`，在两个域外保留普通文本形式的方括号和连接号：

```text
[ { REF AWBib1 \n \h \# "0" } - { REF AWBib2 \n \h \# "0" } ]
```

上面的空格和花括号仅用于说明域的边界，实际显示时不要加入这些空格。数字格式开关`\# "0"`使两个端点只显示数字，不各自带方括号。在这种编号格式下，仅使用`\t`不能达到这一效果。单条引用和带数字格式开关的引用均已通过原生Word COM测试。

通过书签绑定具体文献，不要只保留写死的显示数字。新增或重排文献后，更新域，并检查每一个引用区间：如果原先相邻的文献不再连续，应展开或重建引用，不要使区间错误地包含中间其他文献。采用顺序编码制时，核对首次引用的先后顺序，并将区间展开检查。不能只重排参考文献列表的编号，却不处理正文引用目标。

`bibliography`目前支持的输入是：一个完全匹配的标题，后面紧接连续、手动输入`[1]`至`[N]`的参考文献段落。遇到已有自动列表编号、AWBib书签、由插件管理的ADDIN域、修订记录或超出范围的引用编号时，脚本会拒绝处理。运行前应检查参考文献列表之外所有“方括号内数字”的匹配结果：脚本不能自行判断它是文献引用，还是数学区间、数据索引。遇到不支持的结构，应编写范围明确的适配处理，而不是删除安全检查。

## 题注与句子移动

- 当前保存的图题注偏好为：`图 `＋`SEQ 图 \* ARABIC`＋一个空格＋题注文字。
- 书签范围只包含`图 `和SEQ编号，不包含后面的空格、题注文字。
- 正文使用`REF <bookmark> \h`，显示效果例如`如图 3所示。`。其中`<bookmark>`应替换为实际书签名。本用户的开题报告要求完整短语`如图 3所示`为14 pt；要同时格式化“如”、REF显示结果和“所示”，不能只修改引用结果，也不能改写域代码。
- 现有`_Ref...`书签用于维持引用关系，移动或润色题注时应保留。修改题注文字不代表可以重建整个段落。
- `figures`只处理**尚未建立题注书签**的简单原生`SEQ 图`题注，支持复杂域和`w:fldSimple`两种存储形式。遇到额外域、特殊SEQ开关、锁定域、复杂对象或超链接时拒绝处理。默认标签与编号间有空格；仅在用户要求无空格时使用`--label-separator none`。它不负责转换手打题注、处理表格题注、推测缺失题注，也不能判断图件与正文论述在含义上是否对应。
- 不按书签前缀猜测题注类型。Word生成的`_Ref...`及用户命名的书签均可能指向题注，应核对范围中实际的SEQ域和标签编号。处理段落内域时必须限制查询范围，并拼接分散在多个`instrText`中的同一域代码。
- `lead-sentence`在两个完全匹配的章节标题之间，将每个匹配段落的最后一句移到段首，并保留原有REF节点、文字片段格式及正文文字。如果短语重复、移动范围内有书签、句子已经位于段首，或匹配数量与预期不符，脚本会拒绝处理。使用前必须确认所选章节和短语。

## MathType与域

检查`word/embeddings`、`word/media`、OLE的`ProgID`、域代码、关联关系及缓存的显示结果。嵌入文件逐字节未变，只能证明对象得到保留，不能证明已成功编辑MathType公式。

- Word保存时可能重命名媒体文件或改变OLE的ObjectID。审计不能只靠文件名或ObjectID，应结合关联关系与实际载荷，核对内容和数量。现有`audit-assets`仅比较载荷及数量，不证明对象仍出现在原位置，也不检查完整关系图。
- 格式修改应保留公式文字片段的`w:position`、基线、OLE尺寸及预览关联，不能为了统一字体清空整个`rPr`。公式上下被裁切时，检查固定行距和表格固定行高；只修复有关段落或行，不全篇重设。

MathType公式编号可能包含嵌套且锁定的`SEQ MTEqn`域；文献管理软件使用ADDIN域。对整篇文档按F9、批量解锁域、将全部域转成文本，或覆盖整个段落，都可能破坏这些结构。`word_finalize.ps1`只更新明确选中的图题SEQ/REF域或AWBib REF域，不更新也不解锁MathType/ADDIN域。

通过COM导出时，不应操作用户正在使用的Word实例。应单独启动Word，禁止宏自动运行，以只读方式处理输入文件，再通过SaveAs另存新路径。正常清理时不得使用taskkill强制结束进程。

传给`Documents.Open`、`SaveAs2`和`ExportAsFixedFormat`的路径显式转为`[string]`，避免PowerShell路径包装对象引起COM参数绑定异常。脚本采用UTF-8，建议使用PowerShell 7（`pwsh`）；若在Windows PowerShell 5.1运行，先确保脚本为UTF-8 BOM，避免中文域名被误读。只读导出无需更新任何域。

## 命令示例

以下路径仅为示例，应根据当前环境选择实际Python解释器和文件。先创建输出目录和质量检查目录（示例中的`qa`）。DOCX和PDF的输出路径必须是尚不存在的新文件。

```powershell
python scripts/docx_ops.py inspect source.docx --report qa/source.json
python scripts/docx_ops.py bibliography source.docx --output qa/numbered.docx --report qa/conversion.json
pwsh -NoProfile -File scripts/word_finalize.ps1 -InputPath qa/numbered.docx -OutputPath final.docx -PdfPath qa/final.pdf -ReportPath qa/fields.json -UpdateBibliography
python scripts/docx_ops.py audit-assets source.docx --compare final.docx
python scripts/render_pdf.py qa/final.pdf qa/pages
```

上述命令依次用于：检查原稿、转换参考文献编号和引用、通过Word更新并另存、核对图片及嵌入对象、生成排版检查图片。

检查章节后，在保留域的前提下移动引导句：

```powershell
python scripts/docx_ops.py lead-sentence source.docx --output moved.docx --start-heading "研究方法与技术路线：" --end-heading "理论分析与计算：" --phrase "具体研究流程如" --expected-count 4
```

这里的数字4只是示例，**不代表研究内容必须有四点**，应改为本次任务实际要求的数量。

处理符合条件的简单题注：

```powershell
python scripts/docx_ops.py figures source.docx --output qa/captions.docx --caption-size 14
pwsh -NoProfile -File scripts/word_finalize.ps1 -InputPath qa/captions.docx -OutputPath final.docx -PdfPath qa/final.pdf -UpdateFigures
```

`--caption-size 14`指定四号题注，省略时保留原有字号。已有题注书签时不要重复运行`figures`；用范围明确的编辑保留其身份。`-UpdateFigures`先更新正文中的原生图/表SEQ，再更新书签内确有这些SEQ且只含标签编号的REF。可选`-FigureReferenceSize 12`统一这些引用结果为12 pt；只在正文的适用字号确为12 pt时使用，字号混合的正文应逐处匹配。无须更新的MathType和ADDIN域保持原样。

当用户明确要求完整的`如图 x 所示`短语统一为四号时，可在核对原生REF域后单独执行：

```powershell
python scripts/format_figure_ref_phrases.py source.docx --output formatted.docx --size-pt 14 --expected-count 23
```

示例中的23必须替换为当前文档实际匹配数。脚本只调整正文普通段落中的完整短语，保留域结构、短语外文字及其他DOCX部件；输出必须是新文件。它不处理页眉、页脚或文本框，仍需通过Word导出并检查受影响页面。

如果只需只读渲染、不更新域，则不传入OutputPath和更新开关。文档渲染工具依赖LibreOffice而当前环境未安装时，可以在已安装Word的环境中使用以下替代方式：

```powershell
pwsh -NoProfile -File scripts/word_finalize.ps1 -InputPath source.docx -PdfPath qa/source.pdf
```

安装修改过的脚本前，执行以下测试：

```powershell
python -m unittest discover -s scripts -p "test_*.py" -v
pwsh -NoProfile -File scripts/test_word_fields.ps1
```

离线测试覆盖：文字分散在多个格式片段中的情况、区间引用、上标保留、域和OLE对象的保护边界、Word生成的分页标记、不覆盖已有文件、按需规范格式、简单题注，以及保留域的句子移动。

实际Word测试会更改临时原生编号列表的起始编号，并通过`word_finalize.ps1`检查任意命名的题注书签、正文引用字号和锁定公式编号域。测试使用新建合成文档，保存在独立临时目录，也可用`-OutputDirectory`指定尚不存在的目录。两类测试都不能代替针对具体文档的渲染、MathType插件内部编辑和视觉检查。COM脚本可核对原生列表编号及REF域的缓存显示结果，最终文档的科学内容仍需单独核验。

## 必须另外检查的限制

- OOXML检查脚本主要面向正文。页眉、页脚和文本框可能需要单独检查。COM更新脚本使用正文的Fields集合，不会遍历全部StoryRange（如页眉、页脚等独立文字区域）。
- 不得把文献管理软件的域直接转成普通文本。只有用户明确要求时，才通过其所属插件更新或转换。
- 手打参考文献的解析要求列表连续。如果列表被非参考文献段落打断，需要范围明确的专用解析处理。
- 格式规范化是可选操作，不是每次交付都必须执行的步骤。脚本有意避开域结果文字和OLE内容；段内MathType周围的空格、符号字体，仍可能需要定位到具体文字片段后处理。
- 脚本能够保留DOCX内部各文件，但无法自动掌握所有学校的格式要求。仍应在渲染结果中检查表格行高、内容遮挡、跨页边框和题注。
