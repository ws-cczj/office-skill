# Word

状态为 `validating`。当前控制器通过 pywin32 连接桌面 Word，支持 DOCX 工作副本、以 DOTX 建稿、保存重开、结构回读和 PDF 导出。需要 Windows、Python 3.10+ 和桌面 Word。MathType 模式另需可运行的 MathType 7 桌面版及已加载的 Word Commands 插件。

从仓库根目录执行。

```powershell
python skills/word/scripts/word.py inspect
python skills/word/scripts/word.py probe
python skills/word/scripts/word.py run --plan runs/demo/word/plan.json
python skills/word/scripts/word.py snapshot --target (Resolve-Path runs/demo/word/paper.docx)
python skills/word/scripts/check_doc.py runs/demo/word/paper.docx
```

`inspect` 只检查环境。`probe` 实际连接 Word 并报告 MathType OLE 注册和插件加载状态。真正能否转换以本次公式生成和保存重开为准。

固定的 SCI 排版可以直接用 [可复用 Word 流程](../../references/reusable-workflows.md#word-sci-稿件)。

## 编排 SCI 稿件

优先使用给定期刊的文档模板及其作者指南。`configure.profile: "sci"` 提供 Times New Roman、正文 12 磅双倍行距和基础标题层级。纸张、页边距、分栏和样式覆盖按本次期刊要求设置。使用模板的样式时直接提供样式名，按需启用通用设置。

`run` 接收 UTF-8 JSON。必填 `schema_version: 1` 与新建的绝对 `target` 路径。可选绝对 `source`（已保存 DOCX）或 `template`（DOTX），以及 `operations`、`snapshot: {"text":true}` 和 `exports`。每项 PDF 导出用绝对 `path`，可同时指定 `first_page`、`last_page`。新建目标已存在时拒绝覆盖。后续修订可传 `update:true` 及原工作稿 `target`，省略 `source` 和 `template`。保存重开与导出验证完成后，更新同一路径的工作稿和指定导出。每个批次至多复用一个隐藏、未落盘的 MathType 转换文档。

操作按顺序执行。先确定页面与样式，再放置公式。以下是计划的 `operations` 示例。

```json
[
  {"op":"configure","profile":"sci","page":{"paper":"a4","columns":1,"left_pt":72,"right_pt":72}},
  {"op":"paragraph","style":"title","text":"Transport Measurements"},
  {"op":"paragraph","style":"heading1","text":"Methods"},
  {"op":"paragraph","runs":[{"text":"H"},{"text":"2","script":"sub"},{"text":"O and x"},{"text":"2","script":"super"},{"text":" follow "},{"ref":"source_a"},{"text":". See Eq. "},{"ref":"balance","kind":"equation"},{"text":"."}]},
  {"op":"equation","id":"balance","engine":"mathtype","text":"\\frac{x_1^2}{y}"},
  {"op":"paragraph","style":"heading1","text":"References"},
  {"op":"reference","id":"source_a","text":"Author A. Journal article title. Journal. Year;volume:pages."}
]
```

| 操作 | 字段及行为 |
| --- | --- |
| `configure` | 可选 `page` 使用 `paper`（`a4` 或 `letter`）、`orientation`、`top_pt`、`bottom_pt`、`left_pt`、`right_pt`、`columns`、`column_gap_pt`。可选 `styles` 按样式名配置 `font`、`font_east_asia`、`size`、`bold`、`italic`、`align`、`line_spacing`、`space_before_pt`、`space_after_pt`、`left_indent_pt`、`first_line_pt`、`keep_with_next`。 |
| `paragraph` | 追加 `text` 或 `runs`。文本 run 可选 `script: "sub"`、`"super"` 或 `"normal"`。`{"math":"x_1^2","engine":"word"}` 插入行内公式。`{"ref":"id","kind":"reference"}` 插入 REF 域，可用 `prefix`、`suffix` 和角标控制引文外观。默认为参考文献，可选 `equation`、`figure`、`table`。 |
| `format_paragraph` | `paragraph` 是现有稿件中从 1 开始的段落序号。可指定 `style` 和段内 `runs`，每个 run 用 UTF-16 `start`、`end` 及可选 `expect` 校对文本跨度，再设置字体或上下角标。 |
| `equation` | `id` 供交叉引用，`text` 给公式，`engine` 默认 `word`（UnicodeMath），或 `mathtype`（MathType TeX 子集，省略外层定界符）。行首用居中制表位，编号为右侧制表位内的 SEQ 域。多栏文档使用等宽栏。需在实际页面验收长公式。 |
| `replace_mathtype` | 在源稿副本中按本工具生成的公式 `id`，或按 Word 中从 1 开始的 `inline_shape` 序号，替换 MathType 对象的内容。公式编号、原 REF 和段落制表位置保留。 |
| `reference` | `id` 加格式化后的书目信息 `text`。序号随参考文献顺序计算，可用 `prefix` 和 `suffix` 定义编号两侧。文字、出版字段与期刊引文格式由任务材料提供。 |
| `figure`、`table`、`caption` | `figure` 的 `image` 为绝对图路径，可选 `width_pt`。`table` 接收矩形 `rows` 和可选 `style`、`header`。`caption` 使用 `kind`（`figure` 或 `table`）、`id`、`text` 和可选 `label`。 |
| `header_footer` | 在指定 `section` 写主 `header` 或 `footer`。需要 `text`，可选 `align`。 |
| `toc`、`refresh_toc`、`page_numbers` | 插入或更新目录。页码写入指定 `section`（默认为 1）的主页面页脚。 |

MathType 对象可以保存后在装有 MathType 的 Word 中继续编辑。转换使用 Wiris 的 Toggle TeX 命令，公式语法取决于本机 MathType 版本。转换失败时保留源稿并报告原因。`check_doc.py` 离线读取主文档的原生公式、MathType OLE、角标、制表位、SEQ 结果和 REF 目标，输出 `quality_verdict: not_assessed`。 [Wiris MathType 7 Word 文档](https://docs.wiris.com/using-mathtype-other-applications-and-websites/mathtype-7-with-microsoft-office-2016-or-later)
