# 可复用办公流程

按应用选用经过验证的流程，填当前材料和变量。控制器继续负责 COM 操作和原生导出。逐页内容、引文和数据口径仍由当前材料决定。

在仓库根目录运行。把参数 JSON 放在本次 `source/`，相对输入路径以参数文件目录为准。

```powershell
python -m core.recipes ppt --run-dir runs/demo --params runs/demo/source/ppt.json
python -m core.recipes visio --run-dir runs/demo --params runs/demo/source/visio.json
python -m core.recipes word --run-dir runs/demo --params runs/demo/source/word.json
python -m core.recipes excel --run-dir runs/demo --params runs/demo/source/excel.json
```

按任务运行其中需要的一项。`--plan-only` 只在标准输出查看展开的原生计划，可按需保存并通过各应用入口进一步调整。正常执行直接使用参数，计划无需另存。工作稿放对应应用分区，预览放 `renders/<应用>/`，PPT 的执行记录和对象快照放 `qa/`。后续修改保持同一运行目录和工作稿路径。

## PowerPoint 模板填充

第一次给出 PPTX 模板路径。使用模板中的稳定页面和对象标识，必要时先用 [PowerPoint 快照](../skills/ppt/references/powerpoint-native.md#运行与探测) 回读。可用 `template_slides` 选源页。

```json
{
  "template": "source-template.pptx",
  "pages": [
    {"slide": 1, "shapes": [{"shape": {"tag": "title"}, "text": "实际页标题"}], "notes": "实际讲述正文"}
  ]
}
```

修改同一份 `ppt/deck.pptx` 时，把 `template` 换成 `"update": true`，填新的 `batch` 名称和受影响页，只重导这些页。每个实际批次在 `qa/` 保留一份执行记录，同一批中断后先核对记录与稿件，再继续余下改动。

## Visio 流程图

提供节点分段、泳道和边，沿用 [Visio 流程图规则](../skills/visio/references/original-design.md)。输出 `visio/diagram.vsdx` 和终版 PNG。

```json
{
  "title": "数据检查流程",
  "nodes": [
    {"key": "Read", "text": "读取输入", "stage": 0},
    {"key": "Check", "text": "核对结果", "stage": 1}
  ],
  "edges": [{"from": "Read", "to": "Check"}]
}
```

修订时传 `"update": true` 和 [Visio apply 操作](../skills/visio/references/bridge.md#批量操作)，以节点 key 定位。已有工作稿继续使用 `diagram.vsdx`，预览沿用原路径。

## Word SCI 稿件

可选 `source` 或 `template` 指定 DOCX 或 DOTX。空白新稿默认应用通用 SCI 样式，给出 `layout` 可明确设置期刊版式。正文及公式沿用 [Word 操作字段](../skills/word/README.md#编排-sci-稿件)。

```json
{
  "layout": {"profile": "sci", "page": {"paper": "a4", "columns": 1}},
  "operations": [
    {"op": "paragraph", "style": "title", "text": "Manuscript title"},
    {"op": "equation", "id": "balance", "engine": "mathtype", "text": "\\frac{x_1^2}{y}"}
  ]
}
```

同一稿修订只给 `"update": true` 和本次操作，例如 `replace_mathtype`。继续写回 `word/manuscript.docx`，并更新对应 PDF。

## Excel CSV 分组汇总

第一行为列名，分组列与求和列用从 1 开始的列号。列出要统计的组，可选 `unique_by` 指定去重字段。流程会清理选定数据、生成真实跨表公式并由 Excel 绘图。

```json
{
  "csv": "measurements.csv",
  "group_column": 1,
  "value_column": 2,
  "groups": ["A", "B"]
}
```

修订时传 `"update": true` 和 [Excel 操作](../skills/excel/README.md)，例如更新 Data 表中的实际值。工作簿和汇总预览沿用原路径。
