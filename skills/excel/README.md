# Excel

状态为 `validating`。需要 Windows、Python 3.10+、pywin32 和桌面 Excel。控制器操作原生工作簿，支持 XLSX 工作副本、XLTX 新建、保存重开与图表导出。

从仓库根目录执行。

```powershell
python skills/excel/scripts/excel.py inspect
python skills/excel/scripts/excel.py probe
python skills/excel/scripts/excel.py run --plan runs/demo/excel/plan.json
python skills/excel/scripts/excel.py snapshot --target (Resolve-Path runs/demo/excel/summary.xlsx)
python skills/excel/scripts/check_sheet.py runs/demo/excel/summary.xlsx --sheet Summary
```

计划为 UTF-8 JSON。必填 `schema_version: 1` 与新建的绝对 `target`，可选绝对 `source`（XLSX）或 `template`（XLTX）。`snapshot.ranges` 可指定多项 `sheet` 和 `range`，保存重开后回读实际值和公式。`exports` 中每项用绝对 `path`，PNG 另填 `sheet` 和 `chart`，PDF 可选 `sheet`。已有目标会被拒绝，成功重开并导出后发布结果。

| 操作 | 字段及行为 |
| --- | --- |
| `sheet` | 按 `name` 新建或复用工作表。可用 `rename_from` 按现有名称或从 1 开始的索引重命名。 |
| `write` | 向 `sheet` 中从 `start`（默认 A1）起批量写入矩形 `rows`。字符串保留为文本，数字按 JSON 数字传入。 |
| `import_csv` | 从 `source` 绝对路径导入 CSV。支持 `start`、`encoding`、`delimiter`、`header`、按列号设置的 `numeric_columns`。 |
| `clean` | 按 `sheet` 和 `range` 处理常量数据。可选 `trim`、`remove_blank_rows`、按相对列号指定的 `unique_by`。重复项保留首次出现，清掉所选范围内剩余单元格。 |
| `formulas` | 在 `start` 批量写入 `rows` 中的公式。可选 `notation` 为 `a1` 或 `r1c1`。由 Excel 重算，可引用同簿其他工作表。 |
| `format` | 在 `range` 设置数字格式、字体、字号、粗体、斜体、换行、填色、字色和列宽行高，可指定 `autofit`。 |
| `table` | 用 `range` 创建名为 `name` 的原生 Excel 表格，可选 `style`、`header`。 |
| `chart` | 在 `anchor` 创建名为 `name` 的原生图表。用 `range` 提供数据区域，设置 `kind`、`width_pt`、`height_pt`、`title`、`x_title`、`y_title`。支持 `column`、`bar`、`line`、`scatter` 和 `pie`。 |

跨表汇总示例。以下为计划中的 `operations`。

```json
[
  {"op":"sheet","name":"Raw","rename_from":1},
  {"op":"write","sheet":"Raw","rows":[["Group","Amount"],["A",2],["B",4],["A",6]]},
  {"op":"sheet","name":"Summary"},
  {"op":"write","sheet":"Summary","rows":[["Group","Total"],["A",null],["B",null]]},
  {"op":"formulas","sheet":"Summary","start":"B2","rows":[["=SUMIF(Raw!A2:A4,A2,Raw!B2:B4)"],["=SUMIF(Raw!A2:A4,A3,Raw!B2:B4)"]]},
  {"op":"chart","sheet":"Summary","name":"Totals","range":"A1:B3","kind":"column","width_pt":720,"height_pt":360}
]
```

`check_sheet.py` 离线读取公式记录与 Excel 保存的缓存值。返回编码 1 表示工作表存在缓存的单元格错误。检查器固定输出 `quality_verdict: not_assessed`，图例、单位、文字和图表裁切用原生导出图验收。
