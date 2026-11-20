# 从纯文字设计 Visio 图

## 先做设计判断

用一句话概括读者看完应明白什么。列出必画对象、它们之间已知的关系、需要凸显的一处信息，以及出图场景。把用户给的事实和设计选择分开：可以决定颜色、方向或留白，但不能凭空补技术模块、数据、指标或因果。只有一个缺失信息会改变图意时，才问那个问题；一般细节自己决定。

| 要讲清的关系 | 版面选择 |
| --- | --- |
| 步骤、分支、回路 | 主线同一行或列，分支移到旁路，决策线上标结果。按职责交接时用泳道。 |
| 组成、分层、依赖 | 按职责或层级分区，连接到具体模块。给主路径和辅助依赖不同视觉权重。 |
| 并列方案或两种状态 | 对齐面板和对应项，用一致的比例、标签和顺序，突出真正的差异。 |
| 结构和机理 | 先立主体形状与视角，再叠加局部放大、箭头或图例。依据不足的科学细节不要画成事实。 |

绘制前做一个内部草案：主线怎么读、哪些内容并列、哪些是次要说明。复杂图先用少数分区安排大关系，再填内容；避免把所有内容排成一串等权方框。先定每区可读字数、边距、留白和走线空间。统一字号层级、颜色和箭头含义，再用原生 Visio 对象绘制。

## 用排版底座画流程

单主线、少量分支的图可用 `scripts/visio_flow.py` 把结构草案转成标准绘图请求。**关系由 AI 根据需求确认**，脚本只处理网格、字号、对齐和可编辑连接线。流程图不适合表达某些结构；复杂泳道、跨区架构或科学机理图，按草案用 [Visio 桥接](bridge.md) 定制。

```json
{
  "title": "采样数据验收",
  "direction": "lr",
  "nodes": [
    {"key":"sample","text":"采集样本","stage":0,"role":"start"},
    {"key":"check","text":"数据完整？","stage":1,"role":"decision"},
    {"key":"save","text":"保存数据","stage":2,"role":"end"},
    {"key":"flag","text":"标记缺项","stage":1,"track":1}
  ],
  "edges": [
    {"from":"sample","to":"check"},
    {"from":"check","to":"save","label":"是"},
    {"from":"check","to":"flag","label":"否"}
  ]
}
```

`stage` 控制主方向顺序，`track` 默认为 0；在 `lr` 画法中正值代表主线下方，在 `tb` 中代表右侧。每个位置只能放一个节点。`role` 默认为 `step`，还可用 `start`、`decision`、`end`。把内部蓝图和请求放临时目录，自动生成请求后，用现有 Visio bridge 执行。保存请求时显式指定 UTF-8：

```powershell
python skills/visio/scripts/visio_flow.py --blueprint C:\work\blueprint.json --document C:\work\delivery\figure.vsdx --preview C:\work\delivery\figure.png | Set-Content -Encoding UTF8 C:\work\request.json
python skills/visio/scripts/visio_bridge.py --request C:\work\request.json
```

## 成图后回看

1. 对照用户需求查必画内容、箭头方向和分支条件。检查保存重开后标签和连接仍可编辑。
2. 看实际导出的整页，确认不看操作说明也能找到焦点、起点和阅读顺序；在正常观看尺寸读最长标签、检查交叉、绕线、裁切和重叠。只有布局位置正确还不算通过。
3. 主线不清就改分区或方向，文字拥挤就改节点大小、间距或表述，绕线遮挡就移节点或给旁路留位置。以检查结果决定改动；简单图无需额外画多个方案给用户选。
