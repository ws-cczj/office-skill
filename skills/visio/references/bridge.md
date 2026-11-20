# 内置连接使用约定

从仓库根目录运行以下命令。示例请求路径换成本次任务的实际路径。输出为 JSON，成功返回码 0，失败返回码 1；`ok` 表示整个请求是否完成。保存成功后导出失败仍可能返回失败，先检查文件与文档状态再重试。

```powershell
python -m pip install -r skills/visio/requirements.txt
python skills/visio/scripts/visio_bridge.py --action status
python skills/visio/scripts/visio_bridge.py --action stencils
python skills/visio/scripts/visio_bridge.py --request C:\work\request.json
```

`status` 会连接正在运行的 Visio，没有实例时启动一个可见实例。它报告打开的绘图、保存状态和可见绘图窗口数量。需要本机交互桌面，不能以此保证 Windows 服务、无人登录或跨平台场景可用。

## 请求

`document`、`save_as` 和 `exports` 必须是绝对路径。请求文件使用 UTF-8，支持 BOM。页面名称使用 Visio 的 `NameU`；`status` 会列出这些名称。

| action | 必要字段与效果 |
| --- | --- |
| status | 列出 Visio 版本和打开的绘图、页面 |
| stencils | 扫描 Office 内置模具与 Visio 实际的 MyShapesPath，支持本机中文“我的形状”路径 |
| masters | `stencil`：模具名或绝对路径；可选 `query`，返回本地名和通用名 `name_u` |
| render | `document`：新文件路径；可选 `page`、`width`、`height`、`template`、`operations`；自动保存 |
| inspect | `document`、可选 `page`；读回形状、位置、文字、字体 ID、连接端点和页面越界警告 |
| apply | `document`、`operations`、可选 `page`；只有 `save:true` 或 `save_as` 时保存 |
| layout | `document`、可选 `page`、`style`、`spacing`、`resize_page`；应用布局并保存 |
| export | `document`、`exports`、可选 `page` |
| save | `document`、可选 `page`、`save_as` |
| open | `document`、可选 `page`；打开或定位文件并读回 |
| open_stencil | `stencil`：模具名或绝对路径；打开模具 |
| pages | `document`；`mode` 为 `list`、`add` 或 `activate`，后两者需要 `name`；新增页面自动保存 |
| page_size | `document`、可选 `page`；设置 `width`/`height` 或 `fit_to_contents:true`；可选 `orientation` 为 `portrait`/`landscape`，默认保存 |
| icons | 列出 `VISIO_SKILL_ICONS_DIR` 或 `~/.visio-skill/icons` 中的标签和绝对路径 |
| snapshot | `document`；无 COM 地读取已保存 VSDX，默认递归全部对象；`recursive:false` 仅读顶层 |
| rebuild | `source`、不同的输出 `document`；通过 Visio 打开已保存源文件的副本并另存；可选 `scene_output` 另存独立几何 JSON |

绘图和导出请求可附加 `exports`（PNG、SVG 或 EMF 路径数组）。多页文件的绘图、编辑和导出必须明确 `page`。`render` 默认页面为 12 × 8 英寸；宽高必须为正数。`render` 和 `save_as` 遇到已有文件会拒绝，除非 `overwrite:true`；已打开的目标不允许重新 `render`，请使用 `apply`。普通原位保存沿用已有路径。

桥接默认在后台打开或创建绘图，请求结束时关闭本次打开且已保存的绘图，保留请求开始前已打开的绘图及未保存的修改。需要跨请求持有后台绘图时设置 `close_after:false`；要在界面显示结果时用 `show:true`，重复调用会复用已有窗口。`pages` 的 `activate` 模式也会显示绘图。后台编辑并落盘用 `apply` 加 `save:true`。

有 PNG 导出时，`png_dpi` 可设 36 至 600 的有限数值（如 300）。只在导出期间生效，随后恢复 Visio 全局栅格设置。PNG 仍按 Visio 的内容范围导出；要保留空白页边，可加一块覆盖页面的底色形状。SVG 和 EMF 可用于矢量交付。

## 批量操作

按数组顺序执行。新形状 `key` 唯一，使用 ASCII 字母开头及字母、数字、下划线、连字符，写入 `Shape.NameU`；后续请求可以继续使用。`target`、`from`、`to` 也可以使用当前页面内的数字形状 ID。

```json
{
  "action": "render",
  "document": "C:/work/flow.vsdx",
  "page": "Flow",
  "width": 10,
  "height": 6,
  "operations": [
    {"op":"add","key":"input","kind":"rectangle","text":"输入","x":2,"y":3,"width":2,"height":0.8,"font":"Microsoft YaHei","fill":"#E9F5F3"},
    {"op":"add","key":"output","text":"输出","x":6,"y":3,"width":2,"height":0.8},
    {"op":"connect","key":"flow","from":"input","to":"output","route":"straight"}
  ],
  "exports": ["C:/work/flow.png"]
}
```

- `add`：`key`、`x`、`y` 必填；`kind` 支持 `rectangle`（默认）、`ellipse`、`diamond`、`text`、`master`。宽高默认 2 × 0.8 英寸，坐标为中心点且 Y 向上。`master` 另外需要 `stencil` 和通用 `master` 名称，可用 `masters` 查询。
- `add` 的 `kind:"polyline"` 接受页面坐标 `points`（交替的 X、Y 英寸值），围合路径时首尾相同。
- `add` 的 `kind:"cylinder"` 使用 `rear`、`front` 两个端面中心 `[x,y]`（英寸，Y 向上），以及 `radius`（前端面长半轴）和 `end_ratio`（短半轴与长半轴的比例，0 到 1，不含 0）。可选 `rear_scale`（远端与近端的尺寸比，默认 1）。生成一个可移动组合，侧壁与前端盖分别是 `<key>_Side`、`<key>_Cap`。用 `fill` 和 `cap_fill` 分别设填色。例子和适用范围见[形体复绘](geometry.md)。
- `connect`：`key`、`from`、`to` 必填；使用动态连接线，端点真正黏附到形状。`route` 为 `right_angle`（默认）、`straight` 或 `curved`，默认带结束箭头。
- `update`：`target` 必填，可更改 `text`、`x`、`y`、`width`、`height`、样式和连接线 `route`。未指定字段保持当前值。
- `delete`：`target` 必填。Visio 可能同时影响相关连线，删除后检查状态。
- `container`：`key`、非空 `members` 数组，另可指定 `text`、`padding`（英寸）、容器模具 `master` 和样式。创建真正的 Visio 容器。
- `members`：`target` 容器、`members` 数组、`mode:add/remove`。移动容器会同步平移其成员，嵌套成员去重处理。
- `image`：`key`、`source` 本地绝对路径或 HTTP(S) URL、`x`、`y`，可选 `width`、`height`。
- `badge`：`key`、`target` 容器、图片 `source` 或 `stencil`+`master`；可选 `size` 和 `corner:top_left/top_right/bottom_left/bottom_right`。徽标会加入容器成员。
- `group`：`key`、`members`，创建可编辑组合。组合内对象可按数字 ID 更新，语义键找不到时递归查找。

`layout.style` 支持 `default`、`flowchart_tb`、`flowchart_lr`、`tree_tb`、`tree_lr`、`radial`、`circular`；`spacing` 单位英寸，`resize_page` 默认 true。径向/圆形预设按顶层二维对象布局，复杂容器图先检查副本上的结果。

`inspect` 默认完整读回当前页顶层形状；`recursive:true` 包含组合内对象，子形状坐标标为父组合坐标。`detail:summary` 只返回尺寸和顶层数量，不检查每个对象。`snapshot` 始终只读磁盘保存状态，输出 `quality_verdict: not_assessed`。未保存修改和 master 继承后的有效样式超出快照范围。

`rebuild` 使用 Visio 的打开副本模式读取已保存的源 VSDX，并用 COM 另存到不同路径，保留原件及已打开文件。可选 `scene_output` 只为无 master、media 或 ForeignData 依赖的独立几何记录 JSON。另存副本后按应用交付规则重开并检查。

绘图阶段默认暂缓屏幕刷新和公式重算，并在 finally 中恢复原值；样式使用 `SetFormulas` 批量写入并覆盖相关单元格保护，模具在批次内缓存。可用 `optimize:false` 关闭屏幕/重算暂缓，但仍使用批量样式。返回的 `timings` 含重建、绘图、保存、导出和检查耗时；`target_total` 为累计值，不能和子阶段简单相加。

样式字段：`fill`、`line`、`text_color` 为 `#RRGGBB`；`font` 同时设置西文和东亚字体；`font_size`、`line_weight` 单位为磅；`rounding`、`text_margin` 单位为英寸；`bold` 布尔值；`align` 为 0 左、1 中、2 右；`no_fill`/`no_line` 隐藏填充或轮廓。`line_pattern`、`begin_arrow`、`end_arrow` 使用 Visio ShapeSheet 数值（例如线型 1 实线、2 虚线；结束箭头 0 无、4 三角）。

请求的绘图部分在一个撤销作用域内执行；中途失败则请求撤销整批绘图操作，错误消息报告撤销调用结果。保存、导出、打开文件及模具加载不属于可回滚范围。`inspect` 的越界警告只检查未旋转二维形状的中心/宽高，不是完整的碰撞、文字溢出或连线裁切检测，仍需人工查看 PNG。

实现依据：[Application](https://learn.microsoft.com/en-us/office/vba/api/visio.application)、[动态黏附](https://learn.microsoft.com/en-us/office/vba/api/visio.cell.glueto)、[撤销作用域](https://learn.microsoft.com/en-us/office/vba/api/visio.application.beginundoscope)、[连接线样式](https://learn.microsoft.com/en-us/office/client-developer/visio/routestyle-cell-page-layout-section)、[东亚字体字段](https://github.com/MicrosoftDocs/office-developer-client-docs/blob/main/docs/visio/asianfont-cell-character-section.md)。
