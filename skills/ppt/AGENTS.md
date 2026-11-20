# PowerPoint 开发与训练

本文件适用于 PowerPoint 模块，继承 [仓库协作规范](../../AGENTS.md)。使用或维护 skill 时先读 [SKILL.md](SKILL.md)。下方的本机检查点用于恢复相关任务，按当前用户请求决定是否继续。

## 项目规范

- 模板制作与正式汇报按 [SKILL.md](SKILL.md) 分流。单风格模板按[模板制作与验收](references/template-workflow.md)去重页型并逐页对照。正式汇报由本次要求和材料确定正文，再按[模板复用规则](SKILL.md#按内容复用模板)整页复用、重组组件或重绘正文。工具探测、协作锁和完整目标路径核对按[内置 PowerPoint 控制流程](references/powerpoint-native.md)执行。桌面操作串行进行，调用异常后先回读目标状态，再决定是否继续。正式流程不安装、连接或回退到第三方 PPT-MCP。
- Skill 训练管理者负责验收当轮样稿、定位失误、修改规则及准备下一轮指令。训练语境中的“验收后优化”默认指优化 Skill。样稿制作与返修由收到明确制作指令的会话执行。跨会话使用[训练与交接指令](references/skill-training.md)，本机证据按 `template-analyzer/learning-workflow.md` 处理。
- 报告附件加一句明确的生成要求即可启动制作。按 [一句话生成](SKILL.md#一句话生成)自动确定模板和可推断项，走完 PPTX、发言稿、备注及实际预览的检查流程。
- 正式生成与面向使用者的提示词只依赖本次任务材料、Skill 规范和所选风格包。开发者的私有资料、以往成品、手工质量参考及机器路径留在本地审查记录。风格包的 `learning.md` 写入自足的生成动作和验收条件。
- 同一页的用途允许多种版式。先按内容、证据关系和阅读顺序判断，再复用或调整模板。正式成片验收版式是否服务当前内容，并核对所选风格的识别特征。
- 正式成稿将实际正文和图示关系交给绘制计划，难页先试排并[核对导图](references/visual-quality.md#试排与导图核对)再扩展。保存后的文件按 [文件检查](references/visual-quality.md#从文件核对三类问题)核对图位、时间轴与讲稿。实际逐页导图同时检查论证和视觉效果，记录具体证据。
- 修改已有成品时另存副本，保持稳定对象 ID，并同步受影响的文稿、备注和来源记录。每次运行的材料与结果放在仓库根目录的 `runs/<run-id>/`，按 [跨应用协作契约](../../references/cross-app-workflow.md) 分区。
- 资源按 `resources/<场景>/<风格>/` 组织。`article` 为学术汇报，`work` 为工作汇报，`education` 为教学展示。读取 `template.yaml`、`style.md`、`learning.md` 和 `validation.md`。用户指定复现测试时可使用 `draft` 或 `validating`，正式选型只推荐 `ready`。正式 `preview.png` 来自同版本、实际验收的样例。详见 [资源规范](resources/README.md)。
- 模板只保留版式、样式和可编辑占位符。具体论文、数据、作者和机构身份由任务材料提供。截图中的水印、界面浮层、缩略图栏和具体正文只属于参考素材。截图分析所得比例与字体属于估计或实施建议。
- 仓库根目录的 `template-analyzer/` 是本机工作区，用于保存初始设计图、私有案例和原始证据，目录内遵循其 `AGENTS.md`。每轮将可复用的结论、具体修正与复验范围更新到对应风格的 `learning.md`。生成与验收按该风格包读取学习结果。分析器和运行产物 `runs/` 均受 Git 忽略。
- 风格验收同时检查页头、正文布局、图文层级和组件细节。重点看轮廓与连接、比例、渐变连续性、线型与阴影。检查渐变上的整段文字对比度、图片宽高比、文字填充后的实际边界，以及图表单位、来源和统计语义。区分插图、形状重绘与原生数据图表。依据 [视觉验收规则](references/visual-quality.md) 结合实际导出和对象回读判断。

## 已登记的学术风格

入口为 [学术风格目录](resources/article/README.md)。以各自元数据的版本和状态为准。

| 风格 ID | 规则 | 复现重点 |
| --- | --- | --- |
| `article/literature-blue-gradient` | [文献汇报规则](resources/article/literature-blue-gradient/style.md) | 连续渐变、文献与元数据主次、左窄问题区与右宽证据区、细线卡片 |
| `article/proposal-navy-navigation` | [开题论证规则](resources/article/proposal-navy-navigation/style.md) | 与导航顶部相连的异形高亮、紧凑标识、细实线与虚线卡片、左侧栏目标签和卡内层级 |

藏蓝导航为已验收的 `source` 模板，蓝白渐变为 `validating`，版本和学习文件以各自元数据为准。

## 本机任务检查点（2026-11-12）

藏蓝导航双语模板已独立验收并入库。使用或变更该风格时读取 [维护说明](resources/article/proposal-navy-navigation/maintenance.md) 和 [验证记录](resources/article/proposal-navy-navigation/validation.md)。

## 内置控制工具（2026-11-16）

`scripts/powerpoint.py` 直接控制真实 PowerPoint，依赖 pywin32，无 MCP 服务和 SDK。原型与隔离测试见[内置流程](references/powerpoint-native.md)，实测范围见[验证记录](docs/native-validation.md)。内容与设计仍须逐页验收。旧后端与运行环境仅作为本机对照证据保存在被忽略的 `runs/`，不进入分享包。

## 正文内容学习（2026-11-17）

审查开题成片发现：有实图和模板组件，背景仍可能缺少研究动机，技术正文也可能退化为名词。先从报告保留必要论据并写实际上屏文案，以不看讲稿也能跟上论证为准。见风格包的[内容与构图选择](resources/article/proposal-navy-navigation/content-design.md)及其学习记录。
