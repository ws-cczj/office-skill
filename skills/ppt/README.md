# PowerPoint

面向 Agent 的演示文稿生成技能，拟支持学术汇报、企业汇报、大学课堂作业等场景。

已提供 [SKILL.md](SKILL.md) 技能入口与内置 PowerPoint 控制工具。Agent 按逐页文稿和布局，在真实 PowerPoint 中绘制与修改，以实际导出的图片检查版式。控制工具通过 COM 调用软件，不依赖第三方 PPT-MCP 或 MCP SDK。

把 skill 放在 Agent 可读的工作区，附上报告后可直接说“把这份开题报告做成 PPT”。默认从已验收风格中按用途选型，并交付可编辑 PPTX、逐页发言稿和实际导出预览。想用特定风格时可在同一句中指定模板 ID。

正式生成使用当前用户提交的材料、Skill 内置规范和所选风格包。可直接复制[生成与试排提示词](references/skill-training.md#交给使用者的提示词)。

## 设计方向

- 制作模板：选定一种风格 → 根据参考图去重页型 → 制作可编辑占位页 → 逐页对照和复制试填 → 保存重开验收。
- 正式汇报：按要求和材料规划内容及证据 → 复用风格框架，按页选择整页复用、重组组件或重绘正文 → 保存重开并核对实际导图。详见[模板复用规则](SKILL.md#按内容复用模板)。

从仓库根目录运行 `python skills/ppt/scripts/powerpoint.py inspect` 检查环境，`probe` 实际连接软件。需要 Windows、PowerPoint、Python 3.10+ 与 pywin32，缺依赖时 `setup` 只安装 pywin32。制作与修改使用 `run --plan`，目标、操作、最终导出和保存重开写进本次计划，详细接口见[内置控制工具](references/powerpoint-native.md)。

## 模板资源

内置资源按 `resources/<场景>/<风格>/` 组织，初始场景为 `article`（学术）、`work`（工作汇报）、`education`（课堂教学）。每个可用风格包含 PNG 总览、模板信息和风格规则，并按生成方式提供 PPTX 源模板或已验证样例。只展示真实存在且已验证可用的风格；详细约定见 [模板资源规范](resources/README.md)。

目前已登记两种[学术风格](resources/article/README.md)。藏蓝导航提供已验收的 18 页双语源模板和正式预览。蓝白渐变可用于指定复现测试。两种风格各有 `learning.md`，记录已解决的问题、复查方法和待复验范围。初始设计图归档在本机分析器。

采用 PowerPoint 软件内制作主线，HTML 按用户需要提供网页展示。当前已验证范围和限制见[内置工具验证](docs/native-validation.md)与[开发方案](docs/development-plan.md)。

正式汇报按受众读取[学术](academic/SKILL.md)、[工作汇报](work/SKILL.md)或[教学展示](education/SKILL.md)子 skill，负责选材和逐页发言稿。视觉模板仍从 `resources/` 查找，场景规则可以与不同模板组合使用。

## 开发资料

- [开发方案](docs/development-plan.md)：工作流、目录规划、首版范围及验收。
- [逐页文稿示例](examples/slide-spec.md)：Markdown 与结构化页面信息的约定草案。
- [训练与跨会话交接](references/skill-training.md)：从失败成片和作者参考稿提炼规则，可复制到生成、复审和维护会话。
- [模板制作与验收](references/template-workflow.md)：按参考图确定独立页型、制作占位模板、逐页对照及复用测试。
- [内置 PowerPoint 控制](references/powerpoint-native.md)：环境准备、批量计划、对象定位、脚本扩展与异常恢复。
- [风格复现与页面验收](references/visual-quality.md)：局部组件对照、文本回读、图表编辑模式和修改检查。
- [成片文件检查](scripts/check_deck.py)：只读核对沿用的空图位、指定时间轴截止月、发言稿与备注，并可提示正文重复容器；工具不作内容与美感判定，配合逐页导出图验收。

每批操作绑定目标完整路径，并持有跨进程协作锁。其他客户端或用户不参与此锁，桌面任务仍需串行。失败后先读真实状态，不自动重放可能已生效的修改。

本模块属于 [office-skill](../../README.md)。

未确定项目许可证；当前文档不代表对第三方素材或代码的授权。
