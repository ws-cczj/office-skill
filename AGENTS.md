# office-skill 仓库协作规范

项目仓库为 [ws-cczj/office-skill](https://github.com/ws-cczj/office-skill)。[SKILL.md](SKILL.md) 负责应用路由。应用专属规则、控制器与测试留在对应子目录。共享基础设施见 [core/README.md](core/README.md)。

## 六条仓库铁律

1. **训练边界** — 训练中的“优化”默认修订 Skill 规则、风格 `learning.md` 和确定性校验脚本。历史成片与运行产物作为证据保留。制作或返修以明确任务为准，`runs/` 受 Git 忽略。
2. **原生桌面 COM** — Office 绘制、排版、保存和导出统一通过仓库内控制工具，以 pywin32 调用 Windows 上真实的 Office 或 Visio，使用 Python 3.10+。禁止引入或回退到第三方 PPT、Word、Excel 或 Visio MCP。
3. **规则通用自足** — 使用当次材料、应用规则和选中资源执行。Skill、模板规则和校验脚本不得依赖案例人名、学校、课题名、成片固定页码或机器路径。源模板的对象名与源页索引只用于定位可复用结构。
4. **克制排版** — 以阅读效果、证据关系、文档规范和用户约束决定构图与篇幅。禁止设置机械的全局字数、卡片数或图文比例上限。
5. **只读确定性校验** — 校验以同一已保存产物与同一组参数为输入，结果可复核且不改动源稿、模板或交付物。统一使用 `quality_verdict: not_assessed`，视觉效果与内容判断另据实际导出验收。
6. **单软件独立完成** — 禁止隐式串联其他桌面软件现场制图。PPT 缺图先交付轻量示意或可替换图位与清单，经用户确认才独立运行 Visio 并定点回填。

## 开发与跨会话交接

- 运行和验收使用工作副本，按 [跨应用协作契约](references/cross-app-workflow.md) 记录输入角色、规则或模板版本、产物路径、受影响位置和待复验项。修改规则只使用可复用结论。私有案例和证据留在被忽略的本机目录。
- 控制器以完整路径绑定本批目标，桌面操作使用共享锁串行执行。异常后先核对已执行操作与目标状态，再决定如何恢复。
- 从仓库根目录执行各子 Skill 已有测试。移动入口或资源时核对引用、命令示例、模板元数据和适用回归。改动真实 COM 行为时，在相应应用中验证保存重开、导出和文档生命周期。缺少环境或测试时按实际情况记录。

## 子 Skill 成熟度

| 子 Skill | 状态 | 维护入口 |
| --- | --- | --- |
| PowerPoint | `ready` | [PPT 开发与训练](skills/ppt/AGENTS.md) |
| Visio | `validating` | [Visio 开发与训练](skills/visio/AGENTS.md) |
| Word | `validating` | [Word COM 与稿件规范](skills/word/README.md) |
| Excel | `validating` | [Excel COM 与工作簿规范](skills/excel/README.md) |

`ready` 已具备正式工作流。 `validating` 已有实现并继续验证。 `draft` 为可评审规则。 `planned` 为职责与实现规划。
