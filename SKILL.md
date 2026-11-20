---
name: office-skill
description: 按交付物路由 Windows 桌面 Office 任务。用于 PowerPoint 演示、Visio 图件、Word 长文档和 Excel 数据表的制作、修改或校验。
---

# Office Skill

先读 [仓库协作规范](AGENTS.md) 中的单机运行与文件管理，再按当前交付物读取对应入口。用于制作的输入格式不改变路由，例如用 Word 报告制作 PPT 时只加载 PowerPoint。

| 交付物 | 子 Skill |
| --- | --- |
| 演示文稿 | [PowerPoint](skills/ppt/SKILL.md) |
| 可编辑工程图或流程图 | [Visio](skills/visio/SKILL.md) |
| 长文档或源稿结构与素材提取 | [Word](skills/word/SKILL.md) |
| 工作簿或数据图表 | [Excel](skills/excel/SKILL.md) |

单软件任务独立完成。禁止隐式启动另一款桌面应用现场制图。出现需要另绘的缺口时，先完成已授权的交付，再由用户决定是否单独补绘。

产物交接、缺图确认和定点回填使用 [跨应用协作契约](references/cross-app-workflow.md) 的 `runs/<run-id>/` 分区。
