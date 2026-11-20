---
name: office-word
description: Word 模块规划入口。用于界定桌面 Word 长文档排版、源稿结构与内嵌图件提取和文档一致性预检。
---

# Word

当前状态为 `planned`。[规划与边界](README.md) 记录拟实现能力。处理 Word 任务时先明确源文件、目标样式和交付范围。长文档排版、保存和导出使用 Windows 上的 Word COM。若任务仅要求提取源稿内容，可以只读提取并保留段落、表格、图注与图件的对应关系。若目标是制作 PPT，按 [跨应用协作契约](../../references/cross-app-workflow.md) 交接已提取的资产。
