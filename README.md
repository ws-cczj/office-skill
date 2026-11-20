# office-skill

面向 Agent 的 Windows 桌面办公技能仓库。按当前交付物选择 PowerPoint、Visio、Word 或 Excel，使用内置 Python 工具通过 COM 操作对应的桌面应用。

| 应用 | 状态 | 用途与入口 |
| --- | --- | --- |
| PowerPoint | `ready` | 学术、工作和教学演示，含可编辑 PPTX、逐页发言稿、备注和终版预览。[使用说明](skills/ppt/README.md) |
| Visio | `validating` | 可编辑工程图、机理图、流程图和技术路线图。[使用说明](skills/visio/README.md) |
| Word | `validating` | SCI 稿件、原生公式与 MathType、编号交叉引用和 PDF 导出。[使用说明](skills/word/README.md) |
| Excel | `validating` | 数据清洗、多表公式、原生图表与公式审计。[使用说明](skills/excel/README.md) |

## 快速开始

让 Agent 读取本仓库的 [SKILL.md](SKILL.md)，附上材料并说明交付物。比如“用这份开题报告做成可编辑 PPT，附逐页发言稿”，或“用 Visio 画一张可编辑的数据采集流程图，附透明背景图”。

运行环境需要 Windows、Python 3.10+、pywin32，以及对应应用的桌面版。以下命令从仓库根目录运行。PowerPoint 缺 pywin32 时可先运行 `python skills/ppt/scripts/powerpoint.py setup`。Visio 依赖见 [Visio 使用说明](skills/visio/README.md)。

```powershell
python skills/ppt/scripts/powerpoint.py inspect
python skills/ppt/scripts/powerpoint.py probe
python skills/visio/scripts/visio_bridge.py --action status
python skills/word/scripts/word.py probe
python skills/excel/scripts/excel.py probe
```

PowerPoint `inspect` 检查依赖。各入口的 `probe` 或 Visio `status` 连接真实桌面应用。工作文件按 [跨应用协作契约](references/cross-app-workflow.md) 放入受 Git 忽略的 `runs/<run-id>/`。

## 维护

从仓库根目录运行各子 Skill 的回归。Word 和 Excel 测试需要相应桌面应用，MathType 用例需要 MathType 7 和 Word Commands 插件。按应用依次运行以共用同一桌面锁。

```powershell
python -m unittest discover -s skills/ppt/tests -v
python -m unittest discover -s skills/word/tests -v
python -m unittest discover -s skills/excel/tests -v
```

项目级训练、校验和成熟度约定见 [AGENTS.md](AGENTS.md)，共享锁和保存重开约定见 [core/README.md](core/README.md)。

仓库地址为 [ws-cczj/office-skill](https://github.com/ws-cczj/office-skill)。项目许可证尚未确定。文档不代表对第三方素材或代码的授权。
