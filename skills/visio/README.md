# Visio

`validating`。基于上游 [ws-cczj/visio-skill](https://github.com/ws-cczj/visio-skill) 的 `53be9fa`。

用自然语言在 Microsoft Visio 中画图、修改图纸，或根据一张参考图生成可编辑的 VSDX。

启用 [Visio skill](SKILL.md) 后，发一句需求就可以；有参考图或现有 Visio 文件时，直接作为附件发送。例如：

> 用 Visio 复绘这张图，文字、标注和示意结构要能编辑，显微照片保留原样。

也可以一次发多张参考图，例如：“把这些图合成一张 Visio 图：用第一张的布局、第二张的结构，第三张保留为显微照片。”

AI 会选择绘制方式、在后台操作 Visio，检查并交付图纸。运行环境需要 Windows、Python 3.10+、pywin32 和桌面版 Visio。从仓库根目录运行以下命令，完整接口见 [内置连接](references/bridge.md)。

```powershell
python -m pip install -r skills/visio/requirements.txt
python skills/visio/scripts/visio_bridge.py --action status
```

## 图件交付

- 可编辑 `.vsdx` 放入本次运行的 `visio/`。保存后从最终文件重开并核对对象、标签和连接关系。导出图置于 `visio/assets/`，供 Word 或 PPT 使用。
- 按目标栏宽或图位明确导出尺寸。PNG 记录实际像素与 DPI，SVG 或 EMF 记录画布尺寸。按用户要求提供纯白或透明背景，在实际目标尺寸下核对标注字号、内容边界和清晰度。
- 在 `visio/assets/manifest.md` 逐件记录图件 ID、源 VSDX 与页面、导出文件、格式、尺寸、背景、建议放置位置和已验证范围。若涉及跨应用回填，记录目标页或图位标识与用户确认的版本。
