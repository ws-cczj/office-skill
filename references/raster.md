# 从 PNG 生成可编辑 Visio

只有位图参考时，`raster_to_visio.py` 将色块和线条的轮廓写成 Visio 几何，并用桌面 Visio 后台打开、核对形状数量及导出对照图。源 PNG 不会作为背景图片嵌入 VSDX。

```powershell
python -m pip install -r requirements-raster.txt
python scripts/raster_to_visio.py --image C:\work\source.png --output C:\work\traced.vsdx
```

普通模式默认使用 `--colors 32 --min-area 2 --epsilon 0.22`。细线和浅色层较多时，可用 `--colors 64 --min-area 1.5 --epsilon 0.12` 增加轮廓细节，但会生成更多对象。脚本拒绝覆盖已有 VSDX；旁边只保存 Visio 导出的 PNG。用于描摹的裁片应保留在工作目录；混合复绘优先在最终文档中单独嵌入照片裁片，并把标注、框和引线做成独立原生对象。

需要真文字时，可先让 Windows 本地 OCR 返回文字和坐标，检查后写一个 UTF-8 JSON 字符串数组，例如 `["Sealed container","Carbonization"]`，再给描摹命令加 `--reviewed-text C:\work\reviewed.json`。脚本也可用 `--ocr C:\work\ocr.json` 读取之前的结果。未核对的文字留作轮廓。对于子脚标、非白底、重叠文字和化学键，需要人工检查；如果希望箭头随节点移动或某个部件有语义，应使用原生连接线和组合重新构建这些部件。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/ocr_windows.ps1 -ImagePath C:\work\source.png -Language en-US
```

预览默认以截图的像素尺寸为核对基准。需要放大检查原生 Visio 线条时，另导出 `png_dpi:300` 或 SVG；提高导出分辨率不会补回源 PNG 已经丢失的细节。
