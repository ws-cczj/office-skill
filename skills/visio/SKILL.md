---
name: visio-skill
description: Create, edit, or recreate editable Microsoft Visio drawings from a natural-language brief, one or more reference images, or existing VSDX. Use for flowcharts, architecture diagrams, and scientific illustrations on Windows.
---

# Visio

当前状态为 `validating`。接受独立绘图任务和用户明确确认的补图清单。流程图、力学或化学机理图、系统架构图和技术路线图均通过原生 Visio 绘制。跨应用交付按 [交接契约](../../references/cross-app-workflow.md) 执行。

Treat the user's brief and attachments as the drawing request. Infer routine layout, dimensions, and technique. Ask only when missing information would materially change the drawing or make its content unreliable. Handle script calls, intermediate files, and verification yourself.

## Choose the drawing method

- **Text brief / original design:** Start with the message the reader should take away, required facts, and the relationships they imply. Pick a layout that makes those relationships clear, then draft the reading order before placing shapes. Use [original design](references/original-design.md) to choose a layout, plan an editable diagram, and verify it. A simple ranked flow can use the bundled `scripts/visio_flow.py`; other layouts use [the bridge](references/bridge.md).
- **Reference image:** Identify text, symbols, diagram objects, and photographs separately. Prefer native objects for text, layout, connectors, and model parts the user may edit. For parts with three-dimensional volume, reconstruct the shape and projection before styling; use [shape reconstruction](references/geometry.md) to choose and check the geometry. For microscopy and photos, crop without reducing the supplied pixels; avoid drawing an editable annotation over the same mark already baked into the crop. Preserve panel order and correspondence. Use [PNG outlines](references/raster.md) when visual contour matters more than semantic editing. Avoid embedding the whole screenshot as the drawing unless the user asks for a bitmap.
- **Multiple reference images:** Read them together before drawing. Decide which source controls layout, structure, details, or evidence. Match different views only when their correspondence is supported; do not combine incompatible objects or conditions. When the user wants one figure, compose one page unless they specify otherwise. Use [multiple references](references/multiple-references.md) to plan the composition and check source conflicts.
- **Existing VSDX:** Inspect by path and page before editing. Work on a copy unless the user wants in-place changes. For dense native geometry, use `snapshot` to read saved state or `rebuild` to save a separate copy through Visio COM. Refer to the bridge reference for file support and limitations.

For scientific figures, keep the original evidence and meaning. Do not guess illegible labels or invent relationships. Flag only uncertainties that affect the result.

## Draw in the background

Use `skills/visio/scripts/visio_bridge.py` from the repository root. It connects to desktop Visio on Windows with Python 3.10+ and pywin32. Check the connection with `python skills/visio/scripts/visio_bridge.py --action status`. Handle ordinary dependency setup on the execution host using `requirements.txt` or, for PNG tracing, `requirements-raster.txt` in this module. The bridge holds the [shared desktop lock](../../core/README.md) for its COM request. If Windows, Visio, or local execution is unavailable, state the missing prerequisite; do not claim to have rendered or verified a VSDX.

- Run `render` for a new document and `apply` with `save:true` for edits. Pass an explicit page for a multipage document and use absolute paths in requests. Keep user documents and unsaved changes intact.
- Keep window handling at its default; set `show:true` only when the user wants the result opened. Isolate temporary requests, crops, OCR data, and diagnostics from the delivery files. If a drawing batch fails, inspect what was saved before retrying.

## Verify and deliver

Reopen the saved VSDX and check that the intended text, objects, and connections are present. Inspect a full-page preview for legibility, clipping, and alignment. For reference reproductions with complex model parts, compare the key parts against the source at the same scale and check silhouettes, seams, and occlusion before accepting the full page. For dense figures, export PNG at `png_dpi:300` or use SVG for zooming when suitable. Increasing DPI does not restore detail lost in the supplied image.

Deliver the editable VSDX. For figure handoff, export PNG, SVG or EMF with the requested background and readable labels, then record the usable exports in `assets/manifest.md`. Follow the [figure handoff specifications](README.md#图件交付). Mention relevant limitations, such as regions retained as images or labels that were unreadable. Keep diagnostic JSON, logs, benchmarks, and internal checks local; the user does not need to run scripts, validate the result, or send it back before delivery.
