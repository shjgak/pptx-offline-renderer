---
name: pptx-offline-renderer
description: "Renders a .pptx / .ppt presentation to per-slide PNG images entirely offline with Python and Pillow, so a deck can be visually inspected, QA'd, or screenshotted when no PowerPoint, Keynote, or LibreOffice is installed. Use it to preview slides, catch text overflow, out-of-bounds shapes, or missing pictures, and verify colors without any office software. Triggers: render pptx to image, preview the slides, check if text overflows, convert pptx to png, let me see the slides."
---

# PPTX Offline Renderer

## Overview

Render a `.pptx` (PowerPoint) file into per-slide PNG images entirely offline,
using only Python + Pillow and the operating system's built-in CJK font. No
PowerPoint, LibreOffice, or network access is required. The output is an
*approximate* reconstruction meant for visual QA (layout, colors, shapes,
embedded pictures, text overflow), not a pixel-perfect reproduction.

## When to use

- The environment has no office software but you need to "see" a slide deck.
- You are building or editing a `.pptx` (e.g. with `python-pptx`) and must verify
  it does not overflow, misalign, or drop pictures before handing it to a user.
- A user asks to preview / screenshot a presentation.

Do NOT use this for:
- Exact typography / print proofing (use the real office app).
- Editing content (use `python-pptx` or the office SDK instead).

## Quick start

```bash
pip install pillow
python scripts/render_pptx.py path/to/deck.pptx [output_dir]
# 轻量视觉模型（Flash 类）看图预算小，出小图：
python scripts/render_pptx.py deck.pptx out/ --width 1280
# 强视觉模型（Pro 类）可出大图看细节：
python scripts/render_pptx.py deck.pptx out/ --width 2400
```

Outputs `slide_1.png`, `slide_2.png`, … and a `contact_sheet.png` (all slides
side by side) into `output_dir` (defaults to the deck's folder). The script also
prints any problems it detects:

- `OVERFLOW` – rendered text is taller than its text box (text will be clipped in
  real PowerPoint too).
- `BOUNDS` – a shape extends past the slide edge.
- `PIC` – an embedded picture failed to composite (missing/filtered media).

Read those warnings and fix the source `.pptx` (shorten text, enlarge boxes,
re-embed images) rather than treating them as cosmetic.

## Vision-model compatibility (国产大模型适配)

The renderer is **model-agnostic**: it only writes PNG files. Any
vision-capable model can consume them by attaching `slide_N.png` (one slide per
image) or `contact_sheet.png` (all slides in one image) as an image input. No
model-specific code is needed — pick the resolution that fits the model's image
budget:

| 目标模型 / Target model | 建议 --width | 建议传入方式 | 说明 |
| --- | --- | --- | --- |
| DeepSeek V4 **Flash** (轻量/高速) | 1280 | 优先 `contact_sheet.png` 单图，或分页 | 视觉预算小，小图更快更稳；长段文字可能被弱模型漏读，关键结论放标题/短句 |
| DeepSeek V4 **Pro** (强视觉) | 2000–2400 | 分页或拼图皆可 | 细节识别强，可用大图看排版/溢出 |
| 通义千问 Qwen-VL / 智谱 GLM-4V / 阶跃 Step-1V 等国产 VLM | 1600–2000 | 分页 | 同为本地产视觉模型，直接把 PNG 当图片传入即可 |
| Claude / GPT-4V 等 | 2000 | 分页或拼图 | 通用做法 |

注意：渲染器只管"把 pptx 变成图"，**不调用任何模型、不发网络请求**。适配不同模型只需在调用时调整 `--width` 和选择传哪张图，无需改动脚本。

## How the renderer works (what it reads)

The script unzips the `.pptx` (it is a ZIP of XML) and for each
`ppt/slides/slideN.xml`:

1. Reads `ppt/presentation.xml` for slide size, and `ppt/theme/themeN.xml` for the
   color scheme.
2. Walks the shape tree (`sp`, `pic`, `graphicFrame`), drawing each with its
   `xfrm` geometry, solid fill / line, and text runs (size, bold, color, align).
3. Composites embedded `<pic>` media (backgrounds, photos) from the slide's
   `.rels` so full-bleed images show up.
4. Wraps CJK + Latin text per run and compares total text height to the box
   height to flag overflow.

## Font handling (cross-platform)

Text needs a CJK font to render Chinese/Japanese/Korean. The script auto-detects
a usable one at runtime — no bundling required:

- macOS: `STHeiti Medium.ttc`, `PingFang.ttc`, `Arial Unicode.ttf`
- Windows: `msyh.ttc` (Microsoft YaHei), `simhei.ttf`, `simsun.ttc`
- Linux: Noto Sans CJK (`NotoSansCJK-Regular.ttc`), WenQuanYi (`wqy-*.ttc`)

If none is found it prints a WARNING and CJK glyphs render as boxes — install any
CJK font and re-run. To force a specific font, edit `FONT_PATH` / `FONT_INDEX` at
the top of `scripts/render_pptx.py`.

## Limitations

- Approximate: spacing, kerning, and exact wrapping differ from PowerPoint.
- Does not render charts, SmartArt, transitions, animations, or grouped/rotated
  shapes precisely (they are approximated as bounding boxes).
- Tables render as plain text blocks, not gridded cells.
- Use overflow/bounds flags as a conservative signal, not as a guarantee.

## Resources

### scripts/render_pptx.py
The renderer. Standalone, no imports beyond the standard library + Pillow.
Accepts a pptx path plus an optional output dir, and optional flags
`--width N` (output px width, default 2000) and `--no-contact` (skip the combined
sheet). Writes one PNG per slide plus a contact_sheet.png.
