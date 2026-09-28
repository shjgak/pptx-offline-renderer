# PPTX Offline Renderer · PPTX 离线渲染器

[English](#english) | [中文](#%E4%B8%AD%E6%96%87)

---

## English

Render a `.pptx` deck into one PNG per slide, **entirely offline** — pure Python + Pillow plus the OS's built-in CJK font. No PowerPoint, Keynote, or LibreOffice, no network. Built so an AI agent (or a human) can "see" a deck and do visual QA: layout, colors, shapes, embedded pictures, text overflow.

The renderer is **model-agnostic**: it only writes PNG files. Attach `slide_N.png` or `contact_sheet.png` to any vision-capable model — DeepSeek V4 Flash / Pro, Qwen-VL, GLM-4V, Step-1V, Claude, GPT-4V.

### Install

As an Agent Skill (Codex, Claude Code, Cursor — anything that reads SKILL.md):

```bash
npx skills add shjgak/pptx-offline-renderer -g -a codex -y
```

Script only:

```bash
git clone https://github.com/shjgak/pptx-offline-renderer.git
cd pptx-offline-renderer
pip install pillow
```

### Usage

```bash
python scripts/render_pptx.py deck.pptx out/                 # slides + contact sheet
python scripts/render_pptx.py deck.pptx out/ --width 1280    # small images for lightweight VLMs
python scripts/render_pptx.py deck.pptx out/ --width 2400    # large images for strong VLMs
python scripts/render_pptx.py deck.pptx out/ --no-contact    # per-slide PNGs only, no sheet
```

Outputs `slide_1.png`, `slide_2.png`, … plus a `contact_sheet.png` (all slides side by side).

### What it checks

- `OVERFLOW` — rendered text is taller than its text box (it would clip in PowerPoint too)
- `BOUNDS` — a shape extends past the slide edge
- `PIC` — an embedded picture failed to composite

Fix the source `.pptx` (shorten text, enlarge boxes, re-embed images) instead of treating these as cosmetic.

### Choosing `--width`

| Target model | Suggested `--width` | Suggested input |
| --- | --- | --- |
| DeepSeek V4 Flash (fast, light) | 1280 | `contact_sheet.png` first, or one slide per image |
| DeepSeek V4 Pro (strong vision) | 2000–2400 | either |
| Qwen-VL / GLM-4V / Step-1V | 1600–2000 | one slide per image |
| Claude / GPT-4V | 2000 | either |

The renderer never calls a model and never touches the network. Tuning for a model means changing `--width` and which image you attach.

### Repo layout

```
SKILL.md                  Agent Skill definition (instructions for the AI)
scripts/render_pptx.py    The renderer (stdlib + Pillow only)
```

### Limitations

- Approximate: spacing, kerning and wrapping differ from PowerPoint
- Charts, SmartArt, animations, grouped/rotated shapes render as bounding boxes
- Tables render as plain text blocks, not grids
- Overflow/bounds warnings are a conservative signal, not a guarantee

---

## 中文

在没有安装 PowerPoint / Keynote / LibreOffice 的环境（无显卡服务器、CI、没有 Office 的 Mac）里，把 `.pptx` 逐页渲染成 PNG 图片。纯 Python + Pillow + 系统自带中文字体，离线完成，不联网、不依赖任何办公软件，让 AI 或人能"看见"幻灯片，从而做视觉质检（版面、配色、形状、内嵌图片、文字溢出）。

渲染器是**模型无关**的：它只产出图片。DeepSeek V4 Flash / Pro、通义千问 Qwen-VL、智谱 GLM-4V、阶跃 Step-1V、Claude、GPT-4V 等任意带视觉能力的模型，把生成的 `slide_N.png` 或 `contact_sheet.png` 作为图片传入即可使用。

### 安装

作为 Agent Skill（Codex / Claude Code / Cursor 等任何能读 SKILL.md 的运行时）：

```bash
npx skills add shjgak/pptx-offline-renderer -g -a codex -y
```

只要脚本也行：

```bash
git clone https://github.com/shjgak/pptx-offline-renderer.git
cd pptx-offline-renderer
pip install pillow
```

### 用法

```bash
# 逐页渲染 + 生成拼图
python scripts/render_pptx.py deck.pptx out/

# 轻量视觉模型（Flash 类）：小图更快更稳
python scripts/render_pptx.py deck.pptx out/ --width 1280

# 强视觉模型（Pro 类）：大图看细节
python scripts/render_pptx.py deck.pptx out/ --width 2400

# 只要分页图，不要拼图
python scripts/render_pptx.py deck.pptx out/ --no-contact
```

输出 `slide_1.png`、`slide_2.png`… 外加一张 `contact_sheet.png`（所有页拼在一张，适合小视觉预算的模型）。

### 自动检测的问题

- `OVERFLOW` —— 文字超出文本框（在真实 PowerPoint 里也会被裁掉）
- `BOUNDS` —— 形状越界
- `PIC` —— 内嵌图片合成失败

这些要在源 `.pptx` 里修（缩短文字、放大文本框、重新嵌入图片），不要当成小瑕疵忽略。

### 不同视觉模型怎么调

| 目标模型 | 建议 --width | 建议传入方式 |
| --- | --- | --- |
| DeepSeek V4 Flash（轻量/高速） | 1280 | 优先 `contact_sheet.png` 单图，或分页 |
| DeepSeek V4 Pro（强视觉） | 2000–2400 | 分页或拼图皆可 |
| Qwen-VL / GLM-4V / Step-1V 等国产 VLM | 1600–2000 | 分页 |
| Claude / GPT-4V 等 | 2000 | 分页或拼图 |

渲染器本身**不调用任何模型、不发网络请求**。适配不同模型只需调整 `--width` 和选择传哪张图，无需改脚本。

### 仓库结构

```
SKILL.md                  Agent Skill 定义（给 AI 用的说明书）
scripts/render_pptx.py    渲染器本体（只依赖标准库 + Pillow）
```

### 限制

- 近似渲染：字距、断行与 PowerPoint 不完全一致
- 图表、SmartArt、动画、组合/旋转形状只能近似为包围盒
- 表格按纯文本块渲染，不是网格
- 溢出/越界告警是保守信号，不构成保证
