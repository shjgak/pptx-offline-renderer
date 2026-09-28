# PPTX 离线渲染器（pptx-offline-renderer）

在没有安装 PowerPoint / Keynote / LibreOffice 的环境（无显卡服务器、CI、没有 Office 的 Mac）里，把 `.pptx` 逐页渲染成 PNG 图片。纯 Python + Pillow + 系统自带中文字体，离线完成，不联网、不依赖任何办公软件，让 AI 或人能"看见"幻灯片，从而做视觉质检（版面、配色、形状、内嵌图片、文字溢出）。

渲染器是**模型无关**的：它只产出图片。DeepSeek V4 Flash / Pro、通义千问 Qwen-VL、智谱 GLM-4V、阶跃 Step-1V、Claude、GPT-4V 等任意带视觉能力的模型，把生成的 `slide_N.png` 或 `contact_sheet.png` 作为图片传入即可使用。

## 安装

作为 Agent Skill（Claude Code / Codex / Cursor 等任何能读 SKILL.md 的运行时）：

```bash
npx skills add shjgak/pptx-offline-renderer -g -a codex -y
```

只要脚本也行：

```bash
git clone https://github.com/shjgak/pptx-offline-renderer.git
cd pptx-offline-renderer
pip install pillow
```

## 用法

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

## 核心能力

- **离线渲染**：解包 pptx（本质是 ZIP+XML），按形状树绘制几何、纯色填充/描边、文字（字号/加粗/颜色/对齐）与内嵌图片
- **溢出检测**：自动告警 `OVERFLOW`（文字超出文本框）、`BOUNDS`（形状越界）、`PIC`（图片缺失），交付前就能发现排版事故
- **中文字体自动探测**：macOS / Windows / Linux 通吃，无需打包字体
- **分辨率可调**：`--width` 控制输出尺寸，轻量视觉模型用 1280，强视觉模型用 2400

## 不同视觉模型怎么调

| 目标模型 | 建议 --width | 建议传入方式 |
| --- | --- | --- |
| DeepSeek V4 Flash（轻量/高速） | 1280 | 优先 `contact_sheet.png` 单图，或分页 |
| DeepSeek V4 Pro（强视觉） | 2000–2400 | 分页或拼图皆可 |
| Qwen-VL / GLM-4V / Step-1V 等国产 VLM | 1600–2000 | 分页 |
| Claude / GPT-4V 等 | 2000 | 分页或拼图 |

渲染器本身**不调用任何模型、不发网络请求**。适配不同模型只需调整 `--width` 和选择传哪张图，无需改脚本。

## 仓库结构

```
SKILL.md                  Agent Skill 定义（给 AI 用的说明书）
scripts/render_pptx.py    渲染器本体（只依赖标准库 + Pillow）
```

## 限制

- 近似渲染：字距、断行与 PowerPoint 不完全一致
- 图表、SmartArt、动画、组合/旋转形状只能近似为包围盒
- 表格按纯文本块渲染，不是网格
- 溢出/越界告警是保守信号，不构成保证
