PPTX 离线渲染器（pptx-offline-renderer）

在没有安装 PowerPoint / Keynote / LibreOffice 的环境（无显卡服务器、CI、没有 Office 的 Mac）里，把 .pptx 逐页渲染成 PNG 图片。它用纯 Python + Pillow 和系统自带的中文字体离线完成，不联网、不依赖任何办公软件，让 AI 或人能"看见"幻灯片，从而做视觉质检（版面、配色、形状、内嵌图片、文字溢出）。

渲染器是模型无关的：它只产出图片，DeepSeek V4 Flash / Pro、通义千问 Qwen-VL、智谱 GLM-4V、阶跃 Step-1V、Claude、GPT-4V 等任意带视觉能力的模型，把生成的 slide_N.png 或 contact_sheet.png 作为图片传入即可使用。

核心能力：

离线渲染：解包 pptx（本质是 ZIP+XML），按形状树绘制几何、纯色填充/描边、文字（字号/加粗/颜色/对齐）与内嵌图片；
溢出检测：自动告警 OVERFLOW（文字超出文本框）、BOUNDS（形状越界）、PIC（图片缺失），帮你在交付前修掉排版事故；
跨平台中文字体自动探测（macOS / Windows / Linux），无需打包字体；
--width 分辨率可调：轻量视觉模型（Flash 类）用 1280 小图，强视觉模型（Pro 类）用 2400 大图。
