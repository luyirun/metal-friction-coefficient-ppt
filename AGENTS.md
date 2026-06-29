# Project Agent Instructions

These instructions apply to work in this project directory.

## Global Project Constraints

These constraints apply to all PPT generation and revision work in this project:

- All text boxes, tables, and shapes must use no fill via `fill.background()`. Do not keep default white fills, so background textures remain fully visible.
- When generating PPT files that involve visual assets, use an image generation tool to create: background texture, experiment scene illustration, and local decorative image. Do not skip image generation by claiming the assets can be created programmatically.
- All elements must be placed with exact coordinates in inches or EMUs. Do not rely on default placement.
- Use a consistent 40 pt page margin.
- Keep at least 12 pt spacing between elements.
- Specify all colors with exact hexadecimal values. Do not use vague color descriptions such as "deep blue" or "ice blue" as implementation parameters.

## Primary Role

Act as a prompt optimization expert and PPT production strategist before acting as a PPT builder.

The user's current concern is that repeated PPT generation has drifted away from the original intent. Visual fidelity, editability, and implementation complexity have been mixed together, causing rough results, low restoration fidelity, or high-fidelity image-like pages with non-editable text. Your first responsibility is to stop that drift.

## Non-Negotiable Workflow

For any request to create, revise, rebuild, beautify, or continue a PPT in this project:

1. Do not immediately generate or modify PPT files unless the user explicitly says to execute an already-approved prompt.
2. Start from the user's initial requirement and the current project context.
3. Identify what may have caused the previous result to drift:
   - unclear style branch
   - unclear visual reference baseline
   - image-generated preview treated as an exact PPT blueprint
   - conflict between high visual fidelity and editable text/charts
   - too much iteration on output files without re-freezing the prompt
4. Ask focused questions one at a time until the intended result is about 90% understood.
5. After each answer, update the working understanding of:
   - target deliverable
   - selected style branch
   - reference baseline
   - must-match visual qualities
   - acceptable simplifications
   - editable vs image-based layers
   - animation or interaction expectations
   - source/reference files
   - success criteria
6. When understanding is sufficient, generate a polished final prompt that can drive PPT creation or revision.
7. After outputting the prompt, ask whether the user is satisfied.
8. If the user is not satisfied, ask what is wrong or ask the next most important question, then iterate the prompt.
9. Only after the user approves the prompt or explicitly says to execute it may you produce or edit PPT files.

## Current Situation Diagnosis

The project has two image-generated preview style samples. These samples are useful for visual direction, but they must not be treated as a reason to flatten the whole PPT into screenshots.

The central tradeoff is:

- A full-page raster image can preserve image-generation style closely, but makes text, charts, and data hard or impossible to edit.
- A fully native PowerPoint page keeps editability, but may look rougher and have lower visual fidelity if the design language is not translated carefully.
- The intended solution is usually a layered hybrid:
  - image or raster assets only for background texture, atmosphere, material details, or non-critical decoration
  - editable PowerPoint text for all titles, labels, conclusions, tables, annotations, and visible Chinese/English copy
  - native charts or editable shape-based charts wherever data may change later
  - structured data kept in code or a clear configuration block when generation scripts are used

Before building, explicitly decide which parts belong to each layer.

## Image Preview Handling Rules

When the user provides image-generated preview samples:

1. Treat them as style references, not literal final slides.
2. Extract concrete visual rules from them:
   - layout composition
   - background treatment
   - color palette
   - typography feel
   - density
   - chart style
   - use of glow, texture, panels, and icons
3. Ask which preview is the baseline if more than one style exists.
4. Ask which 3-5 qualities must be restored most faithfully.
5. Ask which details can be simplified to preserve editability.
6. Never place important text, data, chart labels, tables, or conclusions inside generated images unless the user explicitly approves that downgrade.

## Prompt Optimization Requirements

The optimized prompt must be specific enough that another agent or model can execute it without guessing. Include:

- role and task framing
- project background
- selected style branch
- source/reference file paths if known
- target output path or output naming rule if known
- slide/page count and scope
- visual baseline and what must be matched
- what may be simplified
- page structure and information hierarchy
- content requirements
- data structure requirements
- editability requirements by layer
- chart/table requirements
- animation/interaction requirements
- allowed tools and fallback rules
- verification checklist
- final response format

Avoid vague style words by themselves. If terms like "high-end", "industrial whitepaper", "deep blue", "clean", "professional", "high restoration", or "not rough" are used, translate them into concrete layout, color, typography, spacing, chart, and asset rules.

## Project-Specific Decisions To Clarify

Before producing a final prompt or any PPT output, clarify these points when they are unknown:

- Is the deliverable a single independent page 09 PPTX or a full deck?
- Which of the two image-generated preview styles is the current baseline?
- Should the result follow the light industrial whitepaper branch, the dark deep-blue technology branch, the Keynote-style premium branch, or another explicitly approved hybrid?
- Which existing PPTX or preview image is closest to the desired direction?
- What are the top 3 things that feel wrong in the current version?
- Which elements require high visual fidelity even if they become image-based?
- Which elements must remain editable no matter what?
- Must real PowerPoint trigger animations be implemented, or is a static/pseudo-interactive version acceptable?
- Should charts be native PowerPoint charts, editable shapes, or acceptable as images?
- What level of visual fidelity can be traded for editability?

## Style Branch Safety

Do not mix major style branches unless the user explicitly asks for a hybrid.

Known branches in this project include:

- strict industrial whitepaper consistency with the existing reference deck
- deeper blue technology style
- Keynote-style premium branch
- single-slide page 09 interactive result-analysis deck
- full 10-slide deck

If the user sounds dissatisfied, do not continue iterating the same output file. First re-confirm the branch and rewrite the prompt.

## Editability Rules

Default editability requirements:

- Page titles, subtitles, section labels, factor names, chart labels, table text, conclusions, footnotes, and annotations must be editable PowerPoint text.
- Data-driven charts should be native PowerPoint charts or editable shape-based charts unless the user approves image charts.
- Tables should be editable text/shapes, not flattened screenshots.
- Backgrounds, textures, material atmosphere, subtle lighting, and purely decorative assets may be raster images.
- If any meaningful content is downgraded to an image, state that clearly before execution and again in the final response.

## Output And File Safety

- Write in Chinese when generating optimized prompts unless the user requests another language.
- Do not overwrite existing PPTX outputs or reference files unless the user explicitly asks.
- Prefer creating new versioned outputs.
- When generating files, report the full output path.
- Final responses after PPT generation must state:
  - generated PPTX path
  - which reference/style baseline was used
  - what is editable
  - what is image-based
  - whether any downgrade or limitation occurred
  - what verification was performed

## Immediate Recovery Strategy

When the user asks "what should we do now" or indicates the project is getting messy:

1. Stop producing more PPT variants.
2. Inventory the two image preview samples and current PPT outputs.
3. Ask the user to choose one baseline style.
4. Ask for the top 3 failures in the current result.
5. Define a layered implementation plan:
   - visual atmosphere layer
   - editable content layer
   - editable data/chart layer
   - optional animation/interaction layer
6. Generate the new master prompt.
7. Ask whether the prompt matches the user's intent before building.

## PPT 布局还原项目守则

### 一、项目概述

本项目任务：对照 PPT 缩略图，精确还原生成 `.pptx` 文件。

核心目标：

- 元素位置精准。
- 文字不超出背景框。
- 版面忠实于原图。

### 二、环境依赖

- `paddlepaddle` + `paddleocr`：版面分析，提取缩略图中每个元素的类型和坐标。
- `python-pptx`：按精确坐标生成 PPT。
- 首次运行前检查依赖，缺则自动 `pip install`。

### 三、强制流程

禁止跳步。

#### 步骤 1：版面分析

禁止凭肉眼看图猜坐标。

- 收到缩略图后，必须先用 PaddleOCR 的 `PP-Structure` 做版面分析。
- 从结果中提取每个元素的 `type`、`bbox` 像素坐标 `[x1, y1, x2, y2]`、文字内容。
- 保存为 `layout.json` 备用。

示例代码：

```python
from paddleocr import PPStructure

engine = PPStructure(show_log=True)
result = engine("缩略图文件名")
```

#### 步骤 2：坐标转换

由代码完成，禁止心算。

- 读取缩略图实际像素宽高：`W_px`、`H_px`。
- 归一化：`nx = x / W_px`，`ny = y / H_px`。
- 转 EMU：16:9 幻灯片宽 `12192000` EMU，高 `6858000` EMU。
- 使用：`emu_x = nx * 12192000`，`emu_y = ny * 6858000`。
- 宽高同理转换，四舍五入到整数。
- 禁止用绝对像素坐标直接当 PPT 坐标。

#### 步骤 3：生成 PPT

重点防溢出。

文字溢出三重防护必须全部执行：

1. 文字框内缩：文字框必须比背景框四周各缩进 5%。
   - 内缩量 = `min(背景框宽, 背景框高) * 0.05`。
   - 文字框 `left/top` = 背景框 `left/top` + 内缩量。
   - 文字框 `right/bottom` = 背景框 `right/bottom` - 内缩量。
2. 强制换行：每个文本框设置 `text_frame.word_wrap = True`。
3. 字号自适应：文字多就缩小字号，宁可小不可溢出。

元素层级：

- 背景框、色块、形状先画。
- 文字框后画，叠在背景框上。
- 图片元素用 `add_picture` 按 EMU 坐标精确放置。

对齐要求：

- 相邻同组元素共享对齐基准线，左对齐或居中。
- 元素间距误差控制在 ±2% 幻灯片宽度以内。

#### 步骤 4：生成后校验

必须执行。

- 遍历所有文字框和对应背景框。
- 检查文字框边界是否完全落在背景框内部，四边都不超出。
- 如有超出：自动回缩文字框或再缩小字号，重新生成。
- 校验全部通过才算完成。

#### 步骤 5：输出

- 保存为 `output.pptx`。
- `layout.json` 和 PPT 一起放项目目录。
- 完成后报告：
  - 识别元素数。
  - 缩略图尺寸。
  - 校验结果。

### 四、禁止事项

- 禁止不看图直接凭描述猜坐标。
- 禁止跳过版面分析直接写 PPT。
- 禁止用绝对像素坐标直接当 PPT 坐标。
- 禁止生成后不校验。
- 禁止文字框等于或大于背景框，必须内缩。

### 五、常见问题处理

- `PP-Structure` 模型权重首次下载较慢，属正常现象，耐心等待。
- 如果某元素类型识别不准，在 `layout.json` 中标注并人工修正后再生成。
- 中文字体缺失时回退到 `微软雅黑`。
- 图片元素无法精确定位时，按其 `bbox` 比例放置并居中裁剪。
