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
