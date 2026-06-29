# 当前交接说明

更新日期：2026-06-28

## 项目状态

这个项目已经从“生成一页第 09 页结果分析 PPT”扩展出多个实验方向：工业白皮书静态页、PowerPoint 触发器动画测试、Keynote 风格复刻、WPS/HTML 承载方案，以及 GordenImage2PPTX 细节页重建路线。继续动手前应先冻结目标，不要直接在旧输出上继续试错。

## 已确认的工作原则

- 当前最明确的主线是“工业白皮书风格”的第 09 页结果分析页。
- 第 09 页的六个变量是温度、压强、金属材质、表面粗糙度、接触面积、冰的成分；第二个变量必须是“压强 Pressure”。
- 标题、变量名、数据表、图表标签、结论文字必须可编辑。
- 图表应优先使用原生 PowerPoint 图表或可编辑形状，不应做成死图。
- 图片适合承担背景纹理、实验场景、金属/冰质感和非关键信息装饰。
- 若使用图片生成资产，必须留下真实 imagegen 证据；不能用程序图假冒 imagegen 输出。

## 主要分支

- 工业白皮书分支：主提示词在 `projects/slide09_whitepaper/sources/提示词.md`，验证记录在 `projects/slide09_whitepaper/validation/slide09_whitepaper_page09_interactive.md`。
- Gorden 严格细节页路线：待办在 `gorden_detail_pages_todo.md`，目标是先把六个高质量静态参考图还原成可编辑 PPTX，再集成最终全因素动画。
- WPS/HTML 承载路线：最新输出集中在 `outputs/wps_factor_matrix_*`，适合作为演示或粘贴承载，但不等同于可编辑 PowerPoint 原生交互。
- Keynote 风格分支：已有若干预览和扩图提示词，但不应和工业白皮书分支混用。

## 下一步建议

1. 如果目标是最终可交付 PPTX，先完成 `gorden_detail_pages_todo.md` 中六个严格 Gorden 细节页，再运行 `build_factor_matrix_all_animations_from_gorden_pages.ps1`。
2. 如果目标是快速演示给用户看，继续使用 `outputs/wps_factor_matrix_interactive_carrier_v1.pptx` 这类承载版，但要明确它不是最终可编辑原生交互版本。
3. 如果用户对当前结果不满意，先重写并确认主提示词，不要继续生成更多版本。

## 已知风险

- `AGENTS.md` 已经接近规则手册上限，后续不要把历史过程、单次失败或输出流水账继续写进去。
- `projects/slide09_whitepaper_static_v2` 到 `v9` 是连续探索产物，不能默认其中某个就是当前基线。
- `tmp/` 与 `.codexbridge/` 里有大量临时和桥接产物，不应作为交付源。
- 真正的 PowerPoint 触发器动画和 WPS/HTML 交互不是同一类实现，最终回复必须区分。

## 核验入口

- 文本与数据核验：`projects/slide09_whitepaper/validation/slide09_whitepaper_page09_interactive.md`
- 最新 WPS/HTML 输出：`outputs/wps_factor_matrix_interactive_carrier_v1.pptx`
- 严格 Gorden 集成脚本：`build_factor_matrix_all_animations_from_gorden_pages.ps1`
