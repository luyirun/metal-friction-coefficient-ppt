from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parent
REFERENCE = ROOT / "outputs" / "imagegen_ppt_collage_05_industrial_whitepaper.png"
OUT = ROOT / "outputs" / "whitepaper_preview_rebuild"
REFS = OUT / "reference_crops"
PROMPTS = OUT / "prompts"


COMMON_PROMPT = """Use case: productivity-visual
Asset type: single 16:9 PowerPoint preview image for an industrial whitepaper deck
Primary request: Rebuild one standalone 16:9 slide preview for the deck "金属与冰摩擦系数的测量". The input image is a small reference thumbnail from style sample A. Preserve the industrial whitepaper visual system, layout logic, page number system, and refined engineering-report mood, but make the slide cleaner, sharper, more spacious, and more professionally typeset.
Input image role: visual reference for composition, hierarchy, color palette, density, atmosphere, and page-specific layout only. Do not treat distorted AI text as final content truth.
Style/medium: light industrial engineering whitepaper, premium physics experiment report, warm white canvas, subtle ice-blue atmosphere, brushed aluminum gray, deep teal accents, small muted orange highlights.
Composition/framing: exactly one landscape 16:9 slide, full-frame, no collage, no surrounding margins, no extra pages.
Typography direction: serious Chinese sans-serif presentation typography; large clean title, compact body labels, consistent page number block. Text may be approximate in this preview stage; final PPT will rebuild all text as editable PowerPoint text.
Visual system: deep teal top-left page number block, clear title system, fine divider lines, subtle low-opacity shadows, restrained white cards, technical icons with consistent stroke weight, precise apparatus/diagram/chart treatment, generous whitespace.
Image strategy: apparatus and ice/ship scenes may be AI-generated placeholders. Do not bake important final text, formulas, data labels, or conclusions into decorative images.
Avoid: Keynote/premium dark style, deep-blue technology dashboard, neon, dark background, poster look, low-end template, cartoon, clutter, thick borders, heavy shadows, excessive gradients, watermark, 2x5 grid.
Quality target: high visual fidelity to sample A, not rough, not chaotic, formal industrial report polish.
"""


SLIDES = [
    (
        "01",
        "封面",
        """Main visible text to use where legible: 金属与冰摩擦系数的测量; 物理实验创新竞赛答辩; 日期：2024.05.20.
Layout: title and small metadata on the left, large refined inclined-plane ice friction apparatus on the right, circular close-up inset, pale polar mountain/ice atmosphere in the background, bottom teal discipline band.""",
    ),
    (
        "02",
        "选题背景",
        """Main visible text to use where legible: 选题背景; 北极航运与冰阻力问题; 25%-55%; 13%-15%; 研究意义.
Layout: large ship-in-ice image on the left, numeric orange callouts and short explanatory blocks on the right, smaller ice texture image near lower right, clean whitepaper panels.""",
    ),
    (
        "03",
        "题目解读",
        """Main visible text to use where legible: 题目解读; 机理分析; 装置搭建; 结果与不确定度.
Layout: three balanced vertical columns, each with a circular technical icon and concise bullet list, fine dividers, ample white space.""",
    ),
    (
        "04",
        "方法选择",
        """Main visible text to use where legible: 方法选择 -- 动态斜面法; 力学模型; 理论公式; a = g(sinθ - μcosθ); μ = (sinθ - a/g) / cosθ.
Layout: force diagram of block on inclined plane on the left, formula cards and symbol explanation on the right, restrained teal/orange annotation lines.""",
    ),
    (
        "05",
        "装置设计",
        """Main visible text to use where legible: 装置设计; 斜面冰面模块; 垂直冰面模块（压力加载）.
Layout: two apparatus renderings side by side, left inclined ice plane system with callout labels, right vertical pressure loading module, bottom row of four feature icons.""",
    ),
    (
        "06",
        "变量设计",
        """Main visible text to use where legible: 变量设计; 六大影响变量与水平设置; 温度; 压强; 材质; 粗糙度; 面积; 成分.
Layout: six evenly spaced circular icon modules across the slide, each with 2-3 concise parameter levels underneath, pale blue circles and clean teal labels.""",
    ),
    (
        "07",
        "问题改进",
        """Main visible text to use where legible: 问题改进; 冰面制备不平整; 低温环境稳定性差; 传感器可靠性不足; 改进方案.
Layout: three problem rows on the left flowing via arrows to three solution panels on the right, green check marks, restrained engineering-process style.""",
    ),
    (
        "08",
        "数据处理",
        """Main visible text to use where legible: 数据处理; 数据采集; 加速度计算; 摩擦系数计算; 不确定度评估; μ = (sinθ - a/g) / cosθ.
Layout: four process cards connected by arrows, simple line-chart/wave/formula/normal-distribution icons, formula card emphasized but still clean.""",
    ),
    (
        "09",
        "结果分析",
        """Main visible text to use where legible: 结果分析; 摩擦系数对比趋势（示例）; 数据解读; 结果与文献趋势一致，验证方法的有效性.
Layout: large clean line chart on the left with multiple material trends, compact interpretation bullets on the right, small conclusion callout at lower right.""",
    ),
    (
        "10",
        "总结展望",
        """Main visible text to use where legible: 总结展望; 创新点; 工程应用前景.
Layout: structured summary cards on the left, stacked polar ship/ice/apparatus images on the right, calm final-report closing band at the bottom.""",
    ),
]


def crop_reference_thumbnails() -> None:
    image = Image.open(REFERENCE).convert("RGB")
    width, height = image.size

    cols, rows = 2, 5
    cell_w, cell_h = width / cols, height / rows
    margin_x = width * 0.012
    margin_y = height * 0.010

    REFS.mkdir(parents=True, exist_ok=True)
    for index in range(10):
        row = index // cols
        col = index % cols
        left = int(col * cell_w + margin_x)
        top = int(row * cell_h + margin_y)
        right = int((col + 1) * cell_w - margin_x)
        bottom = int((row + 1) * cell_h - margin_y)
        crop = image.crop((left, top, right, bottom))
        crop.save(REFS / f"page_{index + 1:02d}_reference.png")


def write_prompts() -> None:
    PROMPTS.mkdir(parents=True, exist_ok=True)
    for page, title, content in SLIDES:
        prompt = f"""{COMMON_PROMPT}
Slide-specific task:
Rebuild page {page} as a single polished 16:9 preview image.
Page title: {title}
{content}
Preview-stage text rule: keep the main title, page number, key headings, and major numeric callouts readable where possible. Small body text may be approximate because final PPTX will rebuild all text, charts, formulas, labels, and conclusions as editable objects.
Output constraints: one slide only, 16:9 landscape, no collage, no thumbnail grid, no extra slide, no watermark.
"""
        (PROMPTS / f"page_{page}_prompt.txt").write_text(prompt, encoding="utf-8")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    crop_reference_thumbnails()
    write_prompts()
    print(OUT)


if __name__ == "__main__":
    main()
