from pathlib import Path

ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
OUT = ROOT / "outputs" / "single_page_prompts"
OUT.mkdir(parents=True, exist_ok=True)

common = """Use case: productivity-visual
Asset type: single high-resolution 16:9 PowerPoint slide visual draft
Primary request: Expand the provided reference thumbnail into one complete high-definition 16:9 PowerPoint slide. This is not a collage. Generate only this single slide. Strictly preserve the selected moodboard style: industrial whitepaper / engineering report, warm white canvas, graphite text, deep teal accents, muted orange highlights, pale ice blue details, clean technical icons, thin borders, subtle shadows, bottom-right page number.
Input image role: The input image is the exact visual reference for this slide's layout, visual hierarchy, content structure, and style. Recreate it as a sharper, cleaner, higher-resolution single slide, improving details while keeping the same arrangement.
Visual system constraints: serious Chinese sans-serif title hierarchy, refined engineering apparatus/CAD imagery, clean charts/diagrams, restrained academic presentation, no dark dashboard style, no cartoon style, no extra pages, no collage, no watermark.
Text constraints: Main title, core labels, key numbers, formulas, chart labels, and page number must be clear and readable. If tiny reference text is unclear, use the provided verified text below. Do not invent extra facts or precise data.
"""

slides = [
    {
        "title": "金属与冰摩擦系数的测量",
        "page": "01",
        "body": """Verified content:
Title: 金属与冰摩擦系数的测量
Core point: 物理实验创新竞赛答辩
Small labels: 团队：祝冰间隙者; 日期：2024.05.20
Visual: large clean apparatus render: inclined transparent ice plane, metal slider block, sensor/measurement module, small circular close-up inset. Keep the composition from the reference: title on left, apparatus on right, footer band at bottom.""",
    },
    {
        "title": "选题背景",
        "page": "02",
        "body": """Verified content:
Title: 选题背景
Section heading: 北极航运与冰阻力问题
Key numbers: 25%-55%; 13%-15%
Core point: 冰上摩擦影响极地航运阻力与能耗
Visual: two polar ship/ice images on the left and bottom-right, key numeric callouts on the right, short whitepaper text blocks, footer page number 02.""",
    },
    {
        "title": "题目解读",
        "page": "03",
        "body": """Verified content:
Title: 题目解读
Three columns: 机理分析; 装置搭建; 结果与不确定度
Column points:
机理分析: 冰的黏附性变弱; 表面微观水膜; 水膜润滑作用; 温度与压力影响
装置搭建: 斜面摩擦实验平台; 力学测量与采集; 低温环境控制; 稳定与校准
结果与不确定度: 摩擦系数计算; 不确定度评估; 数据性分析; 结果讨论与验证
Visual: three-column whitepaper layout with clean circular icons.""",
    },
    {
        "title": "方法选择 —— 动态斜面法",
        "page": "04",
        "body": """Verified content:
Title: 方法选择 —— 动态斜面法
Left heading: 力学模型
Right heading: 理论公式
Formula labels: a = g(sinθ - μcosθ); μ = (sinθ - a/g) / cosθ
Symbol explanation: a, g, θ, μ
Visual: force diagram of block sliding down inclined plane on left, formula cards on right, clean teal/orange accents, page number 04.""",
    },
    {
        "title": "装置设计",
        "page": "05",
        "body": """Verified content:
Title: 装置设计
Two modules: 斜面冰面模块; 垂直冰面模块（压力加载）
Labels around apparatus: 导轨与滑块; 力传感器; 位移传感器; 冰面板; 温控仓体; 角度调节机构; 支撑框架
Bottom feature labels: 模块化设计; 快速更换; 低温环境模拟; 高精度测量
Visual: two refined grayscale/teal CAD-style apparatus renders, left inclined module, right vertical module, technical callout labels.""",
    },
    {
        "title": "变量设计",
        "page": "06",
        "body": """Verified content:
Title: 变量设计
Subtitle: 六大影响变量与水平设置
Variables: 温度 -5℃ -10℃ -15℃; 压强 0.05 MPa 0.10 MPa 0.20 MPa; 材质 铝合金 不锈钢 钛合金; 粗糙度 Ra 0.2 μm Ra 1.0 μm Ra 3.0 μm; 面积 1 cm² 4 cm² 9 cm²; 成分 淡水冰 海冰 人工掺盐冰
Visual: six circular icon modules across the slide, light blue circles, teal text, consistent spacing, page number 06.""",
    },
    {
        "title": "问题改进",
        "page": "07",
        "body": """Verified content:
Title: 问题改进
Rows:
冰面制备不平整 → 改进方案: 定制制冰机 + 控温体系; 实现光滑透明平整冰面
低温环境稳定性差 → 改进方案: 双层保温 + PID温控; 温度波动 ≤ ±0.2℃
传感器可靠性不足 → 改进方案: 低温型传感器 + 屏蔽布线; 多点校准与滤波处理
Visual: three problem-solution rows with icons, arrows, green check marks, calm engineering report style, page number 07.""",
    },
    {
        "title": "数据处理",
        "page": "08",
        "body": """Verified content:
Title: 数据处理
Four stages: 数据采集; 加速度计算; 摩擦系数计算; 不确定度评估
Stage labels:
数据采集: 力信号, 位移信号, 温度/压力, 采样频率100 Hz
加速度计算: 位移二次微分, 滤波处理, 得到a(t)
摩擦系数计算: μ = (sinθ - a/g) / cosθ; 代入角度与加速度; 计算μ
不确定度评估: A类不确定度; B类不确定度; 合成标准不确定度; 扩展不确定度U
Visual: four process cards connected by arrows, formula card emphasized, page number 08.""",
    },
    {
        "title": "结果分析",
        "page": "09",
        "body": """Verified content:
Title: 结果分析
Chart heading: 摩擦系数对比趋势（示例）
Chart: line chart of 摩擦系数 μ vs 温度/℃ with multiple materials: 铝合金, 不锈钢, 铁合金, 钛合金. Use approximate illustrative trends only, do not claim exact measured values.
Right heading: 数据解读
Points: 温度升高，μ整体下降; 粗糙度增大，μ上升; 压强增大，先升后降; 材质差异显著，钛合金综合表现最优
Callout: 结果与文献趋势一致，验证方法的有效性
Visual: large clean chart left, text interpretation right, page number 09.""",
    },
    {
        "title": "总结展望",
        "page": "10",
        "body": """Verified content:
Title: 总结展望
Two sections: 创新点; 工程应用前景
Innovation points: 模块化多工况装置设计; 低温稳定控制与高精度测量; 多因素系统研究与量化分析; 不确定度评估与可靠性提升
Engineering applications: 极地船舶设计优化; 表面工程与减阻开发; 极地装备与结构安全
Closing sentence: 未来将拓展更多材料与工况，建立摩擦数据库，服务极地工程与安全航行。
Visual: apparatus images and polar ship/ice scenes in right-side image stack, structured summary cards on left, footer page number 10.""",
    },
]

for idx, slide in enumerate(slides, 1):
    prompt = common + f"""
Slide-specific task:
Recreate page {slide['page']} as one single slide.
Page number: {slide['page']}
{slide['body']}
Output constraints: one single 16:9 slide only, no surrounding margins, no collage, no thumbnail grid, no extra slide pages.
"""
    path = OUT / f"page_{idx:02d}_prompt.txt"
    path.write_text(prompt, encoding="utf-8")
    print(path)
