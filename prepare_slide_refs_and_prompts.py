from pathlib import Path
from PIL import Image

ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
OUT = ROOT / "outputs" / "keynote_single_pages"
REF = OUT / "refs"
PROMPTS = OUT / "prompts"
OUT.mkdir(parents=True, exist_ok=True)
REF.mkdir(parents=True, exist_ok=True)
PROMPTS.mkdir(parents=True, exist_ok=True)

collage = ROOT / "outputs" / "imagegen_ppt_collage_04_premium_polar_keynote_corrected.png"
im = Image.open(collage).convert("RGB")
w, h = im.size

slides = [
    {
        "title": "金属与冰摩擦系数的测量",
        "point": "可重复、可调控的测量系统",
        "body": "围绕金属-冰接触面的滑动摩擦特性，搭建实验装置，分析不同条件下摩擦系数的变化规律。",
        "visual": "金属块在透明冰质斜面装置上滑动的高清主视觉，极地冰山背景。",
    },
    {
        "title": "选题背景",
        "point": "冰上摩擦影响极地航运阻力与能耗",
        "body": "北极航道可缩短航程 25%-55%，但冰阻力会使燃油消耗增加约 13%-15%。金属与冰摩擦系数测量可为船冰阻力分析和极地装备设计提供依据。",
        "visual": "破冰船在海冰航道中航行，突出 25%-55% 和 13%-15% 两个关键数字。",
    },
    {
        "title": "题目解读",
        "point": "完成机制分析、装置搭建、结果评估",
        "body": "研究任务包括分析金属与冰之间的滑动摩擦机制，设计并制作实验装置，获得摩擦系数测量结果并讨论不确定度。",
        "visual": "三步流程图：机制分析、装置搭建、结果评估，使用圆形图标和箭头连接。",
    },
    {
        "title": "方法选择",
        "point": "动态斜面法反推出动摩擦系数",
        "body": "通过记录物块沿冰面下滑过程中的运动参数，结合斜面方向动力学方程 ma = mg sinθ - μmg cosθ，反推出动摩擦系数 μ。",
        "visual": "斜面受力分析图，包含 N、f、mg sinθ、θ、金属块和冰面。",
    },
    {
        "title": "装置设计",
        "point": "斜面冰面与竖直冰面双结构验证",
        "body": "实验装置由斜面冰面结构和竖直冰面结构组成，用于动态滑动测量与补充验证，提高实验结果可靠性。",
        "visual": "铝型材实验台、透明冰槽、金属滑块、竖直冰面模块的高质感装置渲染。",
    },
    {
        "title": "变量设计",
        "point": "冰体维度与金属维度分变量控制",
        "body": "围绕温度、压强、材质、粗糙度、接触面积和冰的成分建立因素矩阵，通过单因素控制比较摩擦系数变化。",
        "visual": "以 μ 为中心的六因素矩阵图，冰蓝与深蓝图标系统。",
    },
    {
        "title": "问题改进",
        "point": "解决冰面平整、低温保持与传感稳定",
        "body": "针对冰面气泡和凹坑、实验温度波动、低温传感器稳定性等问题，采用去离子水、平整制冰、低温控制和外置传感器等改进方案。",
        "visual": "三张并列图片卡片：平整冰块、低温显示、传感器/接头细节。",
    },
    {
        "title": "数据处理",
        "point": "由 x、t、θ 得到 a，再计算 μ",
        "body": "实验记录位移、时间、倾角等参数，计算加速度后代入动力学模型得到摩擦系数，并进行不确定度评估。",
        "visual": "四步流程卡片：数据采集 x,t,θ → 加速度 a → 摩擦系数 μ → 不确定度 U(μ)。",
    },
    {
        "title": "结果分析",
        "point": "比较不同条件下摩擦系数变化趋势",
        "body": "通过 μ-温度、μ-压力、μ-粗糙度等趋势图分析变量影响，不凭空给出精确数据，只展示对比分析框架。",
        "visual": "三个并列趋势图，标注 μ-温度关系、μ-压强关系、μ-粗糙度关系。",
    },
    {
        "title": "总结展望",
        "point": "自制装置、多变量测量、工程应用",
        "body": "项目形成从实验装置、变量控制、数据处理到工程应用的完整链条，后续可提升低温控制精度并扩展更多材料与表面处理方式。",
        "visual": "装置渲染与极地应用场景组合，包含破冰船和冰面装备意象。",
    },
]

base = """Use case: productivity-visual
Asset type: independent high-definition 16:9 single PowerPoint slide visual mockup
Primary request: Expand the provided reference thumbnail from the selected "premium polar engineering keynote" collage into one complete high-resolution 16:9 widescreen PPT slide. This is NOT a crop and NOT a redesign; use the reference image as the visual master for layout, style, content structure, page number, color system, imagery direction, and spacing, but redraw it as a clean standalone slide with sharper details.
Input image role: visual reference for this exact slide only. Preserve its layout relationship and visual system.
Style/medium: high-end polar engineering keynote, refined scientific presentation, soft icy atmosphere, realistic metal and ice materials, premium Chinese competition deck.
Color palette: glacier white, frosted blue, deep petroleum teal, silver metal, small orange accent for key numbers.
Typography: bold modern Chinese sans-serif titles, concise readable labels, clean footer and page number.
Visual system: consistent with the selected collage: polar ice background, technical diagrams, realistic apparatus renders, subtle footer bar, page number, refined charts and icons.
Slide number: {num:02d}
Title text: {title}
Core point: {point}
Body content to use if small reference text is unclear: {body}
Main visual direction: {visual}
Constraints: output exactly one full 16:9 landscape slide, no collage, no multiple slides, no cropping frame, no watermark. Main title, core point, key numbers/formulas/chart labels and page number must be clear. Do not invent new project claims or precise data. If any small text from the reference is unreadable, use the provided title/core point/body content.
Strict text rule: do not add author names, dates, project IDs, competition names, school names, aviation/aerospace labels, or any metadata not listed here. Footer text should be minimal, such as "金属与冰摩擦系数测量" and the page number "{num:02d}/10" only.
Avoid: changing to dark dashboard style, adding unrelated content, extra pages, wrong page number, tiny unreadable main title, chaotic layout, cartoon style.
"""

for i in range(10):
    col = i % 2
    row = i // 2
    left = round(col * w / 2)
    right = round((col + 1) * w / 2)
    top = round(row * h / 5)
    bottom = round((row + 1) * h / 5)
    crop = im.crop((left, top, right, bottom))
    crop_path = REF / f"slide_{i+1:02d}_ref.png"
    crop.save(crop_path)
    prompt = base.format(num=i + 1, **slides[i])
    (PROMPTS / f"slide_{i+1:02d}_prompt.txt").write_text(prompt, encoding="utf-8")
    print(crop_path)
