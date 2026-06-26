from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "outputs"
SRC_DIR = OUT_DIR / "single_pages"
ASSET_DIR = OUT_DIR / "whitepaper_native_assets"
OUT = OUT_DIR / "metal_ice_whitepaper_native_editable.pptx"

W, H = 16, 9
FONT = "Microsoft YaHei"

PAPER = RGBColor(248, 250, 250)
WHITE = RGBColor(255, 255, 255)
TEAL = RGBColor(0, 72, 84)
TEAL_2 = RGBColor(0, 96, 112)
TEAL_LIGHT = RGBColor(226, 243, 247)
DARK = RGBColor(28, 44, 49)
MUTED = RGBColor(88, 108, 114)
LINE = RGBColor(211, 225, 229)
ORANGE = RGBColor(226, 126, 39)
GREEN = RGBColor(54, 134, 76)
GRAY_FILL = RGBColor(243, 246, 246)


def crop(page, name, box):
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    src = Image.open(SRC_DIR / f"page_{page:02d}_industrial_whitepaper.png").convert("RGB")
    path = ASSET_DIR / f"p{page:02d}_{name}.png"
    src.crop(box).save(path, quality=95)
    return path


def prep_assets():
    return {
        "hero": crop(1, "hero_apparatus", (880, 0, 2048, 1030)),
        "inset": crop(1, "inset", (710, 690, 1050, 1015)),
        "ship": crop(2, "ship", (55, 175, 1130, 640)),
        "ice": crop(2, "ice", (1180, 700, 1998, 1015)),
        "inclined": crop(5, "inclined_clean", (360, 235, 1070, 830)),
        "vertical": crop(5, "vertical_clean", (1220, 245, 1700, 825)),
        "ship_small": crop(10, "ship_small", (905, 235, 1415, 455)),
        "ice_block": crop(10, "ice_block", (905, 470, 1415, 685)),
        "ship_small_2": crop(10, "ship_small_2", (905, 690, 1415, 910)),
    }


def rect(slide, x, y, w, h, fill=WHITE, line=None, radius=False, transparency=0):
    shp = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.fill.transparency = transparency
    if line:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    return shp


def line(slide, x1, y1, x2, y2, color=TEAL, width=1.4):
    shp = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    shp.line.color.rgb = color
    shp.line.width = Pt(width)
    return shp


def text(slide, x, y, w, h, value, size, color=DARK, bold=False, align=PP_ALIGN.LEFT, name=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.clear()
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    for i, raw in enumerate(value.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = raw
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return box


def para_tab(slide, x, y, w, h, fill=TEAL):
    shp = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = fill
    return shp


def bullet(slide, x, y, value, size=12.5, color=DARK):
    rect(slide, x, y + 0.11, 0.055, 0.055, TEAL, radius=True)
    text(slide, x + 0.17, y, 4.2, 0.24, value, size, color)


def page_base(slide, page, title, subtitle=None):
    rect(slide, 0, 0, W, H, PAPER)
    para_tab(slide, -0.10, 0.08, 1.35, 0.66)
    text(slide, 0.18, 0.20, 0.58, 0.32, f"{page:02d}", 22, WHITE, True, name=f"S{page:02d}_page")
    text(slide, 1.42, 0.22, 7.5, 0.36, title, 23.5, TEAL, True, name=f"S{page:02d}_title")
    line(slide, 0.72, 0.86, 1.28, 0.86, TEAL, 3)
    line(slide, 1.40, 0.86, 2.62, 0.86, TEAL_2, 1.4)
    if subtitle:
        text(slide, 1.42, 0.98, 7.0, 0.24, subtitle, 13, MUTED, name=f"S{page:02d}_subtitle")
    para_tab(slide, 15.12, 8.50, 0.80, 0.36)
    text(slide, 15.36, 8.58, 0.27, 0.12, f"{page:02d}", 10, WHITE, True, PP_ALIGN.CENTER, name=f"S{page:02d}_footer")


def card(slide, x, y, w, h, fill=WHITE, border=LINE, radius=True):
    return rect(slide, x, y, w, h, fill, border, radius)


def circle_icon(slide, cx, cy, label, glyph=None):
    oval = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(cx - 0.42), Inches(cy - 0.42), Inches(0.84), Inches(0.84))
    oval.fill.solid()
    oval.fill.fore_color.rgb = TEAL_LIGHT
    oval.line.color.rgb = RGBColor(183, 216, 224)
    if glyph:
        text(slide, cx - 0.28, cy - 0.17, 0.56, 0.24, glyph, 14, TEAL, True, PP_ALIGN.CENTER)
    text(slide, cx - 0.62, cy + 0.60, 1.24, 0.20, label, 12.5, TEAL, True, PP_ALIGN.CENTER)


def slide01(slide, a):
    rect(slide, 0, 0, W, H, RGBColor(246, 250, 251))
    slide.shapes.add_picture(str(a["hero"]), Inches(6.55), Inches(0.05), width=Inches(9.1), height=Inches(7.95))
    rect(slide, 6.2, 0, 2.0, H, RGBColor(246, 250, 251), transparency=46000)
    para_tab(slide, -0.12, 0.10, 1.50, 0.72)
    text(slide, 0.20, 0.22, 0.62, 0.33, "01", 22, WHITE, True)
    text(slide, 0.72, 1.78, 5.45, 1.38, "金属与冰\n摩擦系数的测量", 34, TEAL, True, name="S01_main_title")
    text(slide, 0.78, 3.58, 3.6, 0.34, "物理实验创新竞赛答辩", 18, DARK, True, name="S01_subtitle")
    line(slide, 0.78, 4.05, 1.38, 4.05, TEAL, 3)
    line(slide, 1.50, 4.05, 2.92, 4.05, LINE, 1.5)
    slide.shapes.add_picture(str(a["inset"]), Inches(4.68), Inches(4.52), width=Inches(1.68), height=Inches(1.68))
    card(slide, 0.78, 4.76, 3.15, 0.58, RGBColor(250, 252, 252), LINE)
    text(slide, 1.18, 4.96, 2.25, 0.18, "团队：祝冰间隙者", 12, DARK, name="S01_team")
    card(slide, 0.78, 5.58, 3.15, 0.58, RGBColor(250, 252, 252), LINE)
    text(slide, 1.18, 5.78, 2.25, 0.18, "日期：2024.05.20", 12, DARK, name="S01_date")
    para_tab(slide, 0.70, 7.95, 4.80, 0.42)
    text(slide, 0.96, 8.06, 4.2, 0.18, "精准测量  ·  理解机理  ·  服务工程", 10.5, WHITE, True)
    para_tab(slide, 15.12, 8.50, 0.80, 0.36)
    text(slide, 15.36, 8.58, 0.27, 0.12, "01", 10, WHITE, True, PP_ALIGN.CENTER)


def slide02(slide, a):
    page_base(slide, 2, "选题背景")
    slide.shapes.add_picture(str(a["ship"]), Inches(0.80), Inches(1.32), width=Inches(7.45), height=Inches(3.45))
    card(slide, 0.80, 5.22, 7.45, 1.78)
    text(slide, 1.18, 5.42, 1.3, 0.22, "研究意义", 13, TEAL, True, name="S02_research_title")
    for i, t in enumerate(["揭示船舶破冰与航行效率的关键因素之一", "为船舶材料选型、表面处理与结构设计提供依据", "推动极地装备与工程安全的进步"]):
        bullet(slide, 1.18, 5.84 + i * 0.36, t, 11)
    text(slide, 9.05, 1.48, 4.8, 0.32, "北极航运与冰阻力问题", 19, TEAL, True, name="S02_right_title")
    text(slide, 9.05, 2.28, 4.6, 0.22, "航行阻力中，冰障碍阻力占比", 12.5, MUTED, name="S02_num_desc1")
    text(slide, 9.05, 2.63, 2.30, 0.42, "25%-55%", 25, ORANGE, True, name="S02_number1")
    line(slide, 9.05, 3.42, 13.75, 3.42, LINE, 1)
    text(slide, 9.05, 3.82, 4.8, 0.22, "冰面相关事故中，摩擦失控问题占比", 12.5, MUTED, name="S02_num_desc2")
    text(slide, 9.05, 4.17, 2.30, 0.42, "13%-15%", 25, ORANGE, True, name="S02_number2")
    slide.shapes.add_picture(str(a["ice"]), Inches(9.05), Inches(5.30), width=Inches(5.70), height=Inches(1.95))


def slide03(slide, a):
    page_base(slide, 3, "题目解读")
    items = [
        ("机理分析", "冰的黏附性变弱\n表面微观水膜\n水膜润滑作用\n温度与压力影响", "□"),
        ("装置搭建", "斜面摩擦实验平台\n力学测量与采集\n低温环境控制\n稳定与校准", "⚒"),
        ("结果与不确定度", "摩擦系数计算\n不确定度评估\n数据性分析\n结果讨论与验证", "↗"),
    ]
    for i, (title, body, glyph) in enumerate(items):
        x = 0.95 + i * 4.75
        card(slide, x, 1.35, 4.35, 5.95)
        text(slide, x + 0.55, 1.72, 3.25, 0.26, title, 16, TEAL, True, PP_ALIGN.CENTER, name=f"S03_col{i+1}_title")
        circle_icon(slide, x + 2.18, 2.95, "", glyph)
        for j, t in enumerate(body.split("\n")):
            bullet(slide, x + 0.72, 4.22 + j * 0.50, t, 13.0)


def slide04(slide, a):
    page_base(slide, 4, "方法选择 —— 动态斜面法")
    text(slide, 1.10, 1.35, 1.6, 0.24, "力学模型", 14.5, TEAL, True, name="S04_model_title")
    line(slide, 1.65, 6.70, 7.45, 3.60, RGBColor(166, 180, 184), 2.2)
    rect(slide, 3.90, 4.70, 1.55, 0.76, RGBColor(218, 224, 225), RGBColor(156, 168, 171))
    line(slide, 4.68, 5.08, 4.68, 2.35, TEAL, 1.8)
    line(slide, 4.68, 5.08, 4.68, 6.92, TEAL, 1.8)
    line(slide, 4.68, 5.08, 6.35, 4.16, ORANGE, 1.7)
    line(slide, 4.68, 5.08, 3.48, 5.72, ORANGE, 1.7)
    text(slide, 4.82, 2.18, 0.35, 0.20, "N", 12, TEAL, True)
    text(slide, 4.86, 6.70, 0.60, 0.22, "mg", 11.5, DARK)
    text(slide, 6.45, 4.02, 0.40, 0.22, "a", 12, ORANGE, True)
    text(slide, 3.20, 5.68, 0.40, 0.22, "f", 12, ORANGE, True)
    text(slide, 1.18, 7.28, 5.8, 0.24, "沿斜面方向建立动力学方程，反推出动摩擦系数", 11.8, MUTED)
    text(slide, 9.05, 1.35, 1.7, 0.24, "理论公式", 14.5, TEAL, True, name="S04_formula_title")
    card(slide, 9.10, 1.92, 4.65, 0.70, GRAY_FILL, None)
    text(slide, 9.46, 2.14, 3.90, 0.22, "a = g(sinθ - μcosθ)", 17, DARK, False, PP_ALIGN.CENTER, name="S04_formula1")
    text(slide, 11.26, 2.88, 0.35, 0.24, "↓", 18, TEAL_2, True, PP_ALIGN.CENTER)
    card(slide, 9.10, 3.30, 4.65, 0.86, GRAY_FILL, None)
    text(slide, 9.43, 3.56, 4.0, 0.24, "μ = (sinθ - a/g) / cosθ", 17, DARK, False, PP_ALIGN.CENTER, name="S04_formula2")
    text(slide, 9.28, 4.70, 4.0, 1.20, "a —— 沿斜面方向的加速度\ng —— 重力加速度\nθ —— 斜面倾角\nμ —— 动摩擦系数", 12.4, DARK, name="S04_symbol_notes")


def slide05(slide, a):
    page_base(slide, 5, "装置设计")
    text(slide, 1.00, 1.18, 2.2, 0.24, "斜面冰面模块", 14, TEAL, True, name="S05_left_title")
    slide.shapes.add_picture(str(a["inclined"]), Inches(1.45), Inches(1.55), width=Inches(6.55), height=Inches(4.85))
    rect(slide, 1.45, 1.55, 2.58, 1.05, PAPER)
    rect(slide, 1.45, 2.45, 0.84, 3.05, PAPER)
    text(slide, 8.52, 1.18, 3.6, 0.24, "垂直冰面模块（压力加载）", 14, TEAL, True, name="S05_right_title")
    slide.shapes.add_picture(str(a["vertical"]), Inches(8.90), Inches(1.60), width=Inches(4.95), height=Inches(4.80))
    rect(slide, 8.90, 1.60, 4.95, 1.05, PAPER)
    rect(slide, 11.75, 1.50, 2.35, 0.60, PAPER)
    rect(slide, 13.05, 1.78, 0.78, 4.05, PAPER)
    labels = ["导轨与滑块", "力传感器", "位移传感器", "冰面板", "角度调节机构", "支撑框架"]
    for i, t in enumerate(labels):
        y = 1.86 + i * 0.48
        text(slide, 0.62, y, 1.42, 0.18, t, 10.6, MUTED, name=f"S05_left_label{i+1}")
        line(slide, 1.82, y + 0.10, 2.60, y + 0.10, LINE, 0.8)
    rlabels = ["压力加载装置", "力传感器", "冰面板", "温控仓体", "底盘"]
    for i, t in enumerate(rlabels):
        y = 1.90 + i * 0.67
        text(slide, 13.48, y, 1.42, 0.18, t, 10.6, MUTED, name=f"S05_right_label{i+1}")
        line(slide, 12.72, y + 0.10, 13.42, y + 0.10, LINE, 0.8)
    for i, (label, glyph) in enumerate([("模块化设计", "◇"), ("快速更换", "×"), ("低温环境模拟", "❄"), ("高精度测量", "◎")]):
        circle_icon(slide, 1.65 + i * 3.45, 7.38, label, glyph)


def slide06(slide, a):
    page_base(slide, 6, "变量设计", "六大影响变量与水平设置")
    items = [
        ("温度", "-5 ℃\n-10 ℃\n-15 ℃", "❄"),
        ("压强", "0.05 MPa\n0.10 MPa\n0.20 MPa", "P"),
        ("材质", "铝合金\n不锈钢\n钛合金", "▰"),
        ("粗糙度", "Ra 0.2 μm\nRa 1.0 μm\nRa 3.0 μm", "≈"),
        ("面积", "1 cm²\n4 cm²\n9 cm²", "□"),
        ("成分", "淡水冰\n海冰\n人工掺盐冰", "⌬"),
    ]
    for i, (title, body, glyph) in enumerate(items):
        x = 1.18 + i * 2.42
        circle_icon(slide, x, 2.75, title, glyph)
        text(slide, x - 0.74, 4.02, 1.48, 1.08, body, 12.0, DARK, False, PP_ALIGN.CENTER, name=f"S06_{title}")
    for i in range(5):
        line(slide, 1.75 + i * 2.42, 2.75, 2.34 + i * 2.42, 2.75, LINE, 1)
    line(slide, 0.78, 7.88, 15.20, 7.88, LINE, 1)


def slide07(slide, a):
    page_base(slide, 7, "问题改进")
    rows = [
        ("冰面制备不平整", "表面粗糙和气泡大\n影响测量重复性", "定制制冰机 + 控温体系\n实现光滑透明平整冰面"),
        ("低温环境稳定性差", "温度波动导致数据漂移\n影响实验结果", "双层保温 + PID温控\n温度波动 ≤ ±0.2 ℃"),
        ("传感器可靠性不足", "低温漂移与噪声干扰大\n影响信号准确性", "低温型传感器 + 屏蔽布线\n多点校准与滤波处理"),
    ]
    glyphs = ["□", "℃", "∿"]
    for i, (problem, desc, solution) in enumerate(rows):
        y = 1.45 + i * 1.88
        card(slide, 0.92, y, 5.32, 1.22, RGBColor(247, 250, 250))
        circle_icon(slide, 1.55, y + 0.60, "", glyphs[i])
        text(slide, 2.18, y + 0.16, 2.35, 0.22, problem, 12.8, DARK, True, name=f"S07_problem{i+1}")
        text(slide, 2.18, y + 0.50, 2.80, 0.45, desc, 11.3, MUTED, name=f"S07_problem_desc{i+1}")
        text(slide, 6.60, y + 0.44, 0.55, 0.28, "→", 21, RGBColor(142, 180, 190), True, PP_ALIGN.CENTER)
        card(slide, 7.35, y, 5.92, 1.22, RGBColor(248, 251, 249))
        text(slide, 7.72, y + 0.16, 1.10, 0.22, "改进方案", 12.8, GREEN, True, name=f"S07_solution_title{i+1}")
        text(slide, 7.72, y + 0.50, 3.88, 0.45, solution, 11.3, DARK, name=f"S07_solution{i+1}")
        text(slide, 12.65, y + 0.42, 0.40, 0.28, "✓", 22, GREEN, True, PP_ALIGN.CENTER)


def slide08(slide, a):
    page_base(slide, 8, "数据处理")
    steps = [
        ("1", "数据采集", "力信号\n位移信号\n温度/压力\n采样频率 100 Hz"),
        ("2", "加速度计算", "位移二次微分\n滤波处理\n得到 a(t)"),
        ("3", "摩擦系数计算", "μ = (sinθ - a/g) / cosθ\n代入角度与加速度\n计算 μ"),
        ("4", "不确定度评估", "A类不确定度\nB类不确定度\n合成标准不确定度\n扩展不确定度 U"),
    ]
    for i, (num, title, body) in enumerate(steps):
        x = 0.92 + i * 3.75
        card(slide, x, 1.45, 2.92, 5.86)
        oval = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x + 1.29), Inches(1.62), Inches(0.36), Inches(0.36))
        oval.fill.solid()
        oval.fill.fore_color.rgb = TEAL
        oval.line.color.rgb = TEAL
        text(slide, x + 1.39, 1.68, 0.16, 0.12, num, 8.5, WHITE, True, PP_ALIGN.CENTER, name=f"S08_step{num}")
        text(slide, x + 0.44, 2.12, 2.04, 0.24, title, 12.5, TEAL, True, PP_ALIGN.CENTER, name=f"S08_title{num}")
        if i == 2:
            text(slide, x + 0.35, 3.82, 2.18, 0.55, "μ = (sinθ - a/g)\n      / cosθ", 14.5, DARK, False, PP_ALIGN.CENTER, name="S08_formula")
            text(slide, x + 0.58, 5.28, 1.72, 0.42, "代入角度与加速度\n计算 μ", 10.5, DARK, name="S08_formula_note")
        else:
            for j, t in enumerate(body.split("\n")):
                bullet(slide, x + 0.52, 4.86 + j * 0.34, t, 10.8)
        if i < 3:
            text(slide, x + 3.05, 4.02, 0.42, 0.25, "→", 20, RGBColor(142, 180, 190), True, PP_ALIGN.CENTER)


def slide09(slide, a):
    page_base(slide, 9, "结果分析")
    text(slide, 2.35, 1.25, 4.5, 0.28, "摩擦系数对比趋势（示例）", 14.5, TEAL, True, PP_ALIGN.CENTER, name="S09_chart_title")
    x0, y0, cw, ch = 1.35, 2.00, 6.75, 4.65
    line(slide, x0, y0 + ch, x0 + cw, y0 + ch, DARK, 1)
    line(slide, x0, y0, x0, y0 + ch, DARK, 1)
    for k in range(5):
        y = y0 + k * ch / 4
        line(slide, x0, y, x0 + cw, y, RGBColor(230, 236, 238), 0.5)
        text(slide, x0 - 0.48, y - 0.07, 0.30, 0.12, f"{0.15 - k * 0.03:.2f}", 8.0, DARK, False, PP_ALIGN.RIGHT)
    xs = [-15, -10, -5, 0]
    for j, val in enumerate(xs):
        x = x0 + 0.55 + j * (cw - 1.1) / 3
        line(slide, x, y0 + ch, x, y0 + ch + 0.06, DARK, 0.8)
        text(slide, x - 0.20, y0 + ch + 0.16, 0.40, 0.12, str(val), 8.5, DARK, False, PP_ALIGN.CENTER)
    series = [
        (RGBColor(41, 137, 190), [0.047, 0.040, 0.030, 0.020], "铝合金"),
        (ORANGE, [0.116, 0.105, 0.087, 0.064], "不锈钢"),
        (RGBColor(120, 168, 42), [0.082, 0.070, 0.058, 0.039], "钛合金"),
    ]
    for col, vals, name in series:
        pts = []
        for j, v in enumerate(vals):
            px = x0 + 0.55 + j * (cw - 1.1) / 3
            py = y0 + ch - (v / 0.15) * ch
            pts.append((px, py))
        for p1, p2 in zip(pts, pts[1:]):
            line(slide, p1[0], p1[1], p2[0], p2[1], col, 1.6)
        for px, py in pts:
            dot = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(px - 0.035), Inches(py - 0.035), Inches(0.07), Inches(0.07))
            dot.fill.solid()
            dot.fill.fore_color.rgb = col
            dot.line.color.rgb = col
    text(slide, 3.55, 7.05, 1.8, 0.20, "温度 / ℃", 10.5, DARK, False, PP_ALIGN.CENTER, name="S09_x_axis")
    text(slide, 0.68, 3.55, 0.36, 1.2, "摩擦系数 μ", 10.5, DARK, name="S09_y_axis")
    text(slide, 9.35, 1.45, 1.6, 0.24, "数据解读", 14.5, TEAL, True, name="S09_interpret_title")
    for i, t in enumerate(["温度升高，μ整体下降", "粗糙度增大，μ上升", "压强增大，先升后降", "材质差异显著，钛合金综合表现最优"]):
        bullet(slide, 9.35, 2.02 + i * 0.56, t, 12.2)
    card(slide, 9.35, 5.45, 3.95, 0.96, RGBColor(242, 248, 248))
    text(slide, 9.85, 5.76, 2.95, 0.34, "结果与文献趋势一致，\n验证方法的有效性", 11.5, TEAL, True, PP_ALIGN.CENTER, name="S09_conclusion")


def slide10(slide, a):
    page_base(slide, 10, "总结展望")
    text(slide, 2.25, 1.12, 1.0, 0.20, "创新点", 13, TEAL, True, name="S10_innovation_title")
    points = [
        "模块化多工况装置设计",
        "低温稳定控制与高精度测量",
        "多因素系统研究与量化分析",
        "不确定度评估与可靠性提升",
    ]
    for i, p in enumerate(points):
        card(slide, 1.10, 1.48 + i * 1.05, 5.55, 0.72)
        text(slide, 1.45, 1.67 + i * 1.05, 4.20, 0.22, p, 11.8, DARK, name=f"S10_point{i+1}")
    text(slide, 10.70, 1.12, 2.1, 0.20, "工程应用前景", 13, TEAL, True, name="S10_app_title")
    imgs = [a["ship_small"], a["ice_block"], a["ship_small_2"]]
    apps = [("极地船舶设计优化", "降低航行阻力"), ("表面工程与减阻开发", "提升抗冰性能"), ("极地装备与结构安全", "提供理论依据")]
    for i, (title, body) in enumerate(apps):
        y = 1.50 + i * 1.43
        slide.shapes.add_picture(str(imgs[i]), Inches(8.00), Inches(y), width=Inches(3.0), height=Inches(1.08))
        card(slide, 11.15, y, 3.45, 1.08)
        text(slide, 11.42, y + 0.19, 2.45, 0.20, title, 11.2, TEAL, True, name=f"S10_app{i+1}_title")
        text(slide, 11.42, y + 0.52, 1.8, 0.18, body, 9.5, MUTED, name=f"S10_app{i+1}_body")
    card(slide, 1.10, 7.30, 13.65, 0.54)
    text(slide, 1.65, 7.48, 12.1, 0.18, "未来将拓展更多材料与工况，建立摩擦数据库，服务极地工程与安全航行。", 10.8, DARK, name="S10_future")


def build():
    assets = prep_assets()
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    builders = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10]
    for builder in builders:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        builder(slide, assets)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
