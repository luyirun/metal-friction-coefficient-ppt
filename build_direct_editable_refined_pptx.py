from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
SRC_DIR = ROOT / "outputs" / "single_pages"
ASSET_DIR = ROOT / "outputs" / "direct_editable_refined_assets"
OUT = ROOT / "outputs" / "metal_ice_direct_editable_refined.pptx"

W, H = 16, 9
FONT = "Microsoft YaHei"
TEAL = RGBColor(0, 74, 87)
DARK = RGBColor(0, 66, 78)
TEXT = RGBColor(30, 45, 50)
MUTED = RGBColor(86, 108, 116)
LINE = RGBColor(201, 219, 224)
PAPER = RGBColor(247, 250, 250)
CARD = RGBColor(255, 255, 255)
ICE = RGBColor(224, 241, 246)
ORANGE = RGBColor(232, 126, 45)
GREEN = RGBColor(62, 137, 82)
WHITE = RGBColor(255, 255, 255)


def crop(page, name, box):
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    im = Image.open(SRC_DIR / f"page_{page:02d}_industrial_whitepaper.png").convert("RGB")
    path = ASSET_DIR / f"p{page:02d}_{name}.png"
    im.crop(box).save(path, quality=95)
    return path


def prep_assets():
    return {
        "mountains": crop(1, "mountains_clean", (0, 0, 2048, 470)),
        "apparatus": crop(1, "apparatus", (1030, 40, 2048, 1100)),
        "inset": crop(1, "inset", (700, 690, 1050, 1025)),
        "ship": crop(2, "ship", (55, 178, 1135, 680)),
        "ice": crop(2, "ice", (1190, 720, 1998, 1040)),
        "inclined": crop(5, "inclined", (120, 205, 1100, 855)),
        "vertical": crop(5, "vertical", (1180, 220, 2048, 895)),
        "ship1": crop(10, "ship1", (910, 240, 1425, 450)),
        "iceblock": crop(10, "iceblock", (910, 470, 1425, 680)),
        "ship2": crop(10, "ship2", (910, 700, 1425, 910)),
    }


def set_text_style(run, size, color=TEXT, bold=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def text(slide, x, y, w, h, value, size, color=TEXT, bold=False, align=PP_ALIGN.LEFT, name=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.clear()
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    lines = value.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        set_text_style(r, size, color, bold if i == 0 else False)
    return box


def shape(slide, typ, x, y, w, h, fill, line=None, radius=False, trans=0):
    shp_type = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else typ
    shp = slide.shapes.add_shape(shp_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.fill.transparency = trans
    if line:
        shp.line.color.rgb = line
    else:
        shp.line.fill.background()
    return shp


def rect(slide, x, y, w, h, fill, line=None, radius=False, trans=0):
    return shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, h, fill, line, radius, trans)


def line(slide, x1, y1, x2, y2, color=LINE, width=1):
    l = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    l.line.color.rgb = color
    l.line.width = Pt(width)
    return l


def para_tab(slide, x, y, w, h, fill=DARK):
    shp = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = fill
    return shp


def base(slide, num, title, subtitle=None):
    rect(slide, 0, 0, W, H, PAPER)
    para_tab(slide, -0.08, 0.02, 1.38, 0.70)
    text(slide, 0.13, 0.13, 0.70, 0.33, f"{num:02d}", 24, WHITE, True, name=f"S{num:02d}_页码")
    text(slide, 1.32, 0.25, 7.5, 0.43, title, 25, TEAL, True, name=f"S{num:02d}_标题")
    line(slide, 1.32, 0.85, 1.95, 0.85, TEAL, 3)
    line(slide, 2.02, 0.85, 3.12, 0.85, LINE, 1.8)
    if subtitle:
        text(slide, 1.32, 0.98, 6.5, 0.26, subtitle, 12.5, TEAL, False, name=f"S{num:02d}_副标题")
    para_tab(slide, 15.12, 8.47, 0.82, 0.38)
    text(slide, 15.35, 8.55, 0.35, 0.16, f"{num:02d}", 11, WHITE, True, PP_ALIGN.CENTER, name=f"S{num:02d}_页脚")


def card(slide, x, y, w, h, title=None, body=None, name=""):
    rect(slide, x, y, w, h, CARD, RGBColor(224, 233, 236), radius=True)
    if title:
        text(slide, x + 0.22, y + 0.18, w - 0.44, 0.28, title, 13.5, TEAL, True, name=f"{name}_标题" if name else None)
    if body:
        text(slide, x + 0.22, y + 0.58, w - 0.44, h - 0.65, body, 10.8, MUTED, False, name=f"{name}_正文" if name else None)


def bullet(slide, x, y, value, size=11.5):
    rect(slide, x, y + 0.10, 0.055, 0.055, TEAL, radius=True)
    text(slide, x + 0.17, y, 3.6, 0.23, value, size)


def icon(slide, cx, cy, glyph, label=None, sub=None, name=None):
    shp = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(cx - 0.42), Inches(cy - 0.42), Inches(0.84), Inches(0.84))
    shp.fill.solid()
    shp.fill.fore_color.rgb = ICE
    shp.line.color.rgb = RGBColor(185, 216, 224)
    text(slide, cx - 0.25, cy - 0.15, 0.5, 0.25, glyph, 15, TEAL, True, PP_ALIGN.CENTER, name=name)
    if label:
        text(slide, cx - 0.55, cy + 0.55, 1.1, 0.22, label, 12.5, TEAL, True, PP_ALIGN.CENTER)
    if sub:
        text(slide, cx - 0.68, cy + 0.92, 1.36, 0.84, sub, 9.7, TEXT, False, PP_ALIGN.CENTER)


def slide1(slide, a):
    rect(slide, 0, 0, W, H, RGBColor(246, 250, 251))
    slide.shapes.add_picture(str(a["mountains"]), Inches(0), Inches(0), width=Inches(16), height=Inches(3.70))
    rect(slide, 0, 0, W, H, WHITE, trans=52000)
    slide.shapes.add_picture(str(a["apparatus"]), Inches(8.00), Inches(0.22), width=Inches(7.92), height=Inches(8.12))
    para_tab(slide, -0.20, 0.00, 2.50, 1.30)
    text(slide, 0.48, 0.26, 1.05, 0.62, "01", 46, WHITE, True, name="S01_页码")
    text(slide, 0.98, 2.03, 6.8, 1.65, "金属与冰\n摩擦系数的测量", 42, TEAL, True, name="S01_主标题")
    text(slide, 1.02, 4.18, 5.0, 0.36, "青少年科学素养大赛 · 实验答辩", 20, TEAL, True, name="S01_副标题")
    line(slide, 1.02, 4.70, 5.15, 4.70, RGBColor(218, 229, 232), 1.8)
    line(slide, 1.02, 4.70, 1.58, 4.70, TEAL, 3)
    icon(slide, 1.08, 5.92, "人")
    text(slide, 1.62, 5.78, 3.8, 0.32, "参赛选手：实验项目组", 16, TEXT, False, name="S01_参赛选手")
    icon(slide, 1.08, 6.86, "日")
    text(slide, 1.62, 6.73, 3.8, 0.32, "日期：2024.05.20", 16, TEXT, False, name="S01_日期")
    slide.shapes.add_picture(str(a["inset"]), Inches(5.55), Inches(5.88), width=Inches(2.20), height=Inches(2.20))
    para_tab(slide, 0.68, 8.12, 6.50, 0.52)
    text(slide, 0.95, 8.24, 5.7, 0.25, "实验测量 · 误差分析 · 答辩汇报", 12.5, WHITE, True, name="S01_底部标签")
    para_tab(slide, 15.02, 8.38, 0.96, 0.52)
    text(slide, 15.30, 8.52, 0.38, 0.2, "01", 16, WHITE, True, PP_ALIGN.CENTER, name="S01_页脚")


def slide2(slide, a):
    base(slide, 2, "选题背景")
    slide.shapes.add_picture(str(a["ship"]), Inches(0.70), Inches(1.38), width=Inches(7.72), height=Inches(3.60))
    card(slide, 0.70, 5.40, 7.72, 2.10, "研究意义", "揭示船舶破冰与航行效率的关键因素之一\n为船舶材料选型、表面处理与结构设计提供依据\n推动极地装备与工程安全的进步", "S02_研究意义")
    text(slide, 9.65, 1.55, 4.8, 0.36, "北极航运与冰阻力问题", 20, TEAL, True, name="S02_右侧标题")
    text(slide, 9.65, 2.40, 4.3, 0.25, "航行阻力中，冰障碍阻力占比", 12.8, TEXT, name="S02_数字说明1")
    text(slide, 9.65, 2.75, 2.8, 0.52, "25%-55%", 25, ORANGE, True, name="S02_关键数字1")
    line(slide, 9.65, 3.58, 14.86, 3.58, RGBColor(220, 234, 237), 1)
    text(slide, 9.65, 4.05, 4.8, 0.25, "冰面相关事故中，摩擦失控问题占比", 12.8, TEXT, name="S02_数字说明2")
    text(slide, 9.65, 4.40, 2.8, 0.52, "13%-15%", 25, ORANGE, True, name="S02_关键数字2")
    slide.shapes.add_picture(str(a["ice"]), Inches(9.65), Inches(5.86), width=Inches(5.34), height=Inches(1.92))


def slide3(slide, a):
    base(slide, 3, "题目解读")
    data = [
        ("机理分析", "冰的黏附性变弱\n表面微观水膜\n水膜润滑作用\n温度与压力影响", "冰"),
        ("装置搭建", "斜面摩擦实验平台\n力学测量与采集\n低温环境控制\n稳定与校准", "器"),
        ("结果与不确定度", "摩擦系数计算\n不确定度评估\n数据性分析\n结果讨论与验证", "图"),
    ]
    for i, (title, body, glyph) in enumerate(data):
        x = 0.95 + i * 4.75
        card(slide, x, 1.28, 4.35, 6.25, None, None)
        text(slide, x + 0.55, 1.58, 3.25, 0.30, title, 15.8, TEAL, True, PP_ALIGN.CENTER, f"S03_{title}_标题")
        icon(slide, x + 2.18, 2.78, glyph)
        y = 4.02
        for line_text in body.split("\n"):
            bullet(slide, x + 0.55, y, line_text, 11.8)
            y += 0.48
        if i < 2:
            line(slide, x + 4.52, 1.58, x + 4.52, 7.25, LINE, 1)


def slide4(slide, a):
    base(slide, 4, "方法选择 —— 动态斜面法")
    text(slide, 1.18, 1.28, 2.2, 0.30, "力学模型", 15, TEAL, True, name="S04_力学模型标题")
    line(slide, 1.75, 6.40, 7.15, 3.85, RGBColor(170, 180, 184), 2.2)
    rect(slide, 3.95, 4.60, 1.40, 0.72, RGBColor(214, 219, 220), RGBColor(160, 168, 170))
    line(slide, 4.65, 4.96, 4.65, 2.28, TEAL, 1.8)
    line(slide, 4.65, 4.96, 4.65, 6.92, TEAL, 1.8)
    line(slide, 4.65, 4.96, 6.35, 4.20, ORANGE, 1.6)
    text(slide, 4.82, 2.15, 0.35, 0.20, "N", 12.5, TEAL, True, name="S04_N")
    text(slide, 4.80, 6.70, 0.60, 0.22, "mg", 11.5, TEXT, name="S04_mg")
    text(slide, 6.40, 4.05, 0.40, 0.22, "f", 12, ORANGE, True, name="S04_f")
    text(slide, 1.25, 7.02, 5.5, 0.28, "斜面方向建立动力学方程，反推出动摩擦系数", 11.5, MUTED, name="S04_说明")
    text(slide, 9.05, 1.28, 2.0, 0.30, "理论公式", 15, TEAL, True, name="S04_理论公式标题")
    rect(slide, 9.08, 1.88, 4.95, 0.76, RGBColor(241, 245, 245), None, True)
    text(slide, 9.48, 2.08, 4.10, 0.28, "a = g(sinθ - μcosθ)", 17, TEXT, False, PP_ALIGN.CENTER, "S04_公式1")
    text(slide, 11.35, 2.90, 0.35, 0.28, "↓", 22, RGBColor(145, 165, 172), True, PP_ALIGN.CENTER)
    rect(slide, 9.08, 3.38, 4.95, 0.92, RGBColor(241, 245, 245), None, True)
    text(slide, 9.40, 3.62, 4.30, 0.32, "μ = (sinθ - a/g) / cosθ", 17, TEXT, False, PP_ALIGN.CENTER, "S04_公式2")
    text(slide, 9.20, 4.82, 4.1, 1.35, "a —— 沿斜面方向的加速度\ng —— 重力加速度\nθ —— 斜面倾角\nμ —— 动摩擦系数", 12, TEXT, name="S04_符号说明")


def slide5(slide, a):
    base(slide, 5, "装置设计")
    text(slide, 1.16, 1.14, 2.4, 0.28, "斜面冰面模块", 14.2, TEAL, True, name="S05_左模块标题")
    slide.shapes.add_picture(str(a["inclined"]), Inches(0.92), Inches(1.48), width=Inches(7.60), height=Inches(5.10))
    text(slide, 8.78, 1.14, 4.0, 0.28, "垂直冰面模块（压力加载）", 14.2, TEAL, True, name="S05_右模块标题")
    slide.shapes.add_picture(str(a["vertical"]), Inches(8.72), Inches(1.48), width=Inches(5.95), height=Inches(5.10))
    for i, label in enumerate(["模块化设计", "快速更换", "低温环境模拟", "高精度测量"]):
        icon(slide, 1.95 + i * 3.45, 7.38, "◎", label, name=f"S05_标签{i+1}")


def slide6(slide, a):
    base(slide, 6, "变量设计", "六大影响变量与水平设置")
    items = [
        ("温度", "-5 ℃\n-10 ℃\n-15 ℃", "❄"),
        ("压强", "0.05 MPa\n0.10 MPa\n0.20 MPa", "压"),
        ("材质", "铝合金\n不锈钢\n钛合金", "层"),
        ("粗糙度", "Ra 0.2 μm\nRa 1.0 μm\nRa 3.0 μm", "Ra"),
        ("面积", "1 cm²\n4 cm²\n9 cm²", "□"),
        ("成分", "淡水冰\n海冰\n人工掺盐冰", "分"),
    ]
    for i, (title, body, glyph) in enumerate(items):
        x = 1.18 + i * 2.42
        icon(slide, x, 2.88, glyph)
        text(slide, x - 0.62, 3.72, 1.24, 0.25, title, 14.5, TEAL, True, PP_ALIGN.CENTER, f"S06_{title}_标题")
        text(slide, x - 0.75, 4.25, 1.50, 1.05, body, 10.8, TEXT, False, PP_ALIGN.CENTER, f"S06_{title}_水平")
    for i in range(5):
        line(slide, 1.73 + i * 2.42, 3.00, 2.55 + i * 2.42, 3.00, RGBColor(200, 219, 224), 1)
    line(slide, 0.78, 8.18, 15.3, 8.18, RGBColor(215, 231, 235), 0.8)


def slide7(slide, a):
    base(slide, 7, "问题改进")
    rows = [
        ("冰面制备不平整", "表面粗糙和沉积过大\n影响测量重复性", "定制制冰机 + 控温体系\n实现光滑透明平整冰面"),
        ("低温环境稳定性差", "温度波动导致数据漂移\n影响实验结果", "双层保温 + PID温控\n温度波动 ≤ ±0.2 ℃"),
        ("传感器可靠性不足", "低温漂移与噪声干扰大\n影响信号准确性", "低温型传感器 + 屏蔽布线\n多点校准与滤波处理"),
    ]
    glyphs = ["冰", "温", "信"]
    for i, (prob, desc, sol) in enumerate(rows):
        y = 1.48 + i * 1.88
        card(slide, 0.92, y, 5.45, 1.18, None, None)
        icon(slide, 1.70, y + 0.58, glyphs[i])
        text(slide, 2.30, y + 0.18, 2.60, 0.24, prob, 12.5, TEAL, True, name=f"S07_问题{i+1}")
        text(slide, 2.30, y + 0.50, 2.95, 0.45, desc, 9.7, MUTED, name=f"S07_问题说明{i+1}")
        text(slide, 6.73, y + 0.38, 0.40, 0.28, "→", 22, RGBColor(150, 174, 181), True, PP_ALIGN.CENTER)
        card(slide, 7.38, y, 6.00, 1.18, None, None)
        text(slide, 7.75, y + 0.16, 1.25, 0.24, "改进方案", 12.5, GREEN, True, name=f"S07_方案标题{i+1}")
        text(slide, 7.75, y + 0.50, 3.70, 0.45, sol, 9.8, TEXT, name=f"S07_方案说明{i+1}")
        text(slide, 12.85, y + 0.38, 0.35, 0.28, "✓", 23, GREEN, True, PP_ALIGN.CENTER)


def slide8(slide, a):
    base(slide, 8, "数据处理")
    steps = [
        ("1", "数据采集", "力信号\n位移信号\n温度/压力\n采样频率 100 Hz"),
        ("2", "加速度计算", "位移二次微分\n滤波处理\n得到 a(t)"),
        ("3", "摩擦系数计算", "μ = (sinθ - a/g) / cosθ\n代入角度与加速度\n计算 μ"),
        ("4", "不确定度评估", "A类不确定度\nB类不确定度\n合成标准不确定度\n扩展不确定度 U"),
    ]
    for i, (num, title, body) in enumerate(steps):
        x = 0.92 + i * 3.75
        card(slide, x, 1.45, 2.92, 5.72, None, None)
        shp = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x + 1.28), Inches(1.62), Inches(0.36), Inches(0.36))
        shp.fill.solid(); shp.fill.fore_color.rgb = DARK; shp.line.color.rgb = DARK
        text(slide, x + 1.38, 1.68, 0.16, 0.12, num, 8.5, WHITE, True, PP_ALIGN.CENTER, f"S08_步骤{num}")
        text(slide, x + 0.48, 2.02, 1.96, 0.25, title, 12.5, TEAL, True, PP_ALIGN.CENTER, f"S08_{title}_标题")
        if i == 2:
            text(slide, x + 0.38, 3.50, 2.15, 0.32, "μ = (sinθ - a/g) / cosθ", 12.8, TEXT, False, PP_ALIGN.CENTER, "S08_公式")
            text(slide, x + 0.48, 5.10, 1.9, 0.76, "代入角度与加速度\n计算 μ", 9.6, TEXT, name="S08_摩擦系数正文")
        else:
            text(slide, x + 0.48, 5.05, 1.90, 1.0, body, 9.6, TEXT, name=f"S08_{title}_正文")
        if i < 3:
            text(slide, x + 3.05, 4.10, 0.40, 0.25, "→", 22, RGBColor(150, 174, 181), True, PP_ALIGN.CENTER)


def slide9(slide, a):
    base(slide, 9, "结果分析")
    text(slide, 2.80, 1.26, 4.50, 0.30, "摩擦系数对比趋势（示例）", 14.5, TEAL, True, PP_ALIGN.CENTER, "S09_图表标题")
    x0, y0, cw, ch = 1.25, 1.92, 7.00, 4.92
    line(slide, x0, y0 + ch, x0 + cw, y0 + ch, RGBColor(75, 85, 90), 1)
    line(slide, x0, y0, x0, y0 + ch, RGBColor(75, 85, 90), 1)
    for k in range(5):
        y = y0 + k * ch / 4
        line(slide, x0, y, x0 + cw, y, RGBColor(229, 236, 238), 0.6)
    series = [
        (RGBColor(39, 132, 197), [0.32, 0.29, 0.25, 0.21, 0.16], "铝合金"),
        (ORANGE, [0.78, 0.71, 0.64, 0.53, 0.42], "不锈钢"),
        (RGBColor(126, 170, 45), [0.56, 0.49, 0.43, 0.34, 0.25], "钛合金"),
    ]
    for col, vals, name in series:
        pts = [(x0 + 0.55 + j * (cw - 1.1) / 4, y0 + ch - vals[j] * ch) for j in range(5)]
        for p1, p2 in zip(pts, pts[1:]):
            line(slide, p1[0], p1[1], p2[0], p2[1], col, 1.5)
        for px, py in pts:
            dot = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(px - 0.035), Inches(py - 0.035), Inches(0.07), Inches(0.07))
            dot.fill.solid(); dot.fill.fore_color.rgb = col; dot.line.color.rgb = col
    text(slide, 3.65, 7.03, 1.8, 0.22, "温度 / ℃", 11, TEXT, False, PP_ALIGN.CENTER, "S09_X轴")
    text(slide, 0.72, 3.72, 0.34, 1.2, "摩擦系数 μ", 11, TEXT, name="S09_Y轴")
    text(slide, 9.55, 1.55, 1.80, 0.30, "数据解读", 14.5, TEAL, True, name="S09_数据解读标题")
    for i, value in enumerate(["温度升高，μ整体下降", "粗糙度增大，μ上升", "压强增大，先升后降", "材质差异显著，钛合金综合表现最优"]):
        bullet(slide, 9.55, 2.18 + i * 0.55, value, 11.5)
    card(slide, 9.55, 5.62, 3.80, 1.02, None, None)
    text(slide, 10.05, 5.92, 2.85, 0.42, "结果与文献趋势一致，\n验证方法的有效性", 11.2, TEAL, True, PP_ALIGN.CENTER, "S09_结论")


def slide10(slide, a):
    base(slide, 10, "总结展望")
    text(slide, 3.05, 1.06, 1.30, 0.25, "创新点", 12.8, TEAL, True, name="S10_创新点标题")
    points = ["模块化多工况装置设计", "低温稳定控制与高精度测量", "多因素系统研究与量化分析", "不确定度评估与可靠性提升"]
    for i, p in enumerate(points):
        card(slide, 1.15, 1.45 + i * 1.18, 5.65, 0.72, None, None)
        icon(slide, 1.60, 1.80 + i * 1.18, ["器", "温", "图", "盾"][i])
        text(slide, 2.18, 1.63 + i * 1.18, 3.8, 0.24, p, 11.5, TEXT, name=f"S10_创新点{i+1}")
    text(slide, 10.85, 1.06, 2.10, 0.25, "工程应用前景", 12.8, TEAL, True, name="S10_应用标题")
    imgs = [a["ship1"], a["iceblock"], a["ship2"]]
    labels = [("极地船舶设计优化", "降低航行阻力"), ("表面工程与减阻开发", "提升抗冰性能"), ("极地装备与结构安全", "提供理论依据")]
    for i, (head, body) in enumerate(labels):
        y = 1.45 + i * 1.48
        slide.shapes.add_picture(str(imgs[i]), Inches(8.30), Inches(y), width=Inches(3.05), height=Inches(1.05))
        card(slide, 11.55, y, 3.35, 1.05, None, None)
        text(slide, 11.80, y + 0.18, 2.5, 0.22, head, 11.2, TEAL, True, name=f"S10_应用{i+1}_标题")
        text(slide, 11.80, y + 0.50, 2.0, 0.20, body, 9.6, MUTED, name=f"S10_应用{i+1}_正文")
    card(slide, 1.15, 7.25, 13.95, 0.72, None, None)
    text(slide, 2.00, 7.47, 11.20, 0.22, "未来将拓展更多材料与工况，建立摩擦数据库，服务极地工程与安全航行。", 10.8, TEXT, name="S10_展望")


def build():
    a = prep_assets()
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    builders = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10]
    for builder in builders:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        builder(slide, a)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
