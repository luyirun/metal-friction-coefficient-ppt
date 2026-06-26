from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "outputs"
SRC = OUT_DIR / "keynote_single_pages"
ASSETS = OUT_DIR / "keynote_pptx_assets_v3"
OUT = OUT_DIR / "metal_ice_keynote_editable_v3.pptx"
ASSETS.mkdir(exist_ok=True)

W, H = 13.333333, 7.5
NAVY = RGBColor(5, 54, 99)
BLUE = RGBColor(0, 94, 156)
TEAL = RGBColor(0, 103, 135)
PAPER = RGBColor(246, 251, 254)
ICE = RGBColor(231, 245, 251)
ICE2 = RGBColor(218, 236, 246)
MID = RGBColor(177, 207, 225)
PALE = RGBColor(241, 248, 252)
ORANGE = RGBColor(238, 103, 30)
GRAY = RGBColor(62, 82, 98)
SOFT = RGBColor(124, 151, 166)
WHITE = RGBColor(255, 255, 255)


def source(slide_no: int) -> Image.Image:
    return Image.open(SRC / f"slide_{slide_no:02d}.png").convert("RGB")


def crop_fit(slide_no: int, name: str, box, aspect: float) -> Path:
    """Crop the selected source region to the target aspect without stretching."""
    im = source(slide_no)
    sw, sh = im.size
    x1, y1, x2, y2 = box
    left, top, right, bottom = int(x1 * sw), int(y1 * sh), int(x2 * sw), int(y2 * sh)
    width, height = right - left, bottom - top
    current = width / height
    if current > aspect:
        new_w = int(height * aspect)
        delta = (width - new_w) // 2
        left += delta
        right = left + new_w
    elif current < aspect:
        new_h = int(width / aspect)
        delta = (height - new_h) // 2
        top += delta
        bottom = top + new_h
    out = ASSETS / f"s{slide_no:02d}_{name}.png"
    im.crop((left, top, right, bottom)).save(out, quality=95)
    return out


def force_diagram_clean() -> Path:
    out = crop_fit(4, "force_diagram_clean", (0.40, 0.05, 0.98, 0.88), 7.35 / 5.9)
    im = Image.open(out).convert("RGB")
    draw = Image.new("RGB", im.size, (234, 244, 250))
    pixels = im.load()
    w, h = im.size
    slide_x, slide_y, slide_w, slide_h = 5.42, 0.72, 7.35, 5.9

    def cover(abs_box):
        x, y, bw, bh = abs_box
        left = max(0, int((x - slide_x) / slide_w * w))
        top = max(0, int((y - slide_y) / slide_h * h))
        right = min(w, int((x + bw - slide_x) / slide_w * w))
        bottom = min(h, int((y + bh - slide_y) / slide_h * h))
        sx = min(w - 1, max(0, left - 12))
        sy = min(h - 1, max(0, top - 12))
        fill = pixels[sx, sy]
        from PIL import ImageDraw

        d = ImageDraw.Draw(im)
        d.rectangle((left, top, right, bottom), fill=fill)

    for box in [
        (8.62, 0.88, 0.78, 0.9),
        (11.35, 1.62, 0.9, 0.98),
        (10.0, 2.62, 1.05, 0.78),
        (6.0, 3.95, 0.98, 0.76),
        (10.28, 4.2, 1.85, 1.05),
        (7.72, 4.82, 2.05, 1.08),
        (5.95, 5.2, 1.0, 0.9),
        (5.48, 6.08, 1.3, 0.62),
    ]:
        cover(box)
    im.save(out, quality=95)
    return out


def add_text(slide, text, x, y, w, h, size, bold=False, color=NAVY, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = "Microsoft YaHei"
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def add_paras(slide, lines, x, y, w, h, size=13, color=GRAY, bullet=False, bold=False, gap=1.08):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = color
        p.line_spacing = gap
        if bullet:
            p.level = 0
            p.text = "• " + line
    return box


def rect(slide, x, y, w, h, fill, line=None, rounded=False):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h),
    )
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    return shp


def pic(slide, path, x, y, w, h):
    slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))


def line(slide, x1, y1, x2, y2, color=MID, width=1):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(width)
    return conn


def scrub(slide, x, y, w, h, fill=RGBColor(234, 244, 250)):
    shp = rect(slide, x, y, w, h, fill)
    return shp


def dot(slide, x, y, size=0.055, fill=BLUE):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x - size / 2), Inches(y - size / 2), Inches(size), Inches(size))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    return shp


def native_chart(slide, x, y, w, h, title, xlabel, ylabel, points, x_ticks, y_ticks, note):
    rect(slide, x, y, w, h, WHITE, RGBColor(222, 234, 241))
    add_text(slide, title, x + 0.08, y - 0.28, w - 0.16, 0.16, 10.5, bold=True, align=PP_ALIGN.CENTER)
    gx, gy = x + 0.42, y + 0.28
    gw, gh = w - 0.72, h - 0.72
    for i in range(4):
        yy = gy + gh * i / 3
        line(slide, gx, yy, gx + gw, yy, RGBColor(226, 235, 241), 0.7)
    line(slide, gx, gy, gx, gy + gh, RGBColor(172, 192, 205), 1)
    line(slide, gx, gy + gh, gx + gw, gy + gh, RGBColor(172, 192, 205), 1)
    mapped = []
    for px, py in points:
        xx = gx + px * gw
        yy = gy + (1 - py) * gh
        mapped.append((xx, yy))
        dot(slide, xx, yy, 0.07, BLUE)
    for (x1, y1), (x2, y2) in zip(mapped, mapped[1:]):
        line(slide, x1, y1, x2, y2, BLUE, 1.2)
    for i, t in enumerate(x_ticks):
        tx = gx + gw * i / (len(x_ticks) - 1)
        add_text(slide, t, tx - 0.17, gy + gh + 0.1, 0.34, 0.12, 6.8, color=GRAY, align=PP_ALIGN.CENTER)
    for i, t in enumerate(y_ticks):
        ty = gy + gh - gh * i / (len(y_ticks) - 1)
        add_text(slide, t, gx - 0.36, ty - 0.06, 0.28, 0.12, 6.8, color=GRAY, align=PP_ALIGN.RIGHT)
    add_text(slide, xlabel, gx + gw / 2 - 0.4, y + h - 0.16, 0.8, 0.12, 7.2, color=GRAY, align=PP_ALIGN.CENTER)
    add_text(slide, ylabel, x + 0.06, y + 0.08, 0.45, 0.12, 7.2, color=GRAY)
    add_text(slide, note, x + 0.18, y + h - 0.5, w - 0.32, 0.22, 8.2, bold=True, color=GRAY, align=PP_ALIGN.CENTER)


def bg(slide):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = PAPER


def footer(slide, page):
    rect(slide, 0, 7.08, W, 0.42, NAVY)
    add_text(slide, "中国大学生机械工程创新创意大赛", 0.62, 7.21, 3.6, 0.18, 8.5, color=WHITE)
    add_text(slide, f"{page:02d} / 10", 11.7, 7.19, 0.9, 0.2, 11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)


def page_mark(slide, page, x=0.55, y=0.45, orange=False):
    add_text(slide, f"{page:02d}", x, y, 1.25, 0.56, 39, bold=True, color=ORANGE if orange else NAVY)
    rect(slide, x + 0.05, y + 0.72, 0.72, 0.045, NAVY)
    rect(slide, x - 0.14, y + 0.08, 0.035, 0.5, ORANGE)


def title_block(slide, page, title, subtitle="", x=0.58, y=1.34, w=5.0, title_size=28, orange=False):
    page_mark(slide, page, x, 0.45, orange)
    add_text(slide, title, x, y, w, 0.85, title_size, bold=True, color=NAVY)
    if subtitle:
        add_text(slide, subtitle, x + 0.02, y + 0.78, w, 0.32, 14.5, bold=True, color=TEAL)


def tag(slide, text, x, y, w=1.75):
    rect(slide, x, y, w, 0.34, WHITE, MID, rounded=True)
    add_text(slide, text, x + 0.14, y + 0.08, w - 0.22, 0.14, 9.5, bold=True, color=TEAL)


def slide01(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    pic(s, crop_fit(1, "hero_clean", (0.52, 0.02, 1.0, 0.94), 7.2 / 6.65), 6.13, 0.0, 7.2, 7.08)
    rect(s, 0, 0, 6.05, 7.08, RGBColor(248, 252, 254))
    page_mark(s, 1, 0.58, 0.48)
    add_text(s, "金属与冰\n摩擦系数的测量", 0.58, 1.48, 5.2, 1.62, 34, bold=True)
    add_text(s, "极地环境下摩擦行为的实验研究", 0.62, 3.36, 4.7, 0.28, 15, color=GRAY)
    for i, (a, b) in enumerate([("参赛项目", "金属-冰接触界面摩擦测量"), ("作品编号", "2025-ICE-001")]):
        y = 4.27 + i * 0.58
        circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.62), Inches(y), Inches(0.28), Inches(0.28))
        circ.fill.solid(); circ.fill.fore_color.rgb = ICE; circ.line.color.rgb = BLUE
        add_text(s, a + "：", 1.02, y + 0.03, 0.82, 0.12, 9.5, bold=True, color=NAVY)
        add_text(s, b, 1.82, y + 0.03, 3.25, 0.14, 9.5, color=GRAY)
    footer(s, 1)


def slide02(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    pic(s, crop_fit(2, "ship_wide", (0.44, 0.0, 1.0, 0.83), 7.35 / 5.82), 5.98, 0.0, 7.35, 6.18)
    rect(s, 0, 0, 5.98, 7.08, RGBColor(249, 252, 254))
    title_block(s, 2, "选题背景", "冰上摩擦影响极地航运阻力与能耗", w=4.9, title_size=30)
    add_paras(s, ["北极航线缩短航程，但海冰显著增加船冰阻力。", "测量金属与冰摩擦系数，可支撑极地装备减阻设计。"], 0.66, 2.78, 4.65, 0.82, 13.5)
    add_text(s, "25%–55%", 1.02, 4.22, 2.1, 0.38, 24, bold=True, color=ORANGE)
    add_text(s, "冰阻力占船舶总阻力比例", 1.04, 4.72, 2.6, 0.18, 10.5, color=GRAY)
    add_text(s, "13%–15%", 1.02, 5.28, 2.1, 0.38, 24, bold=True, color=ORANGE)
    add_text(s, "极地航线节能潜力区间", 1.04, 5.78, 2.6, 0.18, 10.5, color=GRAY)
    footer(s, 2)


def slide03(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    title_block(s, 3, "题目解读", "从机理、装置到结果评估的研究闭环", w=6.0, title_size=30)
    items = [
        ("机理分析", "研究金属与冰界面\n滑动摩擦作用机理"),
        ("装置搭建", "设计并搭建可控的\n摩擦测试实验装置"),
        ("结果评估", "测量摩擦系数并\n评估测量不确定度"),
    ]
    for i, (name, desc) in enumerate(items):
        x = 1.45 + i * 3.95
        oval = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(2.95), Inches(1.58), Inches(1.58))
        oval.fill.solid(); oval.fill.fore_color.rgb = ICE; oval.line.color.rgb = BLUE
        add_text(s, str(i + 1), x + 0.61, 3.44, 0.34, 0.18, 22, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
        add_text(s, name, x - 0.18, 4.94, 1.95, 0.25, 15.5, bold=True, align=PP_ALIGN.CENTER)
        add_text(s, desc, x - 0.42, 5.35, 2.45, 0.48, 10.5, color=GRAY, align=PP_ALIGN.CENTER)
        if i < 2:
            arr = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 2.05), Inches(3.6), Inches(1.05), Inches(0.22))
            arr.fill.solid(); arr.fill.fore_color.rgb = MID; arr.line.fill.background()
    footer(s, 3)


def slide04(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    rect(s, 5.42, 0.72, 7.35, 5.9, RGBColor(231, 244, 251))
    rect(s, 5.42, 4.65, 7.35, 1.97, RGBColor(219, 238, 248))
    slope = s.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(5.72), Inches(2.28), Inches(6.72), Inches(3.9))
    slope.fill.solid(); slope.fill.fore_color.rgb = RGBColor(210, 232, 244); slope.fill.transparency = 35; slope.line.color.rgb = RGBColor(83, 130, 163)
    plane = s.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(6.05), Inches(4.75), Inches(6.2), Inches(0.22))
    plane.rotation = -18
    plane.fill.solid(); plane.fill.fore_color.rgb = RGBColor(7, 64, 111); plane.fill.transparency = 10; plane.line.fill.background()
    block = s.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(8.25), Inches(2.66), Inches(2.12), Inches(1.18))
    block.rotation = -18
    block.fill.solid(); block.fill.fore_color.rgb = RGBColor(169, 184, 194); block.line.color.rgb = RGBColor(74, 99, 117)
    for x, y, w, h, rot in [
        (9.35, 1.55, 0.32, 2.1, -20),
        (9.68, 3.38, 1.75, 0.22, -18),
        (7.02, 4.32, 1.6, 0.22, 162),
        (10.32, 4.0, 0.25, 1.85, 90),
        (10.74, 3.78, 0.25, 1.35, 55),
    ]:
        arr = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
        arr.rotation = rot
        arr.fill.solid(); arr.fill.fore_color.rgb = NAVY; arr.line.fill.background()
    line(s, 5.85, 5.95, 12.25, 5.95, NAVY, 1)
    add_text(s, "N", 9.48, 1.0, 0.3, 0.18, 15, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s, "v", 11.42, 2.05, 0.28, 0.16, 13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s, "f", 10.85, 2.9, 0.18, 0.16, 13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s, "f", 6.68, 4.18, 0.18, 0.16, 13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s, "mg sinθ", 10.78, 4.55, 0.85, 0.16, 11, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s, "mg cosθ", 8.15, 4.98, 0.9, 0.16, 11, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s, "θ", 6.25, 5.58, 0.18, 0.14, 12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s, "mg", 10.25, 6.05, 0.34, 0.15, 11, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    title_block(s, 4, "方法选择", "动态斜面法反推出动摩擦系数", w=4.9, title_size=30)
    rect(s, 0.72, 3.0, 4.25, 0.62, ICE, MID, rounded=True)
    add_text(s, "ma = mg sinθ − μmg cosθ", 0.93, 3.18, 3.8, 0.2, 17, bold=True, align=PP_ALIGN.CENTER)
    add_paras(s, ["记录物块沿冰面下滑过程中的位移与时间，", "结合斜面方向动力学方程反推出 μ。"], 0.76, 3.94, 4.65, 0.72, 13.5)
    add_paras(s, ["m：滑块质量", "a：运动加速度", "θ：斜面倾角", "μ：摩擦系数"], 0.82, 5.25, 2.15, 0.78, 10.5, color=SOFT)
    footer(s, 4)


def slide05(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    pic(s, crop_fit(5, "apparatus", (0.43, 0.12, 0.91, 0.82), 5.45 / 4.85), 4.45, 1.32, 5.45, 4.85)
    rect(s, 9.48, 1.32, 0.43, 4.85, PAPER)
    title_block(s, 5, "装置设计", "斜面冰面与竖直冰面双结构验证", w=4.6, title_size=30)
    add_paras(s, ["双结构实验装置用于动态滑动测量与补充验证，", "提高实验结果可靠性。"], 0.65, 2.82, 3.1, 0.58, 12.2)
    for i, t in enumerate(["斜面冰面结构", "竖直冰面结构", "低温环境模拟", "摩擦系统测量"]):
        tag(s, t, 0.75, 4.04 + i * 0.49, 2.0)
    for i, t in enumerate(["高精度传感器", "温控系统", "数据采集", "模块化结构"]):
        tag(s, t, 10.25, 2.48 + i * 0.7, 1.85)
    footer(s, 5)


def slide06(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    title_block(s, 6, "变量设计", "冰体维度与金属维度分变量控制", w=6.1, title_size=30)
    cx, cy = 6.82, 4.0
    hexagon = s.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(cx - 0.62), Inches(cy - 0.55), Inches(1.24), Inches(1.1))
    hexagon.fill.solid(); hexagon.fill.fore_color.rgb = NAVY; hexagon.line.color.rgb = NAVY
    add_text(s, "μ", cx - 0.25, cy - 0.18, 0.5, 0.26, 28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    factors = [
        ("温度", "-30℃ – 0℃", 6.1, 2.22),
        ("压强", "0.05 – 0.5 MPa", 9.42, 3.0),
        ("粗糙度", "Ra 0.02 – 1.6 μm", 9.15, 4.98),
        ("成分", "纯水 / 海水 / 含盐冰", 7.28, 5.75),
        ("面积", "接触面积变化", 4.28, 5.75),
        ("材质", "铝 / 铜 / 钢 / 合金", 3.42, 3.95),
    ]
    for name, value, x, y in factors:
        line(s, cx, cy, x + 0.28, y + 0.22, ICE2, 1.1)
        add_text(s, name, x, y, 1.08, 0.2, 14.5, bold=True)
        add_text(s, value, x, y + 0.32, 1.8, 0.16, 8.7, color=GRAY)
    footer(s, 6)


def slide07(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    title_block(s, 7, "问题改进", "针对实验关键环节进行优化", w=6.2, title_size=30)
    imgs = [
        crop_fit(7, "ice", (0.05, 0.31, 0.33, 0.72), 3.15 / 1.58),
        crop_fit(7, "temp", (0.35, 0.31, 0.63, 0.72), 3.15 / 1.58),
        crop_fit(7, "sensor", (0.66, 0.31, 0.94, 0.72), 3.15 / 1.58),
    ]
    cards = [
        ("冰面制备", "去离子水、均匀凝固\n减少气泡与裂纹"),
        ("平整度与温控", "低温箱内恒温控制\n温度波动 ±0.5℃"),
        ("传感器可靠性", "外置传感器布置\n降低低温影响"),
    ]
    for i, (title, body) in enumerate(cards):
        x = 0.72 + i * 4.08
        rect(s, x, 2.42, 3.5, 3.78, WHITE, RGBColor(217, 231, 238), rounded=True)
        pic(s, imgs[i], x + 0.16, 2.6, 3.18, 1.62)
        if i == 1:
            rect(s, x + 1.18, 2.92, 0.92, 0.28, RGBColor(2, 32, 50))
            add_text(s, "-10.0℃", x + 1.21, 2.97, 0.82, 0.12, 11.5, bold=True, color=RGBColor(0, 206, 255), align=PP_ALIGN.CENTER)
        add_text(s, title, x + 0.26, 4.55, 2.2, 0.24, 15, bold=True)
        add_text(s, body, x + 0.26, 5.02, 2.55, 0.46, 10.8, color=GRAY)
    footer(s, 7)


def slide08(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    title_block(s, 8, "数据处理", "从测量参数到摩擦系数与不确定度评估", w=7.2, title_size=30)
    steps = [
        ("原始数据采集", "x, t, θ"),
        ("加速度计算", "a"),
        ("摩擦系数计算", "μ"),
        ("不确定度分析", "U(μ)"),
    ]
    for i, (a, b) in enumerate(steps):
        x = 0.92 + i * 3.08
        rect(s, x, 3.12, 2.25, 1.75, WHITE, MID, rounded=True)
        add_text(s, str(i + 1), x + 0.18, 3.34, 0.3, 0.16, 10, color=SOFT)
        add_text(s, a, x + 0.43, 3.42, 1.42, 0.22, 12, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
        add_text(s, b, x + 0.43, 4.16, 1.42, 0.26, 17, bold=True, align=PP_ALIGN.CENTER)
        if i == 2:
            add_text(s, "μ = (g sinθ − a) / g cosθ", x + 0.16, 4.55, 1.9, 0.18, 8.2, color=GRAY, align=PP_ALIGN.CENTER)
        if i < 3:
            arr = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 2.37), Inches(3.88), Inches(0.55), Inches(0.22))
            arr.fill.solid(); arr.fill.fore_color.rgb = RGBColor(112, 170, 207); arr.line.fill.background()
    footer(s, 8)


def slide09(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    page_mark(s, 9, 0.62, 0.45, orange=True)
    add_text(s, "结果分析", 1.82, 0.78, 3.2, 0.5, 32, bold=True)
    add_text(s, "摩擦系数与实验变量的关联趋势", 1.84, 1.55, 5.2, 0.24, 14.5, bold=True, color=TEAL)
    native_chart(
        s, 0.95, 2.32, 3.35, 2.7,
        "μ - 温度关系", "温度 (℃)", "μ",
        [(0.0, 0.86), (0.2, 0.72), (0.4, 0.57), (0.6, 0.42), (0.8, 0.25), (1.0, 0.12)],
        ["-30", "-20", "-10", "0"], ["0", "0.05", "0.10", "0.15", "0.20"],
        "温度升高，摩擦系数呈下降趋势",
    )
    native_chart(
        s, 4.92, 2.32, 3.35, 2.7,
        "μ - 压强关系", "压强 (MPa)", "μ",
        [(0.0, 0.16), (0.2, 0.32), (0.4, 0.48), (0.6, 0.63), (0.8, 0.78), (1.0, 0.86)],
        ["0", "0.2", "0.4", "0.6", "0.8"], ["0", "1.0", "2.0", "3.0", "4.0"],
        "压强增大，摩擦系数同步提升",
    )
    native_chart(
        s, 8.88, 2.32, 3.35, 2.7,
        "μ - 粗糙度关系", "Ra (μm)", "μ",
        [(0.0, 0.18), (0.2, 0.35), (0.4, 0.52), (0.6, 0.66), (0.8, 0.81), (1.0, 0.9)],
        ["0.01", "0.1", "1", "10"], ["0", "1.0", "2.0", "3.0", "4.0"],
        "粗糙度增大，摩擦系数升高",
    )
    rect(s, 0.82, 5.86, 11.7, 0.75, ICE, MID, rounded=True)
    add_text(s, "温度是主要因素，其次为压强与粗糙度，材质与成分存在显著差异。", 1.18, 6.14, 10.2, 0.22, 14.5, bold=True, color=NAVY)
    footer(s, 9)


def slide10(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    pic(s, crop_fit(10, "apparatus", (0.49, 0.13, 0.98, 0.53), 5.95 / 2.55), 6.25, 1.1, 5.95, 2.55)
    pic(s, crop_fit(10, "ship", (0.49, 0.54, 0.98, 0.91), 5.95 / 2.25), 6.25, 4.02, 5.95, 2.25)
    title_block(s, 10, "总结展望", "自制装置、多变量测量、工程应用", w=5.3, title_size=31)
    rect(s, 0.72, 2.82, 4.75, 1.22, WHITE, MID, rounded=True)
    add_text(s, "研究总结", 1.02, 3.08, 1.35, 0.22, 14, bold=True)
    add_paras(s, ["搭建了双结构实验测量装置", "完成多变量影响规律分析", "建立了数据处理与不确定度评估方法"], 1.02, 3.42, 3.5, 0.48, 9.8, bullet=True)
    rect(s, 0.72, 4.42, 4.75, 1.22, WHITE, MID, rounded=True)
    add_text(s, "应用前景", 1.02, 4.68, 1.35, 0.22, 14, bold=True)
    add_paras(s, ["船体减阻与航线优化", "冰区结构与极地安全评估", "极地机器人及交通工具研发"], 1.02, 5.02, 3.5, 0.48, 9.8, bullet=True)
    footer(s, 10)


def main():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    for builder in [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10]:
        builder(prs)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
