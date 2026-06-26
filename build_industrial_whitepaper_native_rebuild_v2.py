from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
PREVIEW_DIR = ROOT / "outputs" / "whitepaper_preview_rebuild"
ASSET_DIR = ROOT / "outputs" / "industrial_whitepaper_rebuild_assets"
OUT = ROOT / "outputs" / "metal_ice_industrial_whitepaper_native_rebuild_v3.pptx"

W, H = 16, 9
FONT = "Microsoft YaHei"
PAPER = RGBColor(247, 250, 250)
WHITE = RGBColor(255, 255, 255)
TEAL = RGBColor(0, 72, 84)
TEAL_2 = RGBColor(0, 94, 110)
TEAL_LIGHT = RGBColor(226, 243, 247)
DARK = RGBColor(25, 42, 47)
MUTED = RGBColor(78, 98, 104)
LINE = RGBColor(215, 229, 233)
ICE = RGBColor(238, 248, 251)
ORANGE = RGBColor(226, 126, 39)
GREEN = RGBColor(55, 139, 82)


def crop_asset(page, name, box):
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    src = Image.open(PREVIEW_DIR / f"page_{page:02d}_preview.png").convert("RGB")
    out = ASSET_DIR / f"p{page:02d}_{name}.png"
    src.crop(box).save(out, quality=95)
    return out


def faded_asset(source: Path, name: str, opacity=0.22):
    im = Image.open(source).convert("RGB")
    paper = Image.new("RGB", im.size, (247, 250, 250))
    out = ASSET_DIR / name
    Image.blend(paper, im, opacity).save(out, quality=94)
    return out


def prepare_assets():
    assets = {
        "hero": crop_asset(1, "hero_apparatus", (590, 0, 1536, 775)),
        "inset": crop_asset(1, "inset", (525, 510, 760, 740)),
        "ship": crop_asset(2, "ship", (75, 135, 850, 425)),
        "ice": crop_asset(2, "ice", (930, 545, 1470, 750)),
        "inclined": crop_asset(5, "inclined", (245, 230, 850, 680)),
        "vertical": crop_asset(5, "vertical", (875, 235, 1195, 700)),
        "ship_small": crop_asset(10, "ship_small", (860, 205, 1210, 360)),
        "block": crop_asset(10, "block", (860, 385, 1210, 540)),
        "ship_small_2": crop_asset(10, "ship_small_2", (860, 565, 1210, 720)),
    }
    assets["ice_fade"] = faded_asset(assets["ice"], "ice_fade.png", 0.20)
    assets["ship_fade"] = faded_asset(assets["ship"], "ship_fade.png", 0.16)
    return assets


def rect(slide, x, y, w, h, fill=WHITE, line=None, radius=False, transparency=0):
    shp = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.fill.transparency = transparency
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
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
    for idx, raw in enumerate(value.split("\n")):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = raw
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return box


def line(slide, x1, y1, x2, y2, color=TEAL, width=1.2):
    shp = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    shp.line.color.rgb = color
    shp.line.width = Pt(width)
    return shp


def bullet(slide, x, y, value, size=10.4, width=3.9):
    rect(slide, x, y + 0.10, 0.055, 0.055, TEAL, radius=True)
    text(slide, x + 0.18, y, width, 0.24, value, size, DARK)


def base(slide, page, title, subtitle=None, assets=None):
    rect(slide, 0, 0, W, H, PAPER)
    rect(slide, 0, 7.90, W, 1.10, RGBColor(240, 248, 250), transparency=15000)
    if assets:
        slide.shapes.add_picture(str(assets["ice_fade"]), Inches(11.0), Inches(7.35), width=Inches(4.75), height=Inches(1.35))
        line(slide, 0.60, 7.86, 15.10, 7.86, RGBColor(224, 236, 240), 0.8)
    rect(slide, -0.08, 0.08, 1.45, 0.66, TEAL)
    text(slide, 0.20, 0.20, 0.62, 0.30, f"{page:02d}", 22, WHITE, True, name=f"S{page:02d}_page")
    text(slide, 1.42, 0.20, 7.2, 0.34, title, 21.5, TEAL, True, name=f"S{page:02d}_title")
    line(slide, 0.72, 0.84, 1.22, 0.84, TEAL, 3)
    line(slide, 1.34, 0.84, 2.56, 0.84, TEAL_2, 1.2)
    if subtitle:
        text(slide, 1.42, 0.96, 6.2, 0.22, subtitle, 11.8, MUTED, name=f"S{page:02d}_subtitle")
    rect(slide, 15.08, 8.51, 0.82, 0.36, TEAL)
    text(slide, 15.34, 8.59, 0.28, 0.12, f"{page:02d}", 10, WHITE, True, PP_ALIGN.CENTER, name=f"S{page:02d}_footer")


def card(slide, x, y, w, h):
    return rect(slide, x, y, w, h, WHITE, LINE, True)


def circle(slide, cx, cy, label, glyph):
    shp = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(cx - 0.42), Inches(cy - 0.42), Inches(0.84), Inches(0.84))
    shp.fill.solid()
    shp.fill.fore_color.rgb = TEAL_LIGHT
    shp.line.color.rgb = RGBColor(177, 216, 225)
    text(slide, cx - 0.22, cy - 0.16, 0.44, 0.18, glyph, 13, TEAL, True, PP_ALIGN.CENTER)
    text(slide, cx - 0.68, cy + 0.54, 1.36, 0.18, label, 10.8, TEAL, True, PP_ALIGN.CENTER)


def s01(slide, a):
    rect(slide, 0, 0, W, H, RGBColor(246, 251, 252))
    slide.shapes.add_picture(str(a["hero"]), Inches(5.90), Inches(0.00), width=Inches(9.85), height=Inches(7.45))
    rect(slide, 5.70, 0, 2.0, H, RGBColor(246, 251, 252), transparency=42000)
    rect(slide, -0.08, 0.08, 1.45, 0.66, TEAL)
    text(slide, 0.20, 0.20, 0.62, 0.30, "01", 22, WHITE, True)
    text(slide, 0.78, 1.48, 5.30, 1.34, "金属与冰\n摩擦系数的测量", 33, TEAL, True)
    line(slide, 0.80, 3.42, 1.34, 3.42, ORANGE, 2.4)
    line(slide, 1.46, 3.42, 2.60, 3.42, LINE, 1.2)
    text(slide, 0.78, 3.72, 3.5, 0.30, "物理实验创新竞赛答辩", 16, DARK, True)
    slide.shapes.add_picture(str(a["inset"]), Inches(4.50), Inches(4.75), width=Inches(1.55), height=Inches(1.55))
    text(slide, 1.18, 4.88, 2.15, 0.20, "团队：极冰钥匙者", 11.5, DARK)
    text(slide, 1.18, 5.58, 2.15, 0.20, "日期：2024.05.20", 11.5, DARK)
    rect(slide, 0.70, 7.95, 5.05, 0.38, TEAL)
    text(slide, 1.02, 8.04, 4.25, 0.16, "物理实验 | 摩擦学研究 | 精密测量 | 工程创新", 9.4, WHITE, True)
    rect(slide, 15.08, 8.51, 0.82, 0.36, TEAL)
    text(slide, 15.34, 8.59, 0.28, 0.12, "01", 10, WHITE, True, PP_ALIGN.CENTER)


def s02(slide, a):
    base(slide, 2, "选题背景", assets=a)
    slide.shapes.add_picture(str(a["ship"]), Inches(0.90), Inches(1.35), width=Inches(7.4), height=Inches(3.15))
    card(slide, 0.90, 5.10, 7.4, 1.50)
    text(slide, 1.25, 5.30, 1.2, 0.20, "研究意义", 12.4, TEAL, True)
    for i, item in enumerate(["揭示船舶破冰与航行效率的关键摩擦机理之一", "为材料选型、表面处理与结构设计提供依据", "提升极地装备低温接触工况下的安全性"]):
        bullet(slide, 1.18, 5.68 + i * 0.30, item, 9.8, 5.2)
    text(slide, 9.00, 1.35, 4.60, 0.26, "北极航运与冰阻力问题", 16.5, TEAL, True)
    text(slide, 9.00, 2.05, 4.2, 0.20, "航行阻力中，冰障碍阻力占比", 11.0, MUTED)
    text(slide, 9.00, 2.38, 2.10, 0.32, "25%-55%", 21, ORANGE, True)
    text(slide, 9.00, 3.40, 4.2, 0.20, "冰面相关事故中，摩擦失控因素占比", 11.0, MUTED)
    text(slide, 9.00, 3.72, 2.10, 0.32, "13%-15%", 21, ORANGE, True)
    slide.shapes.add_picture(str(a["ice"]), Inches(9.00), Inches(5.20), width=Inches(5.7), height=Inches(2.0))


def s03(slide, a):
    base(slide, 3, "题目解读", assets=a)
    columns = [("机理分析", "⬡", ["冰的黏附性变形", "表面微观水膜", "水膜润滑作用", "温度与压力影响"]),
               ("装置搭建", "⚒", ["斜面摩擦实验平台", "力学测量与采集系统", "低温环境控制", "标定与校准"]),
               ("结果与不确定度", "↗", ["摩擦系数计算", "不确定度评估", "敏感性分析", "结果讨论与验证"])]
    for i, (head, glyph, lines_) in enumerate(columns):
        x = 1.05 + i * 4.65
        card(slide, x, 1.28, 4.25, 5.95)
        text(slide, x + 0.55, 1.68, 3.15, 0.24, head, 15.2, TEAL, True, PP_ALIGN.CENTER)
        circle(slide, x + 2.12, 2.95, "", glyph)
        for j, item in enumerate(lines_):
            bullet(slide, x + 0.70, 4.18 + j * 0.44, item, 10.8, 2.8)


def s04(slide, a):
    base(slide, 4, "方法选择 —— 动态斜面法", assets=a)
    text(slide, 1.10, 1.24, 1.4, 0.20, "力学模型", 12.8, TEAL, True)
    line(slide, 1.45, 6.55, 7.10, 3.55, RGBColor(169, 181, 184), 2)
    rect(slide, 3.85, 4.48, 1.42, 0.64, RGBColor(218, 225, 226), RGBColor(150, 166, 170))
    line(slide, 4.55, 4.82, 4.55, 2.32, TEAL, 1.5)
    line(slide, 4.55, 4.82, 6.35, 3.90, ORANGE, 1.5)
    line(slide, 4.55, 4.82, 3.25, 5.48, ORANGE, 1.5)
    line(slide, 1.78, 6.25, 2.42, 6.25, DARK, 0.8)
    text(slide, 4.70, 2.16, 0.25, 0.15, "N", 10, TEAL, True)
    text(slide, 6.42, 3.78, 0.25, 0.15, "a", 10, ORANGE, True)
    text(slide, 3.05, 5.42, 0.25, 0.15, "f", 10, ORANGE, True)
    text(slide, 8.88, 1.24, 1.4, 0.20, "理论公式", 12.8, TEAL, True)
    card(slide, 8.98, 1.72, 4.55, 0.64)
    text(slide, 9.32, 1.92, 3.90, 0.20, "a = g(sinθ − μcosθ)", 15.8, DARK, False, PP_ALIGN.CENTER)
    text(slide, 11.10, 2.60, 0.28, 0.20, "↓", 18, TEAL, True, PP_ALIGN.CENTER)
    card(slide, 8.98, 3.00, 4.55, 0.78)
    text(slide, 9.25, 3.22, 4.05, 0.24, "μ = (sinθ − a/g) / cosθ", 15.8, DARK, False, PP_ALIGN.CENTER)
    text(slide, 9.16, 4.18, 3.7, 1.00, "a —— 沿斜面方向的加速度\ng —— 重力加速度\nθ —— 斜面倾角\nμ —— 动摩擦系数", 11.2, DARK)


def s05(slide, a):
    base(slide, 5, "装置设计", assets=a)
    text(slide, 1.00, 1.14, 2.20, 0.18, "斜面冰面模块", 12.6, TEAL, True)
    slide.shapes.add_picture(str(a["inclined"]), Inches(1.48), Inches(1.58), width=Inches(6.05), height=Inches(4.48))
    text(slide, 8.42, 1.14, 3.85, 0.18, "垂直冰面模块（压力加载）", 12.6, TEAL, True)
    slide.shapes.add_picture(str(a["vertical"]), Inches(9.25), Inches(1.58), width=Inches(3.82), height=Inches(4.48))
    for i, item in enumerate(["导轨与滑块", "力传感器", "位移传感器", "冰面板", "角度调节机构", "支撑框架"]):
        text(slide, 0.72, 1.82 + i * 0.47, 1.18, 0.15, item, 8.8, MUTED)
    for i, item in enumerate(["压力加载装置", "力传感器", "冰面板", "温控舱体", "底盘"]):
        text(slide, 13.10, 1.84 + i * 0.68, 1.24, 0.15, item, 8.8, MUTED)
    for i, item in enumerate(["模块化设计", "快速更换", "低温环境模拟", "高精度测量"]):
        circle(slide, 1.65 + i * 3.45, 7.28, item, ["◎", "↔", "❄", "✓"][i])


def s06(slide, a):
    base(slide, 6, "变量设计", "六大影响变量与水平设置", a)
    items = [("温度", "❄", "-5 °C\n-10 °C\n-15 °C"), ("压强", "P", "0.05 MPa\n0.10 MPa\n0.20 MPa"), ("材质", "▰", "铝合金\n不锈钢\n钛合金"), ("粗糙度", "≈", "Ra 0.2 μm\nRa 1.0 μm\nRa 3.0 μm"), ("面积", "□", "1 cm²\n4 cm²\n9 cm²"), ("成分", "⌬", "淡水冰\n海冰\n人工掺盐冰")]
    for i, (head, glyph, vals) in enumerate(items):
        x = 1.20 + i * 2.42
        circle(slide, x, 2.60, head, glyph)
        text(slide, x - 0.62, 3.85, 1.24, 0.72, vals, 10.0, DARK, False, PP_ALIGN.CENTER)


def s07(slide, a):
    base(slide, 7, "问题改进", assets=a)
    rows = [("冰面制备不平整", "表面粗糙和气泡较大\n影响测量重复性", "定制制冰模具 + 控温体系\n实现均匀透明平整冰面"),
            ("低温环境稳定性差", "温度波动导致数据漂移\n影响测量精度", "双层保温 + PID温控\n温度波动 ≤ ±0.2 °C"),
            ("传感器可靠性不足", "低温漂移与噪声干扰大\n影响数据信号", "低温型传感器 + 屏蔽布线\n多点校准与滤波处理")]
    for i, (p, pd, sv) in enumerate(rows):
        y = 1.45 + i * 1.75
        card(slide, 0.90, y, 5.35, 1.18)
        text(slide, 1.65, y + 0.18, 2.36, 0.20, p, 11.2, DARK, True)
        text(slide, 1.65, y + 0.48, 2.60, 0.36, pd, 9.5, MUTED)
        text(slide, 6.65, y + 0.38, 0.45, 0.25, "→", 20, RGBColor(141, 178, 186), True, PP_ALIGN.CENTER)
        card(slide, 7.28, y, 5.25, 1.18)
        text(slide, 7.60, y + 0.18, 1.10, 0.20, "改进方案", 11.2, GREEN, True)
        text(slide, 7.60, y + 0.48, 3.20, 0.36, sv, 9.5, DARK)
        text(slide, 12.05, y + 0.38, 0.30, 0.20, "✓", 17, GREEN, True, PP_ALIGN.CENTER)


def s08(slide, a):
    base(slide, 8, "数据处理", assets=a)
    steps = [("数据采集", "⌁", ["力信号", "位移信号", "温度/压力", "采样频率 100 Hz"]), ("加速度计算", "↗", ["位移二次微分", "滤波处理", "得到 a(t)"]), ("摩擦系数计算", "μ", ["μ = (sinθ − a/g) / cosθ", "代入角度与加速度", "计算 μ"]), ("不确定度评估", "∩", ["A类不确定度", "B类不确定度", "合成标准不确定度", "扩展不确定度 U"])]
    for i, (head, glyph, vals) in enumerate(steps):
        x = 1.05 + i * 3.65
        card(slide, x, 1.55, 2.75, 5.65)
        text(slide, x + 0.55, 2.10, 1.65, 0.20, head, 11.4, TEAL, True, PP_ALIGN.CENTER)
        text(slide, x + 1.05, 3.05, 0.65, 0.45, glyph, 24, TEAL, False, PP_ALIGN.CENTER)
        if i == 2:
            text(slide, x + 0.22, 4.18, 2.25, 0.42, "μ = (sinθ − a/g)\n      / cosθ", 12.2, DARK, False, PP_ALIGN.CENTER)
            for j, item in enumerate(vals[1:]):
                bullet(slide, x + 0.42, 5.50 + j * 0.30, item, 8.8, 1.8)
        else:
            for j, item in enumerate(vals):
                bullet(slide, x + 0.42, 5.02 + j * 0.30, item, 8.8, 1.8)
        if i < 3:
            text(slide, x + 2.95, 4.20, 0.36, 0.25, "→", 18, RGBColor(141, 178, 186), True, PP_ALIGN.CENTER)


def s09(slide, a):
    base(slide, 9, "结果分析", assets=a)
    text(slide, 2.25, 1.22, 4.10, 0.20, "摩擦系数对比趋势（示例）", 12.4, TEAL, True, PP_ALIGN.CENTER)
    x0, y0, cw, ch = 1.35, 1.90, 6.2, 4.25
    for k in range(6):
        y = y0 + k * ch / 5
        line(slide, x0, y, x0 + cw, y, RGBColor(229, 237, 239), 0.6)
        text(slide, x0 - 0.44, y - 0.06, 0.32, 0.12, f"{0.15 - k * 0.03:.2f}", 7.2, DARK, False, PP_ALIGN.RIGHT)
    line(slide, x0, y0, x0, y0 + ch, DARK, 0.8)
    line(slide, x0, y0 + ch, x0 + cw, y0 + ch, DARK, 0.8)
    for i, val in enumerate([-15, -10, -5, 0]):
        px = x0 + 0.55 + i * (cw - 1.10) / 3
        text(slide, px - 0.20, y0 + ch + 0.14, 0.40, 0.12, str(val), 7.4, DARK, False, PP_ALIGN.CENTER)
    data = [("铝合金", RGBColor(41, 137, 190), [0.047, 0.040, 0.030, 0.020]), ("不锈钢", ORANGE, [0.116, 0.105, 0.087, 0.064]), ("钛合金", RGBColor(112, 164, 52), [0.082, 0.070, 0.058, 0.039])]
    for si, (name, color, vals) in enumerate(data):
        pts = []
        for i, v in enumerate(vals):
            px = x0 + 0.55 + i * (cw - 1.10) / 3
            py = y0 + ch - (v / 0.15) * ch
            pts.append((px, py))
        for p1, p2 in zip(pts, pts[1:]):
            line(slide, p1[0], p1[1], p2[0], p2[1], color, 1.4)
        for px, py in pts:
            dot = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(px - 0.035), Inches(py - 0.035), Inches(0.07), Inches(0.07))
            dot.fill.solid()
            dot.fill.fore_color.rgb = color
            dot.line.color.rgb = color
        text(slide, 7.75, 2.10 + si * 0.28, 0.9, 0.14, name, 7.4, DARK)
    text(slide, 3.50, 6.40, 1.45, 0.16, "温度 / °C", 8.8, DARK, False, PP_ALIGN.CENTER)
    text(slide, 0.70, 3.30, 0.42, 0.90, "摩擦系数 μ", 8.8, DARK)
    text(slide, 9.05, 1.42, 1.25, 0.20, "数据解读", 12.6, TEAL, True)
    for i, item in enumerate(["温度升高，μ整体下降", "粗糙度增大，μ上升", "压强增大，μ先升后降", "材质差异显著，钛合金综合表现更优"]):
        bullet(slide, 9.05, 2.02 + i * 0.45, item, 9.8, 3.7)
    card(slide, 9.05, 5.46, 3.75, 0.88)
    text(slide, 9.42, 5.72, 2.85, 0.34, "结果与文献趋势一致，\n验证方法的有效性。", 10.4, TEAL, True, PP_ALIGN.CENTER)


def s10(slide, a):
    base(slide, 10, "总结展望", assets=a)
    text(slide, 2.20, 1.18, 1.10, 0.18, "创新点", 11.8, TEAL, True)
    for i, item in enumerate(["模块化多工况装置设计", "低温稳定控制与高精度测量", "多因素系统研究与量化分析", "不确定度评估与可靠性提升"]):
        card(slide, 1.05, 1.55 + i * 1.00, 5.50, 0.70)
        text(slide, 1.60, 1.76 + i * 1.00, 3.70, 0.18, item, 10.4, DARK)
    text(slide, 10.72, 1.18, 1.90, 0.18, "工程应用前景", 11.8, TEAL, True)
    imgs = [a["ship_small"], a["block"], a["ship_small_2"]]
    apps = [("极地船舶设计优化", "降低航行阻力"), ("表面工程与减阻开发", "提升抗冰性能"), ("极地装备与结构安全", "提供理论依据")]
    for i, (head, body) in enumerate(apps):
        y = 1.55 + i * 1.35
        slide.shapes.add_picture(str(imgs[i]), Inches(8.05), Inches(y), width=Inches(2.75), height=Inches(1.00))
        card(slide, 10.95, y, 3.45, 1.00)
        text(slide, 11.15, y + 0.18, 2.10, 0.18, head, 10.4, TEAL, True)
        text(slide, 11.15, y + 0.48, 1.60, 0.16, body, 8.8, MUTED)
    card(slide, 1.05, 7.22, 13.55, 0.55)
    text(slide, 1.60, 7.40, 11.70, 0.18, "未来将拓展更多材料与工况，建立摩擦数据库，服务极地工程与安全航行。", 9.0, DARK)


def build():
    assets = prepare_assets()
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    for fn in [s01, s02, s03, s04, s05, s06, s07, s08, s09, s10]:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        fn(slide, assets)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
