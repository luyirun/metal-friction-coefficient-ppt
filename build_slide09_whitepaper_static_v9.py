from pathlib import Path
from math import cos, sin, pi

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
PROJECT = ROOT / "projects" / "slide09_whitepaper_static_v9"
ASSET_DIR = PROJECT / "assets" / "imagegen"
OUT_DIR = PROJECT / "exports"
VALIDATION_DIR = PROJECT / "validation"
OUT = OUT_DIR / "slide09_whitepaper_static_v9.pptx"

W, H = 16, 9
FONT = "Microsoft YaHei"

WHITE = RGBColor(246, 253, 255)       # #F6FDFF
MIST = RGBColor(186, 211, 220)        # #BAD3DC
MUTED = RGBColor(132, 165, 178)       # #84A5B2
ICE = RGBColor(132, 223, 242)         # #84DFF2
ICE_SOFT = RGBColor(81, 178, 202)     # #51B2CA
NAVY_LINE = RGBColor(25, 92, 117)     # #195C75
ORANGE = RGBColor(226, 136, 64)       # #E28840

BG = ASSET_DIR / "slide09_v9_deep_navy_background.png"
EXPERIMENT = ASSET_DIR / "slide09_v9_experiment_evidence.png"
DECOR = ASSET_DIR / "slide09_v9_radial_decor.png"

FACTORS = [
    ("F01", "温度", "Temperature", "-15 / -10 / -5 / 0 ℃"),
    ("F02", "压强", "Pressure", "0.05 / 0.10 / 0.20 MPa"),
    ("F03", "金属材质", "Metal Material", "Al / Steel / Ti"),
    ("F04", "表面粗糙度", "Surface Roughness", "Ra 0.2 / 1.0 / 3.0 μm"),
    ("F05", "接触面积", "Contact Area", "1 / 4 / 9 cm²"),
    ("F06", "冰的成分", "Ice Composition", "淡水冰 / 海冰 / 盐冰"),
]


def hex_line(shp, color=ICE_SOFT, width=0.75, transparency=0):
    shp.line.color.rgb = color
    shp.line.width = Pt(width)
    shp.line.transparency = transparency


def shape(slide, kind, x, y, w, h, line=ICE_SOFT, width=0.75, transparency=0, name=None):
    shp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        shp.name = name
    shp.fill.background()
    if line is None:
        shp.line.fill.background()
    else:
        hex_line(shp, line, width, transparency)
    return shp


def rect(slide, x, y, w, h, line=ICE_SOFT, width=0.75, transparency=0, name=None):
    return shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, h, line, width, transparency, name)


def oval(slide, x, y, w, h, line=ICE_SOFT, width=0.75, transparency=0, name=None):
    return shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, x, y, w, h, line, width, transparency, name)


def line(slide, x1, y1, x2, y2, color=ICE_SOFT, width=0.75, transparency=0, name=None):
    shp = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    if name:
        shp.name = name
    shp.line.color.rgb = color
    shp.line.width = Pt(width)
    shp.line.transparency = transparency
    return shp


def text(slide, x, y, w, h, value, size, color=WHITE, bold=False, align=PP_ALIGN.LEFT, name=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    box.fill.background()
    box.line.fill.background()
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
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


def add_header(slide):
    line(slide, 0.62, 0.58, 1.22, 0.58, ICE, 2.0, 5)
    text(slide, 0.42, 0.28, 0.52, 0.28, "09", 20, WHITE, True, PP_ALIGN.CENTER, "page_no")
    text(slide, 1.14, 0.25, 1.55, 0.30, "结果分析", 21, WHITE, True, name="title_cn")
    text(slide, 2.82, 0.33, 4.90, 0.16, "Controlled Variable Framework for Friction Coefficient", 8.9, MIST, False, name="title_en")
    line(slide, 1.14, 0.68, 2.32, 0.68, ICE, 1.3, 12)
    line(slide, 2.50, 0.68, 6.55, 0.68, NAVY_LINE, 0.7, 25)


def add_hex(slide, cx, cy, r):
    pts = []
    for i in range(6):
        a = pi / 6 + i * pi / 3
        pts.append((cx + r * cos(a), cy + r * sin(a)))
    for i in range(6):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % 6]
        line(slide, x1, y1, x2, y2, ICE, 0.6, 15)
    for x, y in pts[::2]:
        line(slide, cx, cy, x, y, ICE_SOFT, 0.45, 38)
    oval(slide, cx - 0.024, cy - 0.024, 0.048, 0.048, ICE, 0.3, 0)


def add_factor_node(slide, cx, cy, idx):
    oval(slide, cx - 0.44, cy - 0.44, 0.88, 0.88, ICE, 0.9, 12, f"factor_node_{idx:02d}")
    oval(slide, cx - 0.31, cy - 0.31, 0.62, 0.62, ICE_SOFT, 0.55, 42, f"factor_node_inner_{idx:02d}")
    add_hex(slide, cx, cy - 0.02, 0.14)
    text(slide, cx - 0.16, cy + 0.12, 0.32, 0.08, f"F{idx:02d}", 4.8, ICE, True, PP_ALIGN.CENTER, f"factor_node_id_{idx:02d}")
    oval(slide, cx - 0.045, cy + 0.44, 0.09, 0.09, ORANGE, 0.0, 0, f"factor_dot_{idx:02d}")


def add_factor_label(slide, x, y, w, factor, name):
    fid, cn, en, level = factor
    rect(slide, x, y, w, 0.50, ICE_SOFT, 0.45, 58, f"{name}_outline")
    line(slide, x, y, x, y + 0.50, ICE, 1.4, 12, f"{name}_bar")
    text(slide, x + 0.10, y + 0.07, 0.28, 0.08, fid, 5.0, ICE, True, name=f"{name}_id")
    text(slide, x + 0.42, y + 0.05, 0.74, 0.11, cn, 7.8, WHITE, True, name=f"{name}_cn")
    text(slide, x + 0.42, y + 0.22, 1.00, 0.08, en, 4.7, MIST, False, name=f"{name}_en")
    text(slide, x + 0.42, y + 0.36, w - 0.52, 0.08, level, 4.5, MUTED, False, name=f"{name}_level")


def add_matrix(slide):
    # Structural frame only, no fill.
    rect(slide, 0.62, 1.08, 10.68, 6.84, ICE_SOFT, 0.7, 62, "matrix_frame")
    line(slide, 0.62, 1.08, 11.30, 1.08, ICE, 1.7, 20, "matrix_top_rule")
    text(slide, 0.92, 1.38, 1.22, 0.12, "PRIMARY MODEL", 6.1, MIST, True, name="matrix_kicker")
    text(slide, 0.92, 1.59, 2.25, 0.21, "六因素控制变量矩阵", 13.4, WHITE, True, name="matrix_title")
    text(slide, 7.30, 1.38, 3.34, 0.32, "变量围绕 μ 响应量组织；实验装置仅提供测量语境。", 7.0, MIST, False, PP_ALIGN.RIGHT, "matrix_sub")

    cx, cy = 5.96, 4.22
    # Optional decorative image behind the editable matrix, low visual weight.
    if DECOR.exists():
        pic = slide.shapes.add_picture(str(DECOR), Inches(2.35), Inches(1.48), width=Inches(7.10), height=Inches(4.74))
        pic.name = "radial_decor_image"
        # PowerPoint image transparency through python-pptx is limited; keep it behind all native objects.

    for r, alpha, width in [(2.52, 38, 0.7), (1.95, 50, 0.58), (1.30, 60, 0.55), (0.78, 55, 0.5)]:
        oval(slide, cx - r, cy - r, 2 * r, 2 * r, ICE_SOFT, width, alpha)
    for deg in [0, 30, 60, 90, 120, 150]:
        a = deg * pi / 180
        line(slide, cx - 2.55 * cos(a), cy - 2.55 * sin(a), cx + 2.55 * cos(a), cy + 2.55 * sin(a), ICE_SOFT, 0.42, 68)

    positions = [
        (cx, cy - 2.02),
        (cx + 1.92, cy - 1.08),
        (cx + 1.92, cy + 1.08),
        (cx, cy + 2.02),
        (cx - 1.92, cy + 1.08),
        (cx - 1.92, cy - 1.08),
    ]
    for idx, (nx, ny) in enumerate(positions, start=1):
        line(slide, cx, cy, nx, ny, ICE_SOFT, 0.55, 55, f"factor_connector_{idx:02d}")
    for idx, (nx, ny) in enumerate(positions, start=1):
        add_factor_node(slide, nx, ny, idx)

    oval(slide, cx - 0.70, cy - 0.70, 1.40, 1.40, ICE_SOFT, 0.65, 40, "response_ring")
    oval(slide, cx - 0.52, cy - 0.52, 1.04, 1.04, ICE, 0.9, 14, "response_core_outline")
    text(slide, cx - 0.28, cy - 0.34, 0.56, 0.36, "μ", 32, WHITE, True, PP_ALIGN.CENTER, "response_mu")
    text(slide, cx - 0.36, cy + 0.19, 0.72, 0.08, "response", 5.3, MIST, False, PP_ALIGN.CENTER, "response_label")

    label_specs = [
        (4.82, 1.82, 2.25, FACTORS[0], "label_f01"),
        (8.36, 3.08, 2.08, FACTORS[1], "label_f02"),
        (8.36, 5.10, 2.08, FACTORS[2], "label_f03"),
        (4.82, 6.50, 2.25, FACTORS[3], "label_f04"),
        (1.36, 5.10, 2.08, FACTORS[4], "label_f05"),
        (1.36, 3.08, 2.08, FACTORS[5], "label_f06"),
    ]
    for lx, ly, lw, factor, lname in label_specs:
        add_factor_label(slide, lx, ly, lw, factor, lname)

    # Small editable chart accent, no fill.
    line(slide, 0.94, 7.02, 2.14, 7.02, ICE_SOFT, 0.5, 58)
    line(slide, 0.94, 7.02, 0.94, 6.22, ICE_SOFT, 0.5, 58)
    points = [(1.02, 6.91), (1.24, 6.88), (1.46, 6.79), (1.68, 6.65), (1.90, 6.42)]
    for i, (x, y) in enumerate(points):
        if i:
            line(slide, points[i - 1][0], points[i - 1][1], x, y, ICE, 0.45, 38)
        oval(slide, x - 0.018, y - 0.018, 0.036, 0.036, ORANGE if i == 2 else ICE, 0.0, 0)

    labels = ["控制变量", "实验水平", "μ 数据采集", "趋势拟合", "结果解释", "模型回写"]
    x0, y0, step_w = 0.92, 7.52, 1.48
    line(slide, x0, y0 - 0.12, x0 + step_w * 6, y0 - 0.12, ICE_SOFT, 0.55, 62)
    for i, label in enumerate(labels):
        x = x0 + i * step_w
        if i:
            line(slide, x, y0 - 0.12, x, y0 + 0.38, ICE_SOFT, 0.45, 68)
        text(slide, x + 0.08, y0 + 0.00, 0.24, 0.08, f"{i+1:02d}", 4.8, ICE, True)
        text(slide, x + 0.08, y0 + 0.18, 1.10, 0.10, label, 5.9, WHITE, True)


def add_right(slide):
    line(slide, 11.62, 1.10, 11.62, 7.88, ICE_SOFT, 0.8, 65)
    text(slide, 12.00, 1.24, 2.64, 0.22, "实验手段选择 + 模型构建", 12.5, WHITE, True, name="right_title")
    text(slide, 12.00, 1.58, 2.85, 0.34, "作为证据与方法层，解释六因素矩阵如何落到实验方案和建模链路。", 6.8, MIST, False, name="right_intro")
    rect(slide, 12.00, 2.12, 2.74, 1.56, ICE_SOFT, 0.55, 58, "experiment_frame")
    if EXPERIMENT.exists():
        pic = slide.shapes.add_picture(str(EXPERIMENT), Inches(12.14), Inches(2.23), width=Inches(2.46), height=Inches(1.34))
        pic.name = "experiment_evidence_image"
    text(slide, 12.18, 3.78, 2.40, 0.10, "金属-冰接触实验语境（辅助证据层）", 6.0, MIST, False, PP_ALIGN.CENTER, "experiment_caption")

    x_axis = 12.32
    line(slide, x_axis, 4.32, x_axis, 7.23, ICE_SOFT, 0.65, 45)
    methods = [
        ("01", "控制变量法", "一次只改变一个因素，保持其余条件稳定。"),
        ("02", "实验水平设计", "为每个因素设置离散水平，形成可比较序列。"),
        ("03", "μ 数据采集", "从接触界面测量摩擦响应并记录波动。"),
        ("04", "趋势建模", "拟合单因素趋势并识别主导影响区间。"),
        ("05", "结果分析", "回到物理机制解释并支撑结论表达。"),
    ]
    y = 4.28
    for idx, (num, title, desc) in enumerate(methods):
        oval(slide, x_axis - 0.05, y + 0.02, 0.10, 0.10, ICE, 0.75, 15)
        oval(slide, x_axis - 0.022, y + 0.048, 0.044, 0.044, ORANGE if idx == 4 else ICE, 0.0, 0)
        text(slide, 12.54, y - 0.02, 0.28, 0.08, num, 5.1, ICE, True)
        text(slide, 12.88, y - 0.04, 1.32, 0.11, title, 7.0, WHITE, True)
        text(slide, 12.88, y + 0.15, 2.05, 0.19, desc, 5.4, MIST, False)
        y += 0.58
    rect(slide, 12.00, 7.42, 2.88, 0.35, ICE_SOFT, 0.5, 66)
    text(slide, 12.20, 7.53, 2.52, 0.09, "静态版仅呈现模型总览；交互热区将在确认后单独添加。", 5.3, MIST, False, PP_ALIGN.CENTER)


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    VALIDATION_DIR.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if BG.exists():
        slide.shapes.add_picture(str(BG), Inches(0), Inches(0), width=Inches(W), height=Inches(H)).name = "deep_navy_background_texture"
    add_header(slide)
    add_matrix(slide)
    add_right(slide)
    text(slide, 0.74, 8.45, 6.70, 0.12, "注：当前数据为版式示例，后续可替换为真实实验数据；本页未添加触发器动画。", 5.8, MIST, False, name="footer_note")
    end = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM, Inches(14.64), Inches(8.34), Inches(0.88), Inches(0.34))
    end.fill.background()
    hex_line(end, ICE, 0.75, 8)
    text(slide, 14.87, 8.43, 0.36, 0.12, "09", 8.0, WHITE, True, PP_ALIGN.CENTER)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
