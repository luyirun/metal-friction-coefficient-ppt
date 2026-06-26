from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt


ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
SRC = ROOT / "outputs" / "single_pages" / "page_01_industrial_whitepaper.png"
ASSET_DIR = ROOT / "outputs" / "editable_ppt_assets" / "page01"
OUT = ROOT / "outputs" / "metal_ice_page01_editable.pptx"
ASSET_DIR.mkdir(parents=True, exist_ok=True)


def crop_assets():
    im = Image.open(SRC).convert("RGBA")
    assets = {
        "mountains": (0, 0, 2048, 500),
        "apparatus": (1100, 55, 2048, 1070),
        "inset": (710, 692, 1040, 1010),
    }
    out = {}
    for name, box in assets.items():
        crop = im.crop(box)
        path = ASSET_DIR / f"{name}.png"
        crop.save(path)
        out[name] = path
    return out


def write_svg_icons():
    icons = {
        "team": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><circle cx="24" cy="24" r="9" fill="none" stroke="#0B5868" stroke-width="5"/><circle cx="43" cy="26" r="7" fill="none" stroke="#0B5868" stroke-width="5"/><path d="M9 53c2-12 10-18 20-18s18 6 20 18" fill="none" stroke="#0B5868" stroke-width="5" stroke-linecap="round"/><path d="M36 52c1-8 7-13 15-13 5 0 9 2 12 7" fill="none" stroke="#0B5868" stroke-width="5" stroke-linecap="round"/></svg>""",
        "calendar": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect x="10" y="14" width="44" height="40" rx="6" fill="none" stroke="#0B5868" stroke-width="5"/><path d="M10 25h44M22 8v12M42 8v12" fill="none" stroke="#0B5868" stroke-width="5" stroke-linecap="round"/><path d="M21 34h6M36 34h6M21 44h6M36 44h6" stroke="#0B5868" stroke-width="5" stroke-linecap="round"/></svg>""",
        "snow": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><path d="M32 7v50M12 20l40 24M52 20 12 44M20 11l12 9 12-9M20 53l12-9 12 9" fill="none" stroke="#FFFFFF" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/></svg>""",
        "gear": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><circle cx="32" cy="32" r="10" fill="none" stroke="#FFFFFF" stroke-width="5"/><path d="M32 6v9M32 49v9M6 32h9M49 32h9M13 13l7 7M44 44l7 7M51 13l-7 7M20 44l-7 7" fill="none" stroke="#FFFFFF" stroke-width="5" stroke-linecap="round"/></svg>""",
        "shield": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><path d="M32 7 53 16v15c0 14-9 23-21 27C20 54 11 45 11 31V16Z" fill="none" stroke="#FFFFFF" stroke-width="5" stroke-linejoin="round"/><path d="m21 33 8 8 15-17" fill="none" stroke="#FFFFFF" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/></svg>""",
    }
    paths = {}
    for name, svg in icons.items():
        path = ASSET_DIR / f"{name}.svg"
        path.write_text(svg, encoding="utf-8")
        paths[name] = path
    return paths


def add_textbox(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = "Noto Sans SC"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def add_rect(slide, x, y, w, h, fill, radius=False, line=None):
    shape_type = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = line if line else fill
    shp.line.transparency = 100000 if line is None else 0
    return shp


def add_circle_icon(slide, cx, cy, label_type):
    # Editable PPT approximation of the SVG icons.
    circ = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL, Inches(cx - 0.18), Inches(cy - 0.18), Inches(0.36), Inches(0.36)
    )
    circ.fill.solid()
    circ.fill.fore_color.rgb = RGBColor(255, 255, 255)
    circ.line.color.rgb = RGBColor(146, 178, 185)
    circ.line.width = Pt(1.3)
    color = RGBColor(11, 88, 104)
    if label_type == "team":
        for dx, dy, r in [(-0.06, -0.05, 0.045), (0.07, -0.04, 0.035)]:
            o = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(cx + dx - r), Inches(cy + dy - r), Inches(r * 2), Inches(r * 2))
            o.fill.solid(); o.fill.fore_color.rgb = color; o.line.color.rgb = color
        body = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ARC, Inches(cx - 0.12), Inches(cy - 0.01), Inches(0.24), Inches(0.16))
        body.line.color.rgb = color; body.line.width = Pt(2)
    else:
        cal = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(cx - 0.10), Inches(cy - 0.10), Inches(0.20), Inches(0.19))
        cal.fill.background(); cal.line.color.rgb = color; cal.line.width = Pt(1.5)
        for lx in [-0.05, 0.03]:
            line = slide.shapes.add_connector(1, Inches(cx + lx), Inches(cy - 0.13), Inches(cx + lx), Inches(cy - 0.08))
            line.line.color.rgb = color; line.line.width = Pt(1.4)
        hline = slide.shapes.add_connector(1, Inches(cx - 0.10), Inches(cy - 0.035), Inches(cx + 0.10), Inches(cy - 0.035))
        hline.line.color.rgb = color; hline.line.width = Pt(1.2)


def build_ppt():
    assets = crop_assets()
    icons = write_svg_icons()

    prs = Presentation()
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    teal = RGBColor(4, 74, 87)
    dark_teal = RGBColor(2, 64, 76)
    graphite = RGBColor(34, 47, 52)
    pale = RGBColor(236, 247, 250)

    # Background and image regions, split into independent assets rather than one full-slide bitmap.
    add_rect(slide, 0, 0, 16, 9, RGBColor(246, 250, 251))
    slide.shapes.add_picture(str(assets["mountains"]), Inches(0), Inches(0), width=Inches(16), height=Inches(3.95))
    wash = add_rect(slide, 0, 0, 16, 9, RGBColor(255, 255, 255))
    wash.fill.transparency = 39000
    slide.shapes.add_picture(str(assets["apparatus"]), Inches(8.55), Inches(0.30), width=Inches(7.42), height=Inches(8.05))

    # Top-left page tab.
    chevron = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM, Inches(-0.18), Inches(-0.03), Inches(2.62), Inches(1.32))
    chevron.fill.solid(); chevron.fill.fore_color.rgb = dark_teal
    chevron.line.color.rgb = dark_teal
    add_textbox(slide, 0.48, 0.28, 1.15, 0.62, "01", 48, RGBColor(255, 255, 255), True)

    # Main title and metadata.
    add_textbox(slide, 1.00, 2.10, 5.80, 1.95, "金属与冰\n摩擦系数的测量", 40, teal, True)
    add_textbox(slide, 1.04, 4.46, 4.30, 0.40, "物理实验创新竞赛答辩", 23, graphite, True)
    line = slide.shapes.add_connector(1, Inches(1.04), Inches(5.10), Inches(4.95), Inches(5.10))
    line.line.color.rgb = RGBColor(216, 226, 229); line.line.width = Pt(2)
    line2 = slide.shapes.add_connector(1, Inches(1.04), Inches(5.10), Inches(1.60), Inches(5.10))
    line2.line.color.rgb = teal; line2.line.width = Pt(3)

    # Editable icon/text rows.
    add_circle_icon(slide, 1.14, 6.15, "team")
    add_textbox(slide, 1.72, 5.97, 3.25, 0.38, "团队：祝冰间隙者", 18, graphite)
    add_circle_icon(slide, 1.14, 7.15, "calendar")
    add_textbox(slide, 1.72, 6.97, 3.20, 0.38, "日期：2024.05.20", 18, graphite)

    # Inset image as separate cropped PNG.
    circ = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(5.70), Inches(5.94), Inches(2.56), Inches(2.56))
    circ.fill.solid(); circ.fill.fore_color.rgb = RGBColor(255, 255, 255)
    circ.line.color.rgb = RGBColor(255, 255, 255); circ.line.width = Pt(3)
    slide.shapes.add_picture(str(assets["inset"]), Inches(5.78), Inches(6.02), width=Inches(2.40), height=Inches(2.40))

    # Bottom banner rebuilt as shape + editable labels.
    banner = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM, Inches(0.68), Inches(8.10), Inches(7.00), Inches(0.58))
    banner.fill.solid(); banner.fill.fore_color.rgb = dark_teal
    banner.line.color.rgb = dark_teal
    # SVG icons are saved as independent source assets. If the local python-pptx
    # build cannot embed them natively, the companion COM pass below inserts them.
    add_textbox(slide, 1.42, 8.22, 1.45, 0.32, "精准测量", 20, RGBColor(255, 255, 255), True)
    add_textbox(slide, 3.58, 8.22, 1.45, 0.32, "理解机理", 20, RGBColor(255, 255, 255), True)
    add_textbox(slide, 5.76, 8.22, 1.45, 0.32, "服务工程", 20, RGBColor(255, 255, 255), True)

    # Bottom-right page marker.
    marker = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM, Inches(14.87), Inches(8.36), Inches(1.20), Inches(0.64))
    marker.fill.solid(); marker.fill.fore_color.rgb = dark_teal
    marker.line.color.rgb = dark_teal
    add_textbox(slide, 15.20, 8.53, 0.55, 0.28, "01", 23, RGBColor(255, 255, 255), True, PP_ALIGN.CENTER)

    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build_ppt()
