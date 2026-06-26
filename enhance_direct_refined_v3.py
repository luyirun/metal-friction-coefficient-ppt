from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "outputs" / "metal_ice_whitepaper_direct_refined_v2.pptx"
OUT = ROOT / "outputs" / "metal_ice_whitepaper_direct_refined_v4.pptx"
FONT = "Microsoft YaHei"

TEAL = RGBColor(0, 72, 84)
DARK = RGBColor(28, 44, 49)
MUTED = RGBColor(78, 98, 104)
LINE = RGBColor(211, 225, 229)
ICE = RGBColor(226, 243, 247)
ORANGE = RGBColor(226, 126, 39)
GREEN = RGBColor(61, 137, 78)
WHITE = RGBColor(255, 255, 255)
SOFT = RGBColor(247, 250, 250)


def text(slide, x, y, w, h, value, size, color=DARK, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
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


def rect(slide, x, y, w, h, fill=WHITE, line=None, radius=True, transparency=0):
    shp = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h),
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


def line(slide, x1, y1, x2, y2, color=LINE, width=1):
    shp = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    shp.line.color.rgb = color
    shp.line.width = Pt(width)
    return shp


def remove_shapes(slide, predicate):
    for shp in list(slide.shapes):
        try:
            if predicate(shp):
                shp._element.getparent().remove(shp._element)
        except Exception:
            continue


def poly(slide, pts, color=TEAL, width=1.2):
    for a, b in zip(pts, pts[1:]):
        line(slide, a[0], a[1], b[0], b[1], color, width)


def bullet(slide, x, y, value, size=9.8, color=DARK):
    rect(slide, x, y + 0.08, 0.045, 0.045, TEAL, radius=True)
    text(slide, x + 0.14, y, 2.6, 0.18, value, size, color)


def cube_icon(slide, cx, cy):
    # Simple editable line icon drawn over the existing icon circle.
    line(slide, cx - 0.22, cy - 0.05, cx, cy - 0.18, TEAL, 1)
    line(slide, cx, cy - 0.18, cx + 0.22, cy - 0.05, TEAL, 1)
    line(slide, cx - 0.22, cy - 0.05, cx - 0.22, cy + 0.20, TEAL, 1)
    line(slide, cx + 0.22, cy - 0.05, cx + 0.22, cy + 0.20, TEAL, 1)
    line(slide, cx - 0.22, cy + 0.20, cx, cy + 0.34, TEAL, 1)
    line(slide, cx, cy + 0.34, cx + 0.22, cy + 0.20, TEAL, 1)
    line(slide, cx, cy - 0.18, cx, cy + 0.34, TEAL, 1)


def tools_icon(slide, cx, cy):
    line(slide, cx - 0.25, cy - 0.20, cx + 0.24, cy + 0.28, TEAL, 1.3)
    line(slide, cx + 0.25, cy - 0.20, cx - 0.24, cy + 0.28, TEAL, 1.3)
    rect(slide, cx - 0.29, cy + 0.22, 0.10, 0.10, ICE, TEAL, True)
    rect(slide, cx + 0.19, cy + 0.22, 0.10, 0.10, ICE, TEAL, True)


def chart_icon(slide, cx, cy):
    line(slide, cx - 0.27, cy + 0.25, cx - 0.27, cy - 0.20, TEAL, 1)
    line(slide, cx - 0.27, cy + 0.25, cx + 0.28, cy + 0.25, TEAL, 1)
    poly(slide, [(cx - 0.18, cy + 0.15), (cx - 0.02, cy - 0.02), (cx + 0.10, cy + 0.05), (cx + 0.24, cy - 0.16)], TEAL, 1.3)


def enhance_slide3(slide):
    # Make the three-column page look less like empty cards and closer to the original technical columns.
    centers = [(3.08, 2.98), (8.03, 2.98), (12.98, 2.98)]
    cube_icon(slide, *centers[0])
    tools_icon(slide, *centers[1])
    chart_icon(slide, *centers[2])
    labels = ["接触界面", "实验平台", "数据闭环"]
    for (cx, cy), label in zip(centers, labels):
        rect(slide, cx - 0.62, cy + 0.72, 1.24, 0.30, SOFT, LINE, True)
        text(slide, cx - 0.54, cy + 0.80, 1.08, 0.12, label, 8.8, TEAL, True, PP_ALIGN.CENTER)
    remove_shapes(
        slide,
        lambda shp: (
            getattr(shp, "has_text_frame", False)
            and shp.has_text_frame
            and any(t in shp.text for t in ["温度与压力影响", "稳定与校准", "结果讨论与验证", "• "])
        )
        or (
            shp.shape_type == 9
            and shp.top / 914400 >= 6.95
            and shp.left / 914400 <= 14.7
        ),
    )
    remove_shapes(
        slide,
        lambda shp: (
            shp.shape_type == 1
            and shp.width / 914400 <= 0.08
            and shp.height / 914400 <= 0.08
            and shp.top / 914400 >= 5.35
        ),
    )


def enhance_slide6(slide):
    # Add an editable factor matrix baseline and reduce the feeling of empty lower space.
    text(slide, 0.98, 6.18, 1.50, 0.18, "变量水平矩阵", 10.5, TEAL, True)
    cols = [
        ("温度", "-5 / -10 / -15 ℃"),
        ("压强", "0.05 / 0.10 / 0.20 MPa"),
        ("材质", "铝合金 / 不锈钢 / 钛合金"),
        ("粗糙度", "Ra 0.2 / 1.0 / 3.0 μm"),
        ("面积", "1 / 4 / 9 cm²"),
        ("成分", "淡水冰 / 海冰 / 掺盐冰"),
    ]
    for i, (h, v) in enumerate(cols):
        x = 0.96 + i * 2.38
        rect(slide, x, 6.55, 1.70, 0.58, WHITE, LINE, True)
        text(slide, x + 0.12, 6.70, 1.46, 0.12, h, 8.6, TEAL, True, PP_ALIGN.CENTER)
        text(slide, x + 0.10, 6.94, 1.50, 0.12, v, 7.3, MUTED, False, PP_ALIGN.CENTER)
    line(slide, 0.92, 7.45, 15.12, 7.45, LINE, 0.8)
    remove_shapes(
        slide,
        lambda shp: (
            shp.top / 914400 >= 6.15
            and shp.left / 914400 <= 15.3
        )
        or (
            getattr(shp, "has_text_frame", False)
            and shp.has_text_frame
            and shp.text.strip() in {
                "变量水平矩阵",
                "温度",
                "压强",
                "材质",
                "粗糙度",
                "面积",
                "成分",
                "-5 / -10 / -15 ℃",
                "0.05 / 0.10 / 0.20 MPa",
                "铝合金 / 不锈钢 / 钛合金",
                "Ra 0.2 / 1.0 / 3.0 μm",
                "1 / 4 / 9 cm²",
                "淡水冰 / 海冰 / 掺盐冰",
            }
        ),
    )


def enhance_slide8(slide):
    # Add the small source-like visual labels under each step and make the formula panel cleaner.
    for x, label in [(1.28, "原始记录"), (5.02, "曲线处理"), (8.72, "模型代入"), (12.42, "误差合成")]:
        rect(slide, x, 2.95, 1.62, 0.28, SOFT, LINE, True)
        text(slide, x + 0.10, 3.03, 1.42, 0.10, label, 8.5, TEAL, True, PP_ALIGN.CENTER)
    # Add compact bottom captions to visually fill the card bottoms.
    for x, cap in [(1.32, "采样频率 100 Hz"), (5.00, "得到 a(t)"), (8.72, "计算 μ"), (12.38, "扩展不确定度 U")]:
        text(slide, x, 6.22, 1.78, 0.16, cap, 8.8, MUTED, False, PP_ALIGN.CENTER)
    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False) and shp.has_text_frame and "sinθ" in shp.text:
            shp.left = Inches(8.62)
            shp.top = Inches(3.25)
            shp.width = Inches(2.72)
            shp.height = Inches(0.78)
            tf = shp.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.05)
            tf.margin_right = Inches(0.05)
            tf.margin_top = Inches(0.02)
            tf.margin_bottom = Inches(0.02)
            for p in tf.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.size = Pt(12.5)
                    r.font.name = FONT
                    r.font.color.rgb = DARK
            break
    remove_shapes(
        slide,
        lambda shp: (
            getattr(shp, "has_text_frame", False)
            and shp.has_text_frame
            and shp.text.strip() in {
                "原始记录",
                "曲线处理",
                "模型代入",
                "误差合成",
                "采样频率 100 Hz",
                "得到 a(t)",
                "计算 μ",
                "扩展不确定度 U",
            }
        )
        or (
            shp.shape_type == 1
            and abs(shp.top / 914400 - 2.95) < 0.12
            and abs(shp.width / 914400 - 1.62) < 0.15
        )
        or (
            getattr(shp, "has_text_frame", False)
            and shp.has_text_frame
            and "sinθ" in shp.text
            and 8.45 <= shp.left / 914400 <= 9.15
            and 3.15 <= shp.top / 914400 <= 3.75
        ),
    )
    text(slide, 8.90, 3.56, 1.60, 0.18, "μ = (sinθ - a/g) / cosθ", 10.5, DARK, False, PP_ALIGN.CENTER)


def enhance_slide9(slide):
    # Add clearer axis labels and a source-like interpretation divider.
    text(slide, 1.05, 1.70, 0.90, 0.16, "μ", 9.0, MUTED, True, PP_ALIGN.CENTER)
    text(slide, 4.05, 6.88, 1.20, 0.16, "温度 / ℃", 9.0, MUTED, False, PP_ALIGN.CENTER)
    line(slide, 8.95, 1.72, 8.95, 6.40, LINE, 0.8)
    text(slide, 9.28, 1.30, 1.40, 0.18, "数据解读", 12.2, TEAL, True)
    text(slide, 9.28, 5.55, 3.15, 0.30, "结果与文献趋势一致，\n验证方法的有效性", 10.2, TEAL, True, PP_ALIGN.CENTER)
    remove_shapes(
        slide,
        lambda shp: (
            getattr(shp, "has_text_frame", False)
            and shp.has_text_frame
            and shp.text.strip() in {
                "μ",
                "温度 / ℃",
                "数据解读",
                "结果与文献趋势一致，\n验证方法的有效性",
                "结果与文献趋势一致， | 验证方法的有效性",
                "趋势一致性用于验证实验方案可行性",
            }
        )
        or (
            shp.shape_type == 1
            and abs(shp.left / 914400 - 9.1) < 0.25
            and abs(shp.top / 914400 - 6.28) < 0.25
        ),
    )
    rect(slide, 9.55, 5.62, 3.80, 1.02, WHITE, LINE, True)
    text(slide, 9.98, 5.92, 2.95, 0.42, "结果与文献趋势一致，\n验证方法的有效性", 10.8, TEAL, True, PP_ALIGN.CENTER)


def main():
    prs = Presentation(SRC)
    enhance_slide3(prs.slides[2])
    enhance_slide6(prs.slides[5])
    enhance_slide8(prs.slides[7])
    enhance_slide9(prs.slides[8])
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
