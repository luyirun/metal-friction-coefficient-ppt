from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
SRC_DIR = ROOT / "outputs" / "single_pages"
ASSET_DIR = ROOT / "outputs" / "layered_ppt_assets_99fid"
OUT = ROOT / "outputs" / "metal_ice_99fid_visual_first.pptx"

SLIDE_W_IN, SLIDE_H_IN = 16, 9
COLS, ROWS = 8, 6
APPLY_VISIBLE_PAGE1_TEXT_FIX = False


SLIDE_TEXT = {
    1: {
        "title": "金属与冰摩擦系数的测量",
        "subtitle": "青少年科学素养大赛 · 实验答辩",
        "footer": "参赛选手答辩PPT",
    },
    2: {"title": "选题背景", "point": "极地与低温工程场景中，金属-冰界面摩擦是安全与效率的关键变量。"},
    3: {"title": "题目解读", "point": "围绕“金属与冰”的接触状态、受力关系与可测量变量建立实验路径。"},
    4: {"title": "方法选择 —— 动态斜面法", "point": "通过斜面角度和运动加速度反推动摩擦系数，兼顾可操作性与数据稳定性。"},
    5: {"title": "装置设计", "point": "斜面模块与压力加载模块共同覆盖不同接触压力和运动状态。"},
    6: {"title": "变量设计", "point": "控制温度、接触面积、法向压力等变量，保证结果可比较。"},
    7: {"title": "问题改进", "point": "针对冰面融化、读数误差和压力波动提出装置与流程优化。"},
    8: {"title": "数据处理", "point": "用重复测量、均值处理和误差分析提高结论可信度。"},
    9: {"title": "结果分析", "point": "比较不同条件下摩擦系数变化，提炼影响趋势与实验解释。"},
    10: {"title": "总结展望", "point": "实验方案可拓展到低温运输、冰雪装备和材料接触性能评估。"},
}


def add_sliced_page(slide, page_no: int) -> None:
    src = SRC_DIR / f"page_{page_no:02d}_industrial_whitepaper.png"
    im = Image.open(src).convert("RGB")
    w, h = im.size
    tile_w, tile_h = w // COLS, h // ROWS

    for r in range(ROWS):
        for c in range(COLS):
            left = c * tile_w
            top = r * tile_h
            right = w if c == COLS - 1 else (c + 1) * tile_w
            bottom = h if r == ROWS - 1 else (r + 1) * tile_h
            tile = im.crop((left, top, right, bottom))
            tile_path = ASSET_DIR / f"p{page_no:02d}_tile_{r:02d}_{c:02d}.png"
            tile.save(tile_path, optimize=True)

            x = SLIDE_W_IN * left / w
            y = SLIDE_H_IN * top / h
            tw = SLIDE_W_IN * (right - left) / w
            th = SLIDE_H_IN * (bottom - top) / h
            slide.shapes.add_picture(
                str(tile_path), Inches(x), Inches(y), width=Inches(tw), height=Inches(th)
            )


def add_hidden_editable_text(slide, page_no: int) -> None:
    data = SLIDE_TEXT[page_no]
    text = "\n".join([data.get("title", ""), data.get("point", ""), data.get("subtitle", ""), data.get("footer", "")]).strip()
    box = slide.shapes.add_textbox(Inches(16.25), Inches(0.25), Inches(5.5), Inches(1.2))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.name = "微软雅黑"
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(255, 255, 255)
    box.name = f"editable_text_source_slide_{page_no:02d}"
    box.line.fill.background()
    box.fill.background()


def add_page1_answer_context(slide) -> None:
    # The generated reference read too commercial. Mask only the small subtitle
    # band and replace it with visible editable answer-defense wording.
    mask = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.53), Inches(4.05), Inches(5.25), Inches(1.05))
    mask.fill.solid()
    mask.fill.fore_color.rgb = RGBColor(246, 249, 248)
    mask.line.fill.background()

    subtitle = slide.shapes.add_textbox(Inches(0.58), Inches(4.10), Inches(5.10), Inches(0.42))
    tf = subtitle.text_frame
    tf.clear()
    run = tf.paragraphs[0].add_run()
    run.text = "青少年科学素养大赛 · 实验答辩"
    run.font.name = "微软雅黑"
    run.font.size = Pt(17)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 74, 87)

    caption = slide.shapes.add_textbox(Inches(0.58), Inches(4.56), Inches(5.15), Inches(0.34))
    tf = caption.text_frame
    tf.clear()
    run = tf.paragraphs[0].add_run()
    run.text = "参赛选手：金属与冰界面摩擦实验方案汇报"
    run.font.name = "微软雅黑"
    run.font.size = Pt(11.5)
    run.font.color.rgb = RGBColor(67, 85, 93)

    team_mask = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.76), Inches(5.40), Inches(4.35), Inches(0.82))
    team_mask.fill.solid()
    team_mask.fill.fore_color.rgb = RGBColor(246, 249, 248)
    team_mask.line.fill.background()

    team = slide.shapes.add_textbox(Inches(1.15), Inches(5.67), Inches(3.60), Inches(0.28))
    tf = team.text_frame
    tf.clear()
    run = tf.paragraphs[0].add_run()
    run.text = "参赛选手：实验项目组"
    run.font.name = "微软雅黑"
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor(35, 64, 72)

    strip = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(0.70), Inches(8.34), Inches(6.95), Inches(0.55))
    strip.fill.solid()
    strip.fill.fore_color.rgb = RGBColor(0, 76, 89)
    strip.line.fill.background()

    strip_text = slide.shapes.add_textbox(Inches(0.92), Inches(8.48), Inches(5.70), Inches(0.22))
    tf = strip_text.text_frame
    tf.clear()
    run = tf.paragraphs[0].add_run()
    run.text = "实验测量 · 误差分析 · 答辩汇报"
    run.font.name = "微软雅黑"
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)


def build() -> None:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)

    for page_no in range(1, 11):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_sliced_page(slide, page_no)
        if page_no == 1 and APPLY_VISIBLE_PAGE1_TEXT_FIX:
            add_page1_answer_context(slide)
        add_hidden_editable_text(slide, page_no)

    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
