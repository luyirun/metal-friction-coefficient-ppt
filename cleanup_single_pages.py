from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数\outputs\keynote_single_pages")
FONT_REG = ImageFont.truetype(r"C:\Windows\Fonts\Noto Sans SC (TrueType).otf", 25)
FONT_MED = ImageFont.truetype(r"C:\Windows\Fonts\Noto Sans SC Medium (TrueType).otf", 28)


def cleanup_slide_01():
    path = ROOT / "slide_01.png"
    im = Image.open(path).convert("RGBA")
    overlay = Image.new("RGBA", im.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)

    # Cover hallucinated metadata in the lower-left body area with a soft frosted panel.
    panel = (80, 708, 635, 892)
    d.rounded_rectangle(panel, radius=18, fill=(238, 246, 251, 232), outline=(189, 211, 224, 180), width=1)
    d.text((118, 738), "研究要点", font=FONT_MED, fill=(7, 58, 98, 255))
    items = ["金属-冰接触面", "实验装置", "规律分析"]
    y = 788
    for item in items:
        d.ellipse((122, y + 8, 136, y + 22), fill=(0, 96, 145, 255))
        d.text((154, y), item, font=FONT_REG, fill=(20, 51, 78, 255))
        y += 42

    cleaned = Image.alpha_composite(im, overlay).convert("RGB")
    cleaned.save(path)
    print(path)


cleanup_slide_01()
