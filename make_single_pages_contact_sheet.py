from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
SRC = ROOT / "outputs" / "single_pages"
OUT = ROOT / "outputs" / "single_pages_contact_sheet.jpg"

paths = [SRC / f"page_{i:02d}_industrial_whitepaper.png" for i in range(1, 11)]
thumb_w, thumb_h = 640, 360
gap = 26
margin = 36
header = 72
cols = 2
rows = 5

sheet = Image.new("RGB", (cols * thumb_w + gap + margin * 2, rows * thumb_h + gap * (rows - 1) + margin * 2 + header), (238, 244, 245))
d = ImageDraw.Draw(sheet)
font = ImageFont.truetype(r"C:\Windows\Fonts\Noto Sans SC Bold (TrueType).otf", 34)
d.text((margin, 24), "工业白皮书风 - 单页高清视觉稿总览", font=font, fill=(18, 48, 58))

for idx, path in enumerate(paths):
    im = Image.open(path).convert("RGB").resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    x = margin + (idx % cols) * (thumb_w + gap)
    y = margin + header + (idx // cols) * (thumb_h + gap)
    sheet.paste(im, (x, y))
    d.rectangle((x, y, x + thumb_w, y + thumb_h), outline=(190, 205, 210), width=2)

sheet.save(OUT, quality=95)
print(OUT)
