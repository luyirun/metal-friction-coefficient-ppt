from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
PREVIEW_DIR = ROOT / "outputs" / "metal_ice_whitepaper_direct_refined_v2_preview"
OUT = ROOT / "outputs" / "metal_ice_whitepaper_direct_refined_v2_preview_sheet.jpg"


def slide_no(path: Path) -> int:
    digits = "".join(ch for ch in path.stem if ch.isdigit())
    return int(digits or 0)


files = sorted(PREVIEW_DIR.glob("*.PNG"), key=slide_no)
thumb_w, thumb_h = 512, 288
gap = 26
label_h = 34
cols, rows = 2, 5
sheet = Image.new("RGB", (cols * thumb_w + (cols + 1) * gap, rows * (thumb_h + label_h) + (rows + 1) * gap), "#eaf1f2")
draw = ImageDraw.Draw(sheet)
try:
    font = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 20)
except OSError:
    font = None

for idx, path in enumerate(files):
    im = Image.open(path).convert("RGB").resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    row, col = divmod(idx, cols)
    x = gap + col * (thumb_w + gap)
    y = gap + row * (thumb_h + label_h + gap)
    draw.text((x, y), f"Slide {idx + 1:02d}", fill="#163d45", font=font)
    sheet.paste(im, (x, y + label_h))
    draw.rectangle((x, y + label_h, x + thumb_w, y + label_h + thumb_h), outline="#aabcc0", width=1)

sheet.save(OUT, quality=92)
print(OUT)
