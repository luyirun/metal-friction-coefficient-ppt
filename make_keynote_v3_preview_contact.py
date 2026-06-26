from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "outputs"
PREVIEW_DIR = OUT_DIR / "metal_ice_keynote_editable_v3_preview"
OUT = OUT_DIR / "metal_ice_keynote_editable_v3_preview_contact.png"


def slide_no(path: Path) -> int:
    digits = "".join(ch for ch in path.stem if ch.isdigit())
    return int(digits) if digits else 0


files = sorted(PREVIEW_DIR.glob("*.PNG"), key=slide_no)
thumb_w, thumb_h = 512, 288
gap = 26
label_h = 34
cols, rows = 2, 5

sheet = Image.new(
    "RGB",
    (cols * thumb_w + (cols + 1) * gap, rows * (thumb_h + label_h) + (rows + 1) * gap),
    "#eaf2f5",
)
draw = ImageDraw.Draw(sheet)
font_path = Path(r"C:\Windows\Fonts\msyhbd.ttc")
font = ImageFont.truetype(str(font_path), 18) if font_path.exists() else ImageFont.load_default()

for idx, path in enumerate(files):
    im = Image.open(path).convert("RGB").resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    row, col = divmod(idx, cols)
    x = gap + col * (thumb_w + gap)
    y = gap + row * (thumb_h + label_h + gap)
    draw.text((x, y), f"Slide {idx + 1:02d}", font=font, fill="#053663")
    sheet.paste(im, (x, y + label_h))
    draw.rectangle((x, y + label_h, x + thumb_w, y + label_h + thumb_h), outline="#8ca5b2", width=2)

sheet.save(OUT)
print(OUT)
