from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "outputs"
PREVIEW_DIR = OUT_DIR / "metal_ice_keynote_style_samples_v1_preview"
OUT = OUT_DIR / "metal_ice_keynote_style_samples_v1_preview_contact.png"


def slide_no(path: Path) -> int:
    digits = "".join(ch for ch in path.stem if ch.isdigit())
    return int(digits) if digits else 0


files = sorted(PREVIEW_DIR.glob("*.PNG"), key=slide_no)
thumb_w, thumb_h = 640, 360
gap = 24
label_h = 30
cols = 1
rows = len(files)

sheet = Image.new("RGB", (thumb_w + gap * 2, rows * (thumb_h + label_h) + gap * (rows + 1)), "#eaf2f8")
draw = ImageDraw.Draw(sheet)
font_path = Path(r"C:\Windows\Fonts\msyhbd.ttc")
font = ImageFont.truetype(str(font_path), 18) if font_path.exists() else ImageFont.load_default()

for idx, path in enumerate(files):
    im = Image.open(path).convert("RGB").resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    x = gap
    y = gap + idx * (thumb_h + label_h + gap)
    draw.text((x, y), f"Sample {idx + 1:02d}", font=font, fill="#083b66")
    sheet.paste(im, (x, y + label_h))
    draw.rectangle((x, y + label_h, x + thumb_w, y + label_h + thumb_h), outline="#9bb4c2", width=2)

sheet.save(OUT)
print(OUT)
