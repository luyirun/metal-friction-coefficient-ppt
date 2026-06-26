from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
PREVIEW_DIR = ROOT / "outputs" / "metal_ice_99fid_layered_editable_preview"
OUT = ROOT / "outputs" / "metal_ice_99fid_layered_editable_preview_sheet.jpg"

def slide_no(path: Path) -> int:
    stem = path.stem.replace("幻灯片", "")
    return int(stem)


files = sorted(PREVIEW_DIR.glob("*.PNG"), key=slide_no)
thumb_w, thumb_h = 512, 288
gap = 26
label_h = 34
cols = 2
rows = 5

sheet = Image.new("RGB", (cols * thumb_w + (cols + 1) * gap, rows * (thumb_h + label_h) + (rows + 1) * gap), "#eef3f4")
draw = ImageDraw.Draw(sheet)

for idx, path in enumerate(files):
    im = Image.open(path).convert("RGB").resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    row, col = divmod(idx, cols)
    x = gap + col * (thumb_w + gap)
    y = gap + row * (thumb_h + label_h + gap)
    sheet.paste(im, (x, y + label_h))
    draw.text((x, y), f"Slide {idx + 1:02d}", fill="#163d45")

sheet.save(OUT, quality=92)
print(OUT)
