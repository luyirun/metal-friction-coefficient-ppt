from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
SRC = ROOT / "outputs" / "imagegen_ppt_collage_05_industrial_whitepaper_corrected.png"
OUT = ROOT / "outputs" / "moodboard05_page_refs"
OUT.mkdir(parents=True, exist_ok=True)

im = Image.open(SRC).convert("RGB")

# Coordinates measured from the selected 1024x1536 moodboard.
# They intentionally include the full slide thumbnail border but exclude inter-slide gutters.
boxes = [
    (11, 10, 509, 315),    (516, 10, 1014, 315),
    (11, 317, 509, 622),   (516, 317, 1014, 622),
    (11, 625, 509, 929),   (516, 625, 1014, 929),
    (11, 932, 509, 1235),  (516, 932, 1014, 1235),
    (11, 1238, 509, 1531), (516, 1238, 1014, 1531),
]

for i, box in enumerate(boxes, 1):
    crop = im.crop(box)
    # Normalize to an exact 16:9 reference canvas for the image edit model.
    crop = crop.resize((1280, 720), Image.Resampling.LANCZOS)
    path = OUT / f"page_{i:02d}_ref.png"
    crop.save(path)
    print(path)
