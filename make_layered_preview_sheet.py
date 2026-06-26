from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
SRC = ROOT / "outputs" / "metal_ice_high_fidelity_layered_preview"
OUT = ROOT / "outputs" / "metal_ice_high_fidelity_layered_preview_sheet.jpg"

paths = [SRC / f"幻灯片{i}.PNG" for i in range(1, 11)]
thumb_w, thumb_h = 640, 360
gap = 24
margin = 36
header = 70
sheet = Image.new("RGB", (2 * thumb_w + gap + margin * 2, 5 * thumb_h + 4 * gap + margin * 2 + header), (238, 244, 245))
d = ImageDraw.Draw(sheet)
font = ImageFont.truetype(r"C:\Windows\Fonts\msyhbd.ttc", 32)
d.text((margin, 24), "高保真分层 PPTX 导出预览", font=font, fill=(18, 48, 58))
for idx, p in enumerate(paths):
    im = Image.open(p).convert("RGB").resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    x = margin + (idx % 2) * (thumb_w + gap)
    y = margin + header + (idx // 2) * (thumb_h + gap)
    sheet.paste(im, (x, y))
    d.rectangle((x, y, x + thumb_w, y + thumb_h), outline=(190, 205, 210), width=2)
sheet.save(OUT, quality=95)
print(OUT)
