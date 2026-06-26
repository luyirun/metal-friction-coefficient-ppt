from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数\outputs")
paths = sorted(ROOT.glob("pptx_v2_preview_page-*.png"))
thumb_w, thumb_h = 384, 216
gap = 20
margin = 30
header = 55
cols = 2

sheet = Image.new("RGB", (cols * thumb_w + gap + margin * 2, 5 * thumb_h + gap * 4 + margin * 2 + header), (228, 238, 242))
d = ImageDraw.Draw(sheet)
font = ImageFont.truetype(r"C:\Windows\Fonts\Noto Sans SC Bold (TrueType).otf", 26)
d.text((margin, 18), "PPTX v2 PowerPoint 导出预览", font=font, fill=(7, 50, 77))

for i, p in enumerate(paths):
    im = Image.open(p).convert("RGB").resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    x = margin + (i % cols) * (thumb_w + gap)
    y = margin + header + (i // cols) * (thumb_h + gap)
    sheet.paste(im, (x, y))
    d.rectangle((x, y, x + thumb_w, y + thumb_h), outline=(141, 164, 176), width=2)

out = ROOT / "metal_ice_keynote_editable_v2_preview_contact.png"
sheet.save(out)
print(out)
