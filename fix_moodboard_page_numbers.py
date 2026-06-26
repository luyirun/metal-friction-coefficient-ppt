from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

FONT = ImageFont.truetype(r"C:\Windows\Fonts\Noto Sans SC Bold (TrueType).otf", 15)


def patch_label(path, out, patches, fill=(239, 247, 250), text_fill=(14, 58, 88)):
    im = Image.open(path).convert("RGB")
    d = ImageDraw.Draw(im)
    for x, y, label in patches:
        d.rectangle((x, y, x + 48, y + 20), fill=fill)
        d.text((x + 4, y + 1), label, font=FONT, fill=text_fill)
    im.save(out)
    print(out)


root = Path(r"F:\Users\user\Documents\测金属摩擦系数\outputs")

patch_label(
    root / "imagegen_ppt_collage_04_premium_polar_keynote.png",
    root / "imagegen_ppt_collage_04_premium_polar_keynote_corrected.png",
    [
        (455, 585, "03/10"),
        (455, 893, "05/10"),
    ],
)

patch_label(
    root / "imagegen_ppt_collage_05_industrial_whitepaper.png",
    root / "imagegen_ppt_collage_05_industrial_whitepaper_corrected.png",
    [
        (455, 1198, "07"),
    ],
    fill=(248, 251, 251),
    text_fill=(0, 67, 82),
)
