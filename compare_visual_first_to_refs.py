from pathlib import Path

from PIL import Image, ImageChops, ImageStat


ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
REF_DIR = ROOT / "outputs" / "single_pages"
PREVIEW_DIR = ROOT / "outputs" / "metal_ice_99fid_visual_first_preview"


def rmsdiff(a: Image.Image, b: Image.Image) -> float:
    diff = ImageChops.difference(a, b)
    stat = ImageStat.Stat(diff)
    return sum(v ** 2 for v in stat.rms) ** 0.5


for i in range(1, 11):
    ref = Image.open(REF_DIR / f"page_{i:02d}_industrial_whitepaper.png").convert("RGB")
    prev = Image.open(PREVIEW_DIR / f"幻灯片{i}.PNG").convert("RGB")
    if prev.size != ref.size:
        prev = prev.resize(ref.size, Image.Resampling.LANCZOS)
    score = rmsdiff(ref, prev)
    similarity = max(0, 100 - score / 441.67295593 * 100)
    print(f"slide {i:02d}: rms={score:.3f}, approx_similarity={similarity:.2f}%")
