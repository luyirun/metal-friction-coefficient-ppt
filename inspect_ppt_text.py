import json
from pathlib import Path

data = json.loads(Path(r"F:\Users\user\Documents\测金属摩擦系数\materials_reading_summary.json").read_text(encoding="utf-8"))

for f in data["files"]:
    if f["suffix"] == ".pptx":
        print("###", f["relative_path"], f["content"]["slide_count"])
        for s in f["content"]["slides"]:
            text = " / ".join(s["text"].split())
            print(f"SLIDE {s['slide']}: {text[:1400]}")
        print()
