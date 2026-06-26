import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(r"F:\桌面\金属与冰的摩擦系数 _")
OUT = Path(r"F:\Users\user\Documents\测金属摩擦系数\materials_reading_summary.json")


def clean_text(text):
    text = re.sub(r"[ \t\r\f\v]+", " ", text or "")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def fix_mojibake(text):
    if not isinstance(text, str):
        return text
    # Some Windows console/codepage paths turn UTF-8 Chinese into GBK-decoded
    # mojibake. This reverses that case while leaving normal text alone.
    suspicious = ("閲", "妗", "鎽", "绯", "锛", "鈥", "銆")
    if not any(ch in text for ch in suspicious):
        return text
    try:
        repaired = text.encode("gbk", errors="strict").decode("utf-8", errors="strict")
        return repaired
    except Exception:
        return text


def pptx_text(path):
    ns = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    slides = []
    media = []
    with zipfile.ZipFile(path) as z:
        slide_names = sorted(
            [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
            key=lambda n: int(re.search(r"slide(\d+)\.xml", n).group(1)),
        )
        media = [n for n in z.namelist() if n.startswith("ppt/media/")]
        for name in slide_names:
            root = ET.fromstring(z.read(name))
            texts = [t.text for t in root.findall(".//a:t", ns) if t.text]
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            slides.append({"slide": slide_no, "text": fix_mojibake(clean_text("\n".join(texts)))})
    return {"slide_count": len(slides), "slides": slides, "media_count": len(media), "media": media[:50]}


def pdf_text(path):
    result = {"pages": None, "text": "", "page_text": []}
    try:
        import pdfplumber

        with pdfplumber.open(path) as pdf:
            result["pages"] = len(pdf.pages)
            for i, page in enumerate(pdf.pages, 1):
                txt = fix_mojibake(clean_text(page.extract_text() or ""))
                result["page_text"].append({"page": i, "text": txt})
            result["text"] = clean_text("\n\n".join(p["text"] for p in result["page_text"]))
        return result
    except Exception as e:
        result["pdfplumber_error"] = repr(e)
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(path))
        result["pages"] = len(reader.pages)
        result["page_text"] = []
        for i, page in enumerate(reader.pages, 1):
            txt = fix_mojibake(clean_text(page.extract_text() or ""))
            result["page_text"].append({"page": i, "text": txt})
        result["text"] = clean_text("\n\n".join(p["text"] for p in result["page_text"]))
    except Exception as e:
        result["pypdf_error"] = repr(e)
    return result


def image_info(path):
    info = {}
    try:
        from PIL import Image

        with Image.open(path) as im:
            info.update({"format": im.format, "width": im.width, "height": im.height, "mode": im.mode})
    except Exception as e:
        info["error"] = repr(e)
    return info


def main():
    files = [p for p in ROOT.rglob("*") if p.is_file()]
    summary = {"root": str(ROOT), "file_count": len(files), "files": []}
    for path in files:
        item = {
            "path": str(path),
            "relative_path": str(path.relative_to(ROOT)),
            "suffix": path.suffix.lower(),
            "size": path.stat().st_size,
        }
        suffix = path.suffix.lower()
        if suffix == ".pptx":
            item["content"] = pptx_text(path)
        elif suffix == ".pdf":
            item["content"] = pdf_text(path)
        elif suffix in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}:
            item["content"] = image_info(path)
        else:
            try:
                item["content"] = {"text": clean_text(path.read_text(encoding="utf-8"))}
            except Exception as e:
                item["content"] = {"error": repr(e)}
        summary["files"].append(item)
        print(f"read {item['relative_path']}", file=sys.stderr)
    OUT.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(OUT))


if __name__ == "__main__":
    main()
