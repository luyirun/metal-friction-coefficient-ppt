import json
from pathlib import Path

SRC = Path(r"F:\Users\user\Documents\测金属摩擦系数\materials_reading_summary.json")
OUT = Path(r"F:\Users\user\Documents\测金属摩擦系数\materials_reading_index.txt")


def one_line(text, limit=260):
    text = " ".join((text or "").split())
    return text[:limit] + ("..." if len(text) > limit else "")


def main():
    d = json.loads(SRC.read_text(encoding="utf-8"))
    lines = []
    lines.append(f"阅读根目录: {d['root']}")
    lines.append(f"文件总数: {d['file_count']}")
    lines.append("")
    for i, f in enumerate(d["files"], 1):
        c = f.get("content", {})
        lines.append(f"{i}. {f['relative_path']}")
        lines.append(f"   类型: {f['suffix']}  大小: {f['size']} bytes")
        if f["suffix"] == ".pptx":
            lines.append(f"   PPT: {c.get('slide_count')} 页, 媒体 {c.get('media_count')} 个")
            for slide in c.get("slides", [])[:5]:
                if slide.get("text"):
                    lines.append(f"   - 第{slide['slide']}页: {one_line(slide['text'], 220)}")
            if c.get("slide_count", 0) > 5:
                lines.append("   - 其余页文字已保存在 JSON 摘要中")
        elif f["suffix"] == ".pdf":
            txt = c.get("text", "")
            lines.append(f"   PDF: {c.get('pages')} 页, 可提取文本 {len(txt)} 字符")
            lines.append(f"   摘要片段: {one_line(txt, 300) if txt else '未提取到文本，可能是扫描版或图片型 PDF'}")
        elif f["suffix"] in (".jpeg", ".jpg", ".png"):
            lines.append(f"   图片: {c.get('width')}x{c.get('height')} {c.get('format')} {c.get('mode')}")
        lines.append("")
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(str(OUT))


if __name__ == "__main__":
    main()
