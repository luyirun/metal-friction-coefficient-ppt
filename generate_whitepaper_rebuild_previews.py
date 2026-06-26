import base64
import os
from pathlib import Path

import httpx
from openai import OpenAI


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "outputs" / "whitepaper_preview_rebuild"
PROMPTS = OUT / "prompts"
REFS = OUT / "reference_crops"


def generate_page(page: int) -> Path:
    image_key = os.environ.get("OPENAI_IMAGE_API_KEY")
    if image_key:
        os.environ["OPENAI_API_KEY"] = image_key

    prompt = (PROMPTS / f"page_{page:02d}_prompt.txt").read_text(encoding="utf-8")
    out = OUT / f"page_{page:02d}_preview.png"
    image_path = REFS / f"page_{page:02d}_reference.png"

    client = OpenAI()
    with image_path.open("rb") as image_file:
        result = client.images.edit(
            model="gpt-image-2",
            image=image_file,
            prompt=prompt,
            quality="high",
            size="1536x864",
        )

    item = result.data[0]
    b64 = getattr(item, "b64_json", None)
    url = getattr(item, "url", None)
    if b64:
        out.write_bytes(base64.b64decode(b64))
    elif url:
        with httpx.Client(timeout=180) as http:
            response = http.get(url)
            response.raise_for_status()
            out.write_bytes(response.content)
    else:
        raise RuntimeError(f"Image response had neither b64_json nor url: {item!r}")
    return out


def main() -> None:
    pages = [p for p in range(1, 11) if not (OUT / f"page_{p:02d}_preview.png").exists()]
    if not pages:
        print("All previews already exist.")
        return
    for page in pages:
        print(f"Generating page {page:02d}...", flush=True)
        out = generate_page(page)
        print(f"Done page {page:02d}: {out}", flush=True)


if __name__ == "__main__":
    main()
