import argparse
import base64
from pathlib import Path
import time

import httpx
from openai import OpenAI


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["generate", "edit"], required=True)
    parser.add_argument("--prompt-file", required=True)
    parser.add_argument("--image")
    parser.add_argument("--out", required=True)
    parser.add_argument("--model", default="gpt-image-2")
    parser.add_argument("--quality", default="high")
    parser.add_argument("--size", default="1024x1536")
    args = parser.parse_args()

    prompt = Path(args.prompt_file).read_text(encoding="utf-8")
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    client = OpenAI()

    payload = {
        "model": args.model,
        "prompt": prompt,
        "quality": args.quality,
        "size": args.size,
    }

    if args.mode == "edit":
        if not args.image:
            raise SystemExit("--image is required for edit mode")
        with Path(args.image).open("rb") as f:
            result = client.images.edit(image=f, **payload)
    else:
        result = client.images.generate(**payload)

    item = result.data[0]
    b64 = getattr(item, "b64_json", None)
    url = getattr(item, "url", None)

    if b64:
        out.write_bytes(base64.b64decode(b64))
    elif url:
        last_error = None
        for attempt in range(1, 5):
            try:
                with httpx.Client(timeout=180) as http:
                    response = http.get(url)
                    response.raise_for_status()
                    out.write_bytes(response.content)
                break
            except httpx.HTTPError as exc:
                last_error = exc
                if attempt == 4:
                    raise
                time.sleep(5 * attempt)
    else:
        raise RuntimeError(f"Image response had neither b64_json nor url: {item!r}")

    print(out)


if __name__ == "__main__":
    main()
