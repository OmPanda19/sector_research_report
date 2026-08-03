#!/usr/bin/env python3
"""Download a PDF and dump its text to stdout (and cache under research/raw/)."""
import hashlib
import os
import subprocess
import sys

RAW = os.path.join(os.path.dirname(os.path.abspath(__file__)), "raw")
os.makedirs(RAW, exist_ok=True)

UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"


def fetch(url: str) -> str:
    name = hashlib.sha1(url.encode()).hexdigest()[:16] + ".pdf"
    path = os.path.join(RAW, name)
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        subprocess.run(
            ["curl", "-sSLk", "--max-time", "180", "-A", UA, "-o", path, url],
            check=True,
        )
    return path


def text(path: str, first: int = 0, last: int = 0) -> str:
    from pypdf import PdfReader

    reader = PdfReader(path)
    pages = reader.pages
    lo = first if first else 0
    hi = last if last else len(pages)
    out = [f"### TOTAL PAGES: {len(pages)}"]
    for i in range(lo, min(hi, len(pages))):
        out.append(f"\n===== PAGE {i + 1} =====")
        try:
            out.append(pages[i].extract_text() or "")
        except Exception as exc:  # noqa: BLE001
            out.append(f"[extract error: {exc}]")
    return "\n".join(out)


if __name__ == "__main__":
    url = sys.argv[1]
    first = int(sys.argv[2]) if len(sys.argv) > 2 else 0
    last = int(sys.argv[3]) if len(sys.argv) > 3 else 0
    print(text(fetch(url), first, last))
