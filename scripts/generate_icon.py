#!/usr/bin/env python3
"""Generate installer/sobha.ico from sobha_logo_brand.png when Pillow is available."""
from __future__ import annotations

from pathlib import Path
import sys


def main() -> int:
    src = Path("sobha_logo_brand.png")
    dst = Path("installer/sobha.ico")

    if not src.exists():
        print(f"Skip icon: {src} not found")
        return 0

    try:
        from PIL import Image
    except ImportError:
        print("Skip icon: Pillow not installed")
        return 0

    dst.parent.mkdir(parents=True, exist_ok=True)
    img = Image.open(src).convert("RGBA")
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    img.save(dst, format="ICO", sizes=sizes)
    print(f"Icon written: {dst} ({dst.stat().st_size // 1024} KB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
