"""Render ビルド時に日本語フォント（BIZ UDGothic）をダウンロードするスクリプト"""
import urllib.request
from pathlib import Path

FONT_DIR = Path(__file__).parent / "fonts"
FONT_DIR.mkdir(exist_ok=True)

BASE  = "https://github.com/google/fonts/raw/main/ofl/bizudgothic/"
FILES = ["BIZUDGothic-Regular.ttf", "BIZUDGothic-Bold.ttf"]

for fname in FILES:
    dest = FONT_DIR / fname
    if dest.exists():
        print(f"Already exists: {fname} ({dest.stat().st_size:,} bytes)")
        continue
    print(f"Downloading {fname} ...", flush=True)
    try:
        urllib.request.urlretrieve(BASE + fname, dest)
        print(f"  -> {dest.stat().st_size:,} bytes", flush=True)
    except Exception as e:
        print(f"  ERROR: {e}", flush=True)
