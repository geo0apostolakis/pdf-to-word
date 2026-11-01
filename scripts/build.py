#!/usr/bin/env python3
"""Inline src/ and vendor/ into one self-contained page: dist/pdf-to-word.html.
One file, no server, works offline. Run: python3 scripts/build.py"""
import re
from pathlib import Path

root = Path(__file__).resolve().parent.parent
page = (root / "src/index.html").read_text()
page = re.sub(r"/\*@include ([\w./-]+)\*/", lambda m: (root / m.group(1)).read_text(), page)
out = root / "dist/pdf-to-word.html"
out.parent.mkdir(exist_ok=True)
out.write_text(page)
print(f"{out.relative_to(root)}  {out.stat().st_size / 1e6:.2f} MB")
