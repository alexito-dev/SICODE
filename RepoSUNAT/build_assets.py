"""Genera logo.png, favicon.png e icon.ico a partir del logo original (webp/png)."""
import sys
from pathlib import Path

from PIL import Image, ImageChops

src = Path(sys.argv[1])
root = Path(__file__).parent
(root / "frontend" / "public").mkdir(parents=True, exist_ok=True)
(root / "assets").mkdir(exist_ok=True)

im = Image.open(src).convert("RGBA")
bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
mask = ImageChops.difference(im, bg).convert("L").point(lambda p: 255 if p > 12 else 0)
l, t, r, b = mask.getbbox()
half = max(r - l, b - t) // 2
half += int(half * 0.08)
cx, cy = (l + r) // 2, (t + b) // 2
logo = im.crop((cx - half, cy - half, cx + half, cy + half)).resize((512, 512), Image.LANCZOS)

logo.save(root / "frontend" / "public" / "logo.png")
logo.save(root / "frontend" / "public" / "favicon.png")
logo.save(root / "assets" / "icon.ico", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
print("ok")
