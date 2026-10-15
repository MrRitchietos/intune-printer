"""Normalize printer brand logos to 256x256 PNG icons for Intune / PSADT.

Each logo is trimmed, scaled to fit, and centred on a white rounded square so
dark logos stay visible on dark backgrounds and Intune never stretches them.
Usage: python3 normalize-logos.py <source-dir> <dest-dir>
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw

SIZE, RADIUS, PAD = 256, 44, 28

def normalize(src: Path, dst: Path) -> None:
    im = Image.open(src).convert("RGBA")
    bbox = im.getchannel("A").point(lambda a: 255 if a > 8 else 0).getbbox()
    if bbox:
        im = im.crop(bbox)
    box = SIZE - 2 * PAD
    scale = min(box / im.width, box / im.height)
    im = im.resize((max(1, round(im.width * scale)), max(1, round(im.height * scale))), Image.LANCZOS)
    canvas = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    mask = Image.new("L", (SIZE, SIZE), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, SIZE - 1, SIZE - 1), RADIUS, fill=255)
    canvas.paste(Image.new("RGBA", (SIZE, SIZE), (255, 255, 255, 255)), (0, 0), mask)
    canvas.alpha_composite(im, ((SIZE - im.width) // 2, (SIZE - im.height) // 2))
    canvas.save(dst, optimize=True)

if __name__ == "__main__":
    src_dir, dst_dir = Path(sys.argv[1]), Path(sys.argv[2])
    dst_dir.mkdir(parents=True, exist_ok=True)
    for f in sorted(src_dir.glob("*.png")):
        normalize(f, dst_dir / f.name.lower())
        print("normalized", f.name)
