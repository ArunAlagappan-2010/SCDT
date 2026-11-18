"""Draws the app icon (installer/mctm.ico): the mint "M" tile used in the web app."""

import os
import sys
from PIL import Image, ImageDraw, ImageFont

MINT = (161, 255, 203, 255)
MINT_INK = (15, 61, 38, 255)


def tile(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=int(size * 0.24), fill=MINT)
    font = None
    for name in ("arialbd.ttf", "segoeuib.ttf", "DejaVuSans-Bold.ttf"):
        try:
            font = ImageFont.truetype(name, int(size * 0.58))
            break
        except OSError:
            continue
    font = font or ImageFont.load_default()
    box = d.textbbox((0, 0), "M", font=font)
    w, h = box[2] - box[0], box[3] - box[1]
    d.text(((size - w) / 2 - box[0], (size - h) / 2 - box[1]), "M", font=font, fill=MINT_INK)
    return img


def main(out_path):
    sizes = [256, 128, 64, 48, 32, 24, 16]
    tile(256).save(out_path, format="ICO", sizes=[(s, s) for s in sizes])
    print(out_path)


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, "mctm.ico"))
