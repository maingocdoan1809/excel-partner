"""Evidence pictures: a red frame around the element a step is about, a caption strip UNDER the picture (never over the UI), and a hard check
that the frame is really there. Pillow does both the drawing and the check, so a picture that merely claims to be annotated is caught."""
import os
from functools import lru_cache


CAP = 56
RED = (255, 0, 0)
MIN_RED_PIXELS = 400          # a 4px frame around anything useful is thousands of pixels; text in red on the page never reaches this
LATIN = ["C:/Windows/Fonts/arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]
CJK = ["C:/Windows/Fonts/meiryo.ttc", "C:/Windows/Fonts/msgothic.ttc", "/System/Library/Fonts/Hiragino Sans GB.ttc", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"]


@lru_cache(maxsize=None)
def _font(cjk, size):
    from PIL import ImageFont
    for path in (CJK + LATIN) if cjk else (LATIN + CJK):
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default()


def _runs(text):
    """Split into Latin/Vietnamese runs and CJK runs: each needs its own font."""
    out = []
    for ch in text:
        cjk = ord(ch) >= 0x2000
        if out and out[-1][1] == cjk:
            out[-1][0] += ch
        else:
            out.append([ch, cjk])
    return out


def _width(draw, text, size):
    return sum(draw.textlength(run, font=_font(cjk, size)) for run, cjk in _runs(text))


def annotate(raw_path, boxes, label, out_path, ok=True):
    """Frame `boxes` ([(x, y, w, h)] in picture pixels) and write `label` under the picture. The frame stays INSIDE the picture: an element as wide as the
    screen would otherwise lose its sides at the edge and look unboxed. Several boxes are numbered."""
    from PIL import Image, ImageDraw
    image = Image.open(raw_path).convert("RGB")
    width, height = image.size
    canvas = Image.new("RGB", (width, height + CAP), "white")
    canvas.paste(image, (0, 0))
    draw = ImageDraw.Draw(canvas)
    for number, (x, y, w, h) in enumerate(boxes, 1):
        x0, y0, x1, y1 = max(x - 4, 2), max(y - 4, 2), min(x + w + 4, width - 3), min(y + h + 4, height - 3)
        draw.rectangle([x0, y0, x1, y1], outline=RED, width=4)
        if len(boxes) > 1:
            cx, cy = min(max(x1 - 18, 16), width - 18), max(y0 + 14, 16)           # at the box's right end, so it never covers the element's label
            draw.ellipse([cx - 14, cy - 14, cx + 14, cy + 14], fill=RED)
            draw.text((cx, cy), str(number), fill="white", font=_font(False, 18), anchor="mm")
    draw.rectangle([0, height, width, height + CAP], fill=(255, 245, 245) if not ok else (245, 250, 246))
    draw.line([0, height, width, height], fill=RED if not ok else (30, 120, 70), width=3)
    size = 24
    while _width(draw, label, size) > width - 36 and size > 12:
        size -= 1
    x = 18
    for run, cjk in _runs(label):
        draw.text((x, height + CAP / 2), run, font=_font(cjk, size), fill=(200, 0, 0) if not ok else (20, 90, 50), anchor="lm")
        x += draw.textlength(run, font=_font(cjk, size))
    canvas.save(out_path)


def red_pixels(path):
    """How many pure-red pixels the picture holds (the frame is drawn in exactly (255, 0, 0))."""
    from PIL import Image, ImageChops
    r, g, b = Image.open(path).convert("RGB").split()
    mask = ImageChops.multiply(ImageChops.multiply(r.point(lambda v: 255 if v >= 250 else 0), g.point(lambda v: 255 if v <= 5 else 0)),
                               b.point(lambda v: 255 if v <= 5 else 0))
    return mask.histogram()[255]


def check(raw_path, annotated_path):
    """None when the pair is sound, else why it is not: a missing or empty file, or no red frame."""
    for path in (raw_path, annotated_path):
        if not os.path.isfile(path) or os.path.getsize(path) < 1000:
            return f"picture missing or empty: {path}"
    if red_pixels(annotated_path) < MIN_RED_PIXELS:
        return f"no red frame in {annotated_path}: the step's element was not boxed"
    return None
