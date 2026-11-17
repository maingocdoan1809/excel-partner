# -*- coding: utf-8 -*-
"""Evidence images: red boxes around the elements the step is about (numbered when several), a pointer glyph for hover / click, and a
caption strip UNDER the image (never over the UI). Fonts: a Latin font for Latin / Vietnamese, a CJK font for Japanese / Chinese / symbols
(searched on Windows, macOS and Linux; PIL default as last resort)."""
import datetime, os
from functools import lru_cache
from PIL import Image, ImageDraw, ImageFont
from . import browser
from .config import cfg

CAP = 56
LATIN = ["C:/Windows/Fonts/arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf", "/Library/Fonts/Arial.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans.ttf"]
CJK = ["C:/Windows/Fonts/meiryo.ttc", "C:/Windows/Fonts/msgothic.ttc", "C:/Windows/Fonts/YuGothM.ttc",
       "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc", "/System/Library/Fonts/Hiragino Sans GB.ttc",
       "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc"]


def _is_cjk(ch):
    return ord(ch) >= 0x2000          # Latin + Vietnamese are below U+2000


def _runs(text):
    out, cur, kind = [], "", None
    for ch in text:
        k = _is_cjk(ch)
        if kind is None or k == kind:
            cur += ch
        else:
            out.append((cur, kind)); cur = ch
        kind = k
    if cur:
        out.append((cur, kind))
    return out


@lru_cache(maxsize=None)
def font(kind, sz):
    for p in (CJK if kind else LATIN) + (LATIN if kind else CJK):
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, sz)
            except OSError:
                continue
    return ImageFont.load_default()


def text_width(dr, text, sz):
    return sum(dr.textlength(t, font=font(k, sz)) for t, k in _runs(text))


def draw_text(dr, xy, text, sz, fill):
    x, y = xy
    for t, k in _runs(text):
        f = font(k, sz)
        dr.text((x, y), t, font=f, fill=fill, anchor="lm")
        x += dr.textlength(t, font=f)


def as_locator(pg, s):
    """A selector string or a ready Playwright Locator (get_by_role(...)) -> Locator."""
    return pg.locator(s) if isinstance(s, str) else s


def rects(pg, sels):
    """Union bounding box (CSS px, viewport coordinates) per selector."""
    out = []
    for s in sels:
        loc = as_locator(pg, s)
        n = loc.count()
        if n == 0:
            raise RuntimeError(f"selector not found: {s}")
        bs = [b for b in (loc.nth(i).bounding_box() for i in range(min(n, 30))) if b]
        if not bs:
            raise RuntimeError(f"selector not visible: {s}")
        x0 = min(b["x"] for b in bs); y0 = min(b["y"] for b in bs)
        x1 = max(b["x"] + b["width"] for b in bs); y1 = max(b["y"] + b["height"] for b in bs)
        out.append((x0, y0, x1 - x0, y1 - y0))
    return out


def annotate(img, boxes, label, out, cursor=None):
    img = img.convert("RGB")
    W, H = img.size
    canvas = Image.new("RGB", (W, H + CAP), "white")
    canvas.paste(img, (0, 0))
    dr = ImageDraw.Draw(canvas)
    for k, (x, y, w, h) in enumerate(boxes, 1):
        p = 4
        # The frame stays INSIDE the picture: an element as wide as the screen would otherwise lose its left and right sides at the picture's edge and look unboxed.
        bx0, by0, bx1, by1 = max(x - p, 2), max(y - p, 2), min(x + w + p, W - 3), min(y + h + p, H - 3)
        dr.rectangle([bx0, by0, bx1, by1], outline=(255, 0, 0), width=4)
        if len(boxes) > 1:
            r = 14
            cx, cy = max(bx0 + r + 2, r + 2), max(by0 + r, r + 2)
            dr.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(255, 0, 0))
            dr.text((cx, cy), str(k), fill="white", font=font(False, 18), anchor="mm")
    if cursor:
        cx, cy = cursor
        dr.polygon([(cx, cy), (cx, cy + 26), (cx + 7, cy + 20), (cx + 12, cy + 31), (cx + 17, cy + 29), (cx + 12, cy + 18), (cx + 21, cy + 18)],
                   fill="black", outline="white")
    dr.rectangle([0, H, W, H + CAP], fill=(255, 245, 245))
    dr.line([0, H, W, H], fill=(255, 0, 0), width=3)
    sz = 24
    while text_width(dr, label, sz) > W - 36 and sz > 12:
        sz -= 1
    draw_text(dr, (18, H + CAP / 2), label, sz, (200, 0, 0))
    canvas.save(str(out))


def _nearest_visible(pg, sel, vw, vh):
    """Bounding box of the closest ancestor of `sel` that is visible and inside the viewport (an empty table has no body to box, its panel does)."""
    try:
        r = as_locator(pg, sel).first.evaluate("""e => { let n = e; for (let i = 0; i < 6 && n; i++) { const b = n.getBoundingClientRect();
            if (b.width >= 2 && b.height >= 2 && b.bottom > 0 && b.right > 0 && b.top < innerHeight && b.left < innerWidth) return [b.x, b.y, b.width, b.height]; n = n.parentElement; } return null; }""")
        if not r:
            return None
        x, y, w, h = r
        x0, y0, x1, y1 = max(x, 0), max(y, 0), min(x + w, vw), min(y + h, vh)
        return (x0, y0, x1 - x0, y1 - y0) if x1 - x0 >= 2 and y1 - y0 >= 2 else None
    except Exception:
        return None


def _name(tc):
    return f"TC{tc:02d}" if isinstance(tc, int) else str(tc)


def shot(tc, n, caption, hl=None, pg=None, wait=400, scroll=True, k=None, note=None, cursor=None):
    """Capture ONE evidence image. hl = selectors to box; k = image index inside the step (n-k); cursor = selector for the pointer glyph."""
    pg = pg or browser.page()
    pg.wait_for_timeout(wait)
    d = cfg.evidence_dir / _name(tc)
    d.mkdir(parents=True, exist_ok=True)
    tag = f"{n}-{k}" if k else f"{n}"
    base = f"{_name(tc)}_s{tag}"
    hl = [x for x in (hl or []) if x]
    if hl and scroll:
        try:
            as_locator(pg, hl[0]).first.evaluate("e=>e.scrollIntoView({block:'center',inline:'nearest'})")
            pg.wait_for_timeout(350)
        except Exception:
            pass
    css = rects(pg, hl) if hl else []
    vw, vh = pg.evaluate("[window.innerWidth, window.innerHeight]")
    boxes = []
    for (x, y, w, h), s_ in zip(css, hl):
        y0, y1, x0, x1 = max(y, 0), min(y + h, vh), max(x, 0), min(x + w, vw)
        if x1 - x0 < 2 or y1 - y0 < 2:
            fb = _nearest_visible(pg, s_, vw, vh)
            boxes.append(fb or (0, 0, vw, vh))
            caption = f"{caption} (「{str(s_)[:40]}」 not visible: {'nearest visible container' if fb else 'whole screen'} boxed)"
            continue
        boxes.append((x0, y0, x1 - x0, y1 - y0))
    img = browser.grab(pg)
    raw = d / f"{base}_raw.png"
    img.save(str(raw))
    cur = None
    if cursor:
        cb = rects(pg, [cursor])[0]
        cur = (cb[0] + cb[2] / 2, cb[1] + cb[3] / 2)
    ann = d / f"{base}_annotated.png"
    annotate(img, boxes, f"Step {tag}: {caption}", ann, cursor=cur)
    return dict(step=n, k=k, tag=tag, caption=caption, note=note or caption, raw=str(raw), annotated=str(ann), url=pg.url,
                ts=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
