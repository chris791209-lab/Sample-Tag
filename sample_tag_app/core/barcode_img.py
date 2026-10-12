"""Code 128 barcode as a PNG — no font to install, no Chr/ChrW code-page issue,
no internet. The bar pattern comes from python-barcode (it picks the optimal
Code 128 sets incl. the check digit); drawing is done here so module width and
size are exact and crisp at print resolution."""
from __future__ import annotations

import io
from functools import lru_cache

from barcode.codex import Code128
from PIL import Image, ImageDraw, ImageFont

from .config import FONT_DIR


@lru_cache(maxsize=8)
def _font(size: int):
    return ImageFont.truetype(str(FONT_DIR / "Carlito-Regular.ttf"), size)


def modules(data: str) -> str:
    """'1101…' module string (bars=1) incl. start, check digit and stop."""
    return Code128(data).build()[0]


def render(data: str, module_px: int = 3, height_px: int = 150, quiet_modules: int = 10,
           text: bool = True, text_px: int = 30) -> Image.Image:
    mods = modules(data)
    width = (len(mods) + 2 * quiet_modules) * module_px
    th = int(text_px * 1.35) if text else 0
    img = Image.new("RGB", (width, height_px + th), "white")
    d = ImageDraw.Draw(img)
    x = quiet_modules * module_px
    for bit in mods:
        if bit == "1":
            d.rectangle([x, 0, x + module_px - 1, height_px - 1], fill="black")
        x += module_px
    if text:
        f = _font(text_px)
        tw = d.textlength(data, font=f)
        d.text(((width - tw) / 2, height_px + text_px * 0.15), data, fill="black", font=f)
    return img


def png_bytes(data: str, **kw) -> bytes:
    buf = io.BytesIO()
    render(data, **kw).save(buf, format="PNG", dpi=(300, 300))
    return buf.getvalue()
