"""Bigger, auto-fitted fonts for the data cells of a Sample Tag.

Each value gets the largest font size (up to the cap below) that still fits
inside its cell — width *and* the fixed row height — so nothing is cut by the
borders. Carlito has exactly Calibri's metrics, so the size chosen here is what
Excel shows; the workbook builder and the PNG renderer both call fit_size()."""
from __future__ import annotations

import re
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path

from PIL import ImageFont

from .layout import RETAIL_ECHO

FONT_DIR = Path(__file__).resolve().parent.parent / "fonts"

# (row, col) offset in the card -> font size cap (pt). Template default is 11 pt.
TARGET_PT = {
    (2, 1): 14,   # product description (wraps, 2 lines)
    (2, 4): 14,   # sample date
    (5, 1): 14,   # MFG style #
    (5, 4): 14,   # PCN item ID
    (6, 1): 14,   # APEX ID
    (6, 4): 14,   # Spark PID
    (7, 1): 14,   # first cost
    (7, 4): 14,   # country of export
    (8, 1): 14,   # PCN retail
    (8, 4): 14,   # duty rate
    RETAIL_ECHO: 28,   # big retail, bottom right (template 20 pt)
}
MIN_PT = 8.0
STEP = 0.5
LINE = 1.33        # Excel row height needed per pt of Calibri (11 pt -> 15 pt row)
PAD_PX = 5         # cell margin each side, plus a little safety


@lru_cache(maxsize=256)
def _font(bold: bool, px: int):
    name = "Carlito-Bold.ttf" if bold else "Carlito-Regular.ttf"
    return ImageFont.truetype(str(FONT_DIR / name), max(6, px))


def _w(text: str, bold: bool, pt: float) -> float:
    """Text width in screen px (96 dpi) at `pt` points (measured at 4x for precision)."""
    f = _font(bold, int(round(pt * 96 / 72 * 4)))
    return f.getlength(text) / 4


def fmt_value(v, fmt: str):
    """-> (text, accounting) the way Excel shows the value; accounting = '$' left, number right."""
    if isinstance(v, (datetime, date)):
        return (v.strftime("%Y/%m/%d") if "yyyy" in fmt.lower() else v.strftime("%m-%d-%y")), False
    if isinstance(v, bool):
        return ("TRUE" if v else "FALSE"), False
    if isinstance(v, (int, float)):
        if "%" in fmt:
            m = re.search(r"0\.?(0*)%", fmt)
            dec = len(m.group(1)) if m else 0
            return f"{v * 100:.{dec}f}%", False
        if '"$"' in fmt or "$" in fmt:
            if abs(v) < 1e-12:
                return "-", True
            txt = f"{abs(v):,.2f}"
            return (f"({txt})" if v < 0 else txt), True
        if "#,##0.00" in fmt:
            return f"{v:,.2f}", False
        if float(v).is_integer():
            return str(int(v)), False
        return f"{v:.10g}", False
    return str(v), False


def _wrap_lines(text: str, bold: bool, pt: float, maxw: float) -> int:
    n = 0
    for para in str(text).split("\n"):
        cur = ""
        n += 1
        for w in para.split(" "):
            t = (cur + " " + w).strip()
            if _w(t, bold, pt) <= maxw or not cur:
                cur = t
            else:
                n += 1
                cur = w
    return n


def fits(text: str, accounting: bool, bold: bool, wrap: bool, w_px: float, h_pt: float, pt: float) -> bool:
    maxw = w_px - 2 * PAD_PX
    if accounting:
        need = _w("$", bold, pt) + _w(" " + text + " ", bold, pt) + 8
        return need <= maxw and pt * LINE <= h_pt
    if wrap:
        lines = _wrap_lines(text, bold, pt, maxw)
        widest_word = max((_w(w, bold, pt) for w in str(text).split()), default=0)
        return widest_word <= maxw and lines * pt * LINE <= h_pt
    return _w(text, bold, pt) <= maxw and pt * LINE <= h_pt


def fit_size(key, value, *, fmt: str, bold: bool, wrap: bool, w_px: float, h_pt: float,
             base_pt: float) -> float | None:
    """Largest size <= cap that fits; None = leave the template size (not a fitted cell / blank)."""
    cap = TARGET_PT.get(key)
    if cap is None or value is None or (isinstance(value, str) and not value.strip()):
        return None
    try:
        if value != value:          # NaN
            return None
    except Exception:
        pass
    text, acc = fmt_value(value, fmt)
    pt = float(cap)
    while pt > MIN_PT and not fits(text, acc, bold, wrap, w_px, h_pt, pt):
        pt -= STEP
    return pt if pt != base_pt else None


ACCOUNTING = '_("$"* #,##0.00_)'
FIELD_FMT = {(2, 4): "yyyy/mm/dd", (7, 1): ACCOUNTING, (8, 1): ACCOUNTING,
             RETAIL_ECHO: ACCOUNTING, (8, 4): "0.00%"}
WRAP = {(2, 1)}                      # description wraps; the rest are single-line values


def card_boxes(sheet_xml: str, geo) -> dict:
    """{(dr, dc): (width px, height pt)} for every fitted cell, merges included."""
    from .config import BLOCK_START, CARD_COLS, LEFT_COL
    from .layout import parse_range
    merges = []
    for m in re.findall(r'<mergeCell ref="([^"]+)"/>', sheet_xml):
        r1, c1, r2, c2 = parse_range(m)
        merges.append((r1 - BLOCK_START, c1 - LEFT_COL, r2 - BLOCK_START, c2 - LEFT_COL))
    out = {}
    for (dr, dc) in TARGET_PT:
        box = next(((a, b, c, d) for a, b, c, d in merges if a <= dr <= c and b <= dc <= d), (dr, dc, dr, dc))
        r1, c1, r2, c2 = box
        w = sum(geo.col_px(LEFT_COL + c) for c in range(c1, c2 + 1))
        h = sum(geo.row_pt(BLOCK_START + r) for r in range(r1, r2 + 1))
        out[(dr, dc)] = (w, h)
    return out


def card_value(key, card: dict):
    from .layout import FIELD_CELLS
    return card.get("retail") if key == RETAIL_ECHO else card.get(FIELD_CELLS.get(key))
