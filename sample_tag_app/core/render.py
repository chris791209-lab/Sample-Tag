"""Draw each sample-tag card as a PNG straight from the Target template's own
formatting (column widths, row heights, merges, borders, fonts, fills, number
formats, blank-field highlight, IMU traffic light). Replaces the Excel
CopyPicture macro: no Excel, no screen capture, so no blank images, and only
cards that have an item are drawn."""
from __future__ import annotations

import io
import math
import re
from datetime import date, datetime
from functools import lru_cache

import openpyxl
from PIL import Image, ImageDraw, ImageFont

from .builder import FIELD_MSG, imu_value
from .colors import hex_to_rgb, parse_theme, resolve, resolve_xml_color
from .config import BLOCK_START, CARD_COLS, CARD_ROWS, FONT_DIR, LEFT_COL
from .fit import TARGET_PT, WRAP, card_boxes, fit_size
from .layout import (BARCODE_BOX, FIELD_CELLS, IMU_CELLS, RETAIL_ECHO, THUMB_BOX, THUMB_CELL,
                     VENDOR_CELL, Geometry, parse_range)
from .xlsxpkg import Package

BORDER_W = {"hair": 1, "thin": 1, "dotted": 1, "dashed": 1, "dashDot": 1, "dashDotDot": 1,
            "medium": 2, "mediumDashed": 2, "mediumDashDot": 2, "mediumDashDotDot": 2,
            "slantDashDot": 2, "double": 3, "thick": 3}
BLANK_FILL_DEFAULT = "FFFF99"


@lru_cache(maxsize=64)
def _font(bold: bool, italic: bool, px: int):
    name = "Carlito-" + ("BoldItalic" if bold and italic else "Bold" if bold else "Italic" if italic else "Regular") + ".ttf"
    return ImageFont.truetype(str(FONT_DIR / name), max(6, px))


class CardRenderer:
    def __init__(self, template_bytes: bytes, scale: float = 2.0):
        self.s = scale
        pkg = Package(template_bytes)
        name = next((n for n, _, _ in pkg.sheets() if n.strip().lower() in ("template", "sample tag")), None)
        sheet_xml = pkg.text(pkg.sheet_path(name))
        self.theme = parse_theme(pkg.text("xl/theme/theme1.xml") if pkg.has("xl/theme/theme1.xml") else None)
        self.geo = Geometry(sheet_xml)
        wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
        ws = wb[name]

        # geometry of one card (prototype rows 2..25, cols B..F)
        self.col_x = [0]
        for dc in range(CARD_COLS):
            self.col_x.append(self.col_x[-1] + self.geo.col_px(LEFT_COL + dc) * self.s)
        self.row_y = [0]
        for dr in range(CARD_ROWS):
            self.row_y.append(self.row_y[-1] + self.geo.row_px(BLOCK_START + dr) * self.s)

        # cell styles / static labels
        self.cells = {}
        for dr in range(CARD_ROWS):
            for dc in range(CARD_COLS):
                c = ws.cell(BLOCK_START + dr, LEFT_COL + dc)
                fill = None
                if c.fill is not None and c.fill.fill_type == "solid":
                    fill = resolve(c.fill.fgColor, self.theme)
                b = c.border
                self.cells[(dr, dc)] = {
                    "value": c.value,
                    "font": (bool(c.font.b), bool(c.font.i), float(c.font.sz or 11),
                             resolve(c.font.color, self.theme, "000000") or "000000"),
                    "fill": fill,
                    "h": c.alignment.horizontal, "v": c.alignment.vertical or "bottom",
                    "wrap": bool(c.alignment.wrap_text),
                    "fmt": c.number_format or "General",
                    "border": {side: (getattr(b, side).style, resolve(getattr(b, side).color, self.theme, "000000") or "000000")
                               for side in ("left", "right", "top", "bottom")},
                }
        # same tweaks the workbook builder applies
        self.cells[(2, 4)]["fmt"] = "yyyy/mm/dd"     # sample date
        self.cells[(2, 1)]["wrap"] = True            # description
        self.boxes = card_boxes(sheet_xml, self.geo)  # fitted (bigger) data-cell fonts, as in the workbook
        # merges inside the prototype card
        self.merges = []
        for m in re.findall(r'<mergeCell ref="([^"]+)"/>', sheet_xml):
            r1, c1, r2, c2 = parse_range(m)
            if BLOCK_START <= r1 and r2 < BLOCK_START + CARD_ROWS and LEFT_COL <= c1 and c2 < LEFT_COL + CARD_COLS:
                self.merges.append((r1 - BLOCK_START, c1 - LEFT_COL, r2 - BLOCK_START, c2 - LEFT_COL))
        self.merge_of = {}
        for mg in self.merges:
            for r in range(mg[0], mg[2] + 1):
                for c in range(mg[1], mg[3] + 1):
                    self.merge_of[(r, c)] = mg

        # conditional formatting we emulate: "blank" highlight and "equals text" rules
        self.blank_rules, self.text_rules = self._read_cf(pkg, sheet_xml)
        # default IMU colours (template's own dxfs are equal to these in every template seen)
        self.imu_style = {"green": ("00B050", "FFFFFF"), "yellow": ("FFFF00", "000000"), "red": ("FF0000", "FFFFFF")}

    # ------------------------------------------------------------------
    def _read_cf(self, pkg, sheet_xml):
        styles = pkg.text("xl/styles.xml")
        dm = re.search(r"<dxfs[^>]*>(.*?)</dxfs>", styles, re.S)
        dxfs = re.findall(r"<dxf>.*?</dxf>|<dxf/>", dm.group(1), re.S) if dm else []
        blank, text = {}, {}
        for sq, body in re.findall(r'<conditionalFormatting sqref="([^"]+)">(.*?)</conditionalFormatting>', sheet_xml, re.S):
            cells = []
            for rng in sq.split():
                r1, c1, r2, c2 = parse_range(rng)
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        dr, dc = r - BLOCK_START, c - LEFT_COL
                        if 0 <= dr < CARD_ROWS and 0 <= dc < CARD_COLS:
                            cells.append((dr, dc))
            if not cells or all(dr == 0 for dr, _ in cells):
                continue        # vendor bar handled separately
            for rule in re.findall(r"<cfRule\b.*?</cfRule>", body, re.S):
                did = re.search(r'dxfId="(\d+)"', rule)
                d = dxfs[int(did.group(1))] if did and int(did.group(1)) < len(dxfs) else ""
                bg = re.search(r"<bgColor ([^/]*)/>", d)
                fc = re.search(r"<font>.*?<color ([^/]*)/>.*?</font>", d, re.S)
                st = {"fill": resolve_xml_color(bg.group(1), self.theme) if bg else None,
                      "color": resolve_xml_color(fc.group(1), self.theme) if fc else None,
                      "italic": ("<i/>" in d or '<i val="1"/>' in d) if "<font>" in d else None,
                      "bold": (False if '<b val="0"/>' in d else True if "<b/>" in d else None)}
                if 'type="containsBlanks"' in rule or re.search(r"<formula>\"\"</formula>", rule):
                    for k in cells:
                        blank.setdefault(k, st)
                else:
                    lit = re.search(r'operator="equal"[^>]*>\s*<formula>"([^"]+)"</formula>', rule)
                    if lit:
                        for k in cells:
                            text.setdefault(k, {})[lit.group(1).lower()] = st
        return blank, text

    # ------------------------------------------------------------------
    def card_values(self, card: dict, vendor: str) -> dict:
        vals = {}
        for k, field in FIELD_CELLS.items():
            vals[k] = card.get(field)
        vals[VENDOR_CELL] = vendor
        for k, (which, wd) in IMU_CELLS.items():
            if which == "imu":
                v = imu_value(card.get("first_cost"), card.get("retail"), card.get("duty"), wd)
            else:
                v = imu_value(card.get("ly_cost"), card.get("ly_retail"), card.get("ly_duty"), wd)
            vals[k] = FIELD_MSG if v is None else v
        vals[RETAIL_ECHO] = card.get("retail")
        if not card.get("thumb_png"):
            vals[THUMB_CELL] = self.cells[THUMB_CELL]["value"]
        else:
            vals[THUMB_CELL] = None
        for k, c in self.cells.items():
            if k not in vals and isinstance(c["value"], str) and not c["value"].startswith("="):
                vals[k] = c["value"]
        return vals

    # ------------------------------------------------------------------
    def render(self, card: dict, *, vendor: str, vendor_fill: str | None, goal: float | None,
               band: float, margin: int = 16) -> Image.Image:
        s = self.s
        W, H = int(self.col_x[-1]), int(self.row_y[-1])
        img = Image.new("RGB", (W + 2 * margin, H + 2 * margin), "white")
        d = ImageDraw.Draw(img)
        ox, oy = margin, margin
        vals = self.card_values(card, vendor)

        def rect(dr, dc):
            mg = self.merge_of.get((dr, dc))
            if mg:
                r1, c1, r2, c2 = mg
            else:
                r1, c1, r2, c2 = dr, dc, dr, dc
            return (ox + self.col_x[c1], oy + self.row_y[r1], ox + self.col_x[c2 + 1], oy + self.row_y[r2 + 1])

        # 1) fills + text (top-left cell of each merge only)
        for (dr, dc), st in self.cells.items():
            mg = self.merge_of.get((dr, dc))
            if mg and (mg[0], mg[1]) != (dr, dc):
                continue
            x0, y0, x1, y1 = rect(dr, dc)
            v = vals.get((dr, dc))
            fill = st["fill"]
            bold, italic, size, color = st["font"]
            if dr == 0 and vendor_fill:
                fill = vendor_fill
            if self._is_blank(v) and (dr, dc) in self.blank_rules and self.blank_rules[(dr, dc)]["fill"]:
                fill = self.blank_rules[(dr, dc)]["fill"]
            tr = self.text_rules.get((dr, dc), {}).get(str(v).lower()) if isinstance(v, str) else None
            if tr:
                fill = tr["fill"] or fill
                color = tr["color"] or color
                italic = tr["italic"] if tr["italic"] is not None else italic
                bold = tr["bold"] if tr["bold"] is not None else bold
            elif (dr, dc) in IMU_CELLS and isinstance(v, (int, float)) and goal is not None:
                kind = "green" if v >= goal else ("yellow" if v >= goal - band else "red")
                fill, color = self.imu_style[kind]
                bold = True
            if fill:
                d.rectangle([x0, y0, x1, y1], fill="#" + fill)
            if (dr, dc) in TARGET_PT and not self._is_blank(v):
                w, h = self.boxes[(dr, dc)]
                size = fit_size((dr, dc), v, fmt=st["fmt"], bold=bold, wrap=(dr, dc) in WRAP,
                                w_px=w, h_pt=h, base_pt=size) or size
            if not self._is_blank(v):
                self._text(d, (x0, y0, x1, y1), v, st, bold, italic, size, color)

        # 2) pictures
        if card.get("barcode_png"):
            self._picture(img, card["barcode_png"], BARCODE_BOX, ox, oy, pad=6 * s)
        if card.get("thumb_png"):
            self._picture(img, card["thumb_png"], THUMB_BOX, ox, oy, pad=4 * s)

        # 3) borders (only the outer edges of merged areas)
        for (dr, dc), st in self.cells.items():
            x0, y0 = ox + self.col_x[dc], oy + self.row_y[dr]
            x1, y1 = ox + self.col_x[dc + 1], oy + self.row_y[dr + 1]
            mg = self.merge_of.get((dr, dc))
            for side, (style, col) in st["border"].items():
                if not style:
                    continue
                if mg:
                    if side == "left" and dc != mg[1]:
                        continue
                    if side == "right" and dc != mg[3]:
                        continue
                    if side == "top" and dr != mg[0]:
                        continue
                    if side == "bottom" and dr != mg[2]:
                        continue
                w = max(1, int(round(BORDER_W.get(style, 1) * s)))
                fillc = "#" + col
                if side == "left":
                    d.rectangle([x0 - w // 2, y0 - w // 2, x0 + (w - 1) - w // 2, y1 + w // 2], fill=fillc)
                elif side == "right":
                    d.rectangle([x1 - w // 2, y0 - w // 2, x1 + (w - 1) - w // 2, y1 + w // 2], fill=fillc)
                elif side == "top":
                    d.rectangle([x0 - w // 2, y0 - w // 2, x1 + w // 2, y0 + (w - 1) - w // 2], fill=fillc)
                else:
                    d.rectangle([x0 - w // 2, y1 - w // 2, x1 + w // 2, y1 + (w - 1) - w // 2], fill=fillc)
        return img

    # ------------------------------------------------------------------
    @staticmethod
    def _is_blank(v) -> bool:
        if v is None:
            return True
        if isinstance(v, float) and math.isnan(v):
            return True
        return isinstance(v, str) and v.strip() == ""

    def _picture(self, img, png, box, ox, oy, pad):
        (r1, c1), (r2, c2) = box
        x0, y0 = ox + self.col_x[c1], oy + self.row_y[r1]
        x1, y1 = ox + self.col_x[c2 + 1], oy + self.row_y[r2 + 1]
        with Image.open(io.BytesIO(png)) as p:
            p = p.convert("RGBA")
            bw, bh = (x1 - x0) - 2 * pad, (y1 - y0) - 2 * pad
            k = min(bw / p.width, bh / p.height)
            w, h = max(1, int(p.width * k)), max(1, int(p.height * k))
            p = p.resize((w, h), Image.LANCZOS)
            img.paste(p, (int(x0 + (x1 - x0 - w) / 2), int(y0 + (y1 - y0 - h) / 2)), p)

    def _fmt(self, v, fmt: str):
        """-> (text, accounting) ; accounting=True means '$' left / number right."""
        if isinstance(v, (datetime, date)):
            dt = v
            f = fmt.lower()
            if "yyyy" in f:
                return dt.strftime("%Y/%m/%d"), False
            return dt.strftime("%m-%d-%y"), False
        if isinstance(v, bool):
            return ("TRUE" if v else "FALSE"), False
        if isinstance(v, (int, float)):
            if "%" in fmt:
                dec = len(re.search(r"0\.?(0*)%", fmt).group(1)) if re.search(r"0\.?(0*)%", fmt) else 0
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

    def _text(self, d, box, v, st, bold, italic, size, color):
        s = self.s
        x0, y0, x1, y1 = box
        px = int(round(size * 96 / 72 * s))
        f = _font(bold, italic, px)
        text, accounting = self._fmt(v, st["fmt"])
        pad = 3 * s
        fill = "#" + color
        ascent, descent = f.getmetrics()
        line_h = ascent + descent
        if accounting:
            cy = self._vpos(y0, y1, line_h, st["v"])
            d.text((x0 + pad, cy), "$", font=f, fill=fill)
            tw = d.textlength(text, font=f)
            d.text((x1 - pad - tw - 4 * s, cy), text, font=f, fill=fill)
            return
        maxw = (x1 - x0) - 2 * pad
        lines = []
        for para in str(text).split("\n"):
            if st["wrap"]:
                lines.extend(self._wrap(d, para, f, maxw))
            else:
                lines.append(para)
        # shrink a line that still does not fit (no wrap) so nothing is cut mid-word
        if not st["wrap"] and lines and d.textlength(lines[0], font=f) > maxw and px > 8:
            k = maxw / d.textlength(lines[0], font=f)
            f = _font(bold, italic, max(8, int(px * k)))
            ascent, descent = f.getmetrics()
            line_h = ascent + descent
        total = line_h * len(lines)
        y = self._vpos(y0, y1, total, st["v"])
        h = st["h"]
        if h is None or h == "general":
            h = "right" if isinstance(v, (int, float)) and not isinstance(v, bool) else "left"
        for ln in lines:
            tw = d.textlength(ln, font=f)
            if h in ("center", "centerContinuous", "distributed", "justify", "fill"):
                x = x0 + ((x1 - x0) - tw) / 2
            elif h == "right":
                x = x1 - pad - tw
            else:
                x = x0 + pad
            d.text((x, y), ln, font=f, fill=fill)
            y += line_h

    @staticmethod
    def _vpos(y0, y1, h, v):
        if v == "center":
            return y0 + ((y1 - y0) - h) / 2
        if v == "top":
            return y0 + 2
        return y1 - h - 2

    @staticmethod
    def _wrap(d, text, f, maxw):
        words = text.split(" ")
        lines, cur = [], ""
        for w in words:
            t = (cur + " " + w).strip()
            if d.textlength(t, font=f) <= maxw or not cur:
                cur = t
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
        return lines


def to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG", dpi=(192, 192), compress_level=6)
    return buf.getvalue()


def cards_pdf(images: list[Image.Image], per_row: int = 2, per_col: int = 4, dpi: int = 200) -> bytes:
    """A4 portrait pages, 2 x 4 cards per page (same as the Target template print)."""
    pw, ph = int(8.27 * dpi), int(11.69 * dpi)
    m = int(0.25 * dpi)
    gap = int(0.08 * dpi)
    cw = (pw - 2 * m - gap * (per_row - 1)) // per_row
    pages = []
    n = per_row * per_col
    for start in range(0, len(images), n):
        page = Image.new("RGB", (pw, ph), "white")
        chunk = images[start:start + n]
        ch = None
        for i, im in enumerate(chunk):
            k = cw / im.width
            th = int(im.height * k)
            ch = th
            row, col = divmod(i, per_row)
            x = m + col * (cw + gap)
            y = m + row * (th + gap)
            page.paste(im.resize((cw, th), Image.LANCZOS), (x, y))
        pages.append(page)
    if not pages:
        return b""
    buf = io.BytesIO()
    pages[0].save(buf, format="PDF", save_all=True, append_images=pages[1:], resolution=dpi)
    return buf.getvalue()
