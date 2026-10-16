"""Build the Sample Tag workbook from a Target template by XML surgery.

The template's first card (rows 2-26, columns B:F) is the prototype. It is
replicated once per item - left card in B:F, right card in H:L - so every card
has identical borders, merges and fonts. Values are written directly (the app
is now the engine: no PCN Data / sheet001 / VBA needed), IMU formulas stay live,
the barcode and thumbnail are embedded pictures, and the IMU traffic light uses
each card's own program goal. Instructions, Dropdowns, workbook protection,
printer settings and everything else in the template are carried over as-is."""
from __future__ import annotations

import hashlib
import math
import re
import uuid
from html import escape

from PIL import Image
import io

from .config import BLOCK_ROWS, BLOCK_START, BLOCKS_PER_PAGE, CARD_COLS, LAST_COL, LEFT_COL, RIGHT_COL
from .fit import FIELD_FMT, TARGET_PT, WRAP, card_boxes, card_value, fit_size
from .layout import (BARCODE_BOX, EMU_PER_PT, EMU_PER_PX, FIELD_CELLS, IMU_CELLS, NUMERIC_FIELDS,
                     RETAIL_ECHO, THUMB_BOX, THUMB_CELL, VENDOR_CELL, Geometry, card_origin,
                     excel_serial, in_proto_card, parse_range, ref, shift_formula, shift_range)
from .stylebook import StyleBook
from .xlsxpkg import Package

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_SHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

FIELD_MSG = "FIELD WILL AUTO CALC"


# ----------------------------------------------------------------------------
# small XML helpers
# ----------------------------------------------------------------------------
def _x(s) -> str:
    return escape(str(s), quote=False)


def _attr_s(attrs: str) -> str:
    m = re.search(r'\ss="(\d+)"', attrs)
    return f' s="{m.group(1)}"' if m else ""


def _num(v) -> str:
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return repr(float(v)) if isinstance(v, float) else str(v)


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    return isinstance(v, str) and v.strip() == ""


def _parse_cells(row_inner: str):
    for m in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', row_inner, re.S):
        yield m.group(1), m.group(3), m.group(4)


def _col_idx(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def imu_value(cost, retail, duty, with_duty):
    if _is_blank(cost) or _is_blank(retail):
        return None
    try:
        retail = float(retail)
        cost = float(cost)
        d = 0.0 if _is_blank(duty) else float(duty)
        landed = cost * d + cost if with_duty else cost
        return (retail - landed) / retail
    except (TypeError, ValueError, ZeroDivisionError):
        return None


# ----------------------------------------------------------------------------
# main entry
# ----------------------------------------------------------------------------
def build_workbook(template_bytes: bytes, cards: list[dict], *, vendor: str,
                   yellow_band: float, vendor_fill: str | None = None,
                   sheet_name: str = "Sample Tag", add_data_sheet: bool = True) -> bytes:
    pkg = Package(template_bytes)
    sheets = pkg.sheets()
    tname = next((n for n, _, _ in sheets if n.strip().lower() in ("template", "sample tag")), None)
    if tname is None:
        raise ValueError("範本中找不到 Template / Sample Tag 工作表")
    tpath = pkg.sheet_path(tname)
    sx = pkg.text(tpath)
    geo = Geometry(sx)

    head, sheet_data, tail = _split_sheet(sx)
    rows = _parse_rows(sheet_data)
    proto = {}          # (dr, col) -> (attrs, inner)
    proto_row_attrs = {}
    for r in range(BLOCK_START, BLOCK_START + BLOCK_ROWS):
        attrs, inner = rows.get(r, ("", ""))
        proto_row_attrs[r - BLOCK_START] = re.sub(r'\s(?:r|spans)="[^"]*"', "", attrs)
        for col, cattrs, cinner in _parse_cells(inner):
            proto[(r - BLOCK_START, _col_idx(col))] = (re.sub(r'\sr="[^"]*"', "", cattrs), cinner)

    # Style tweaks are clones of the template styles (originals untouched):
    #  * sample date shown as yyyy/mm/dd
    #  * description wraps (long PCN descriptions were cut off)
    #  * vendor bar gets its colour as a plain fill instead of cross-sheet
    #    conditional formatting, so Excel, Mac Numbers, LibreOffice and the PNGs agree
    book = StyleBook(pkg.text("xl/styles.xml"))
    def sidx(dr, dc):
        m = re.search(r'\ss="(\d+)"', proto.get((dr, LEFT_COL + dc), ("", None))[0] or "")
        return int(m.group(1)) if m else 0
    style_over = {(2, 4): book.clone(sidx(2, 4), numfmt="yyyy/mm/dd"),
                  (2, 1): book.clone(sidx(2, 1), wrap=True)}
    if vendor_fill:
        for dc in range(CARD_COLS):
            style_over[(0, dc)] = book.clone(sidx(0, dc), fill_rgb=vendor_fill)
    date_style = style_over[(2, 4)]
    # bigger, auto-fitted fonts for the data cells (never wider/taller than the cell)
    boxes = card_boxes(sx, geo)

    def fitted_style(key, card):
        base = style_over.get(key, sidx(*key))
        fx = book.fonts[int(re.search(r'fontId="(\d+)"', book.xfs[base]).group(1))] if 'fontId="' in book.xfs[base] else ""
        w, h = boxes[key]
        pt = fit_size(key, card_value(key, card), fmt=FIELD_FMT.get(key, "@"), bold="<b/>" in fx or '<b val="1"/>' in fx,
                      wrap=key in WRAP, w_px=w, h_pt=h, base_pt=-1)
        return book.clone(base, font_size=pt) if pt else base

    n = len(cards)
    nblocks = max(1, math.ceil(n / 2))
    last_row = BLOCK_START + nblocks * BLOCK_ROWS - 1

    # ---- sheetData ---------------------------------------------------------
    out_rows = []
    row1 = rows.get(1)
    if row1:
        a, inner = row1
        out_rows.append(f"<row{a}>{inner}</row>" if inner else f"<row{a}/>")
    for b in range(nblocks):
        for dr in range(BLOCK_ROWS):
            r = BLOCK_START + b * BLOCK_ROWS + dr
            cells = []
            for col in range(1, LAST_COL + 1):
                side_card = None
                if LEFT_COL <= col < LEFT_COL + CARD_COLS:
                    side_card, dc = 2 * b, col - LEFT_COL
                elif RIGHT_COL <= col < RIGHT_COL + CARD_COLS:
                    side_card, dc = 2 * b + 1, col - RIGHT_COL
                if side_card is not None:
                    if side_card >= n:
                        continue                      # odd count: no empty right card
                    pa, pi = proto.get((dr, LEFT_COL + dc), ("", None))
                    ds = date_style
                    if (dr, dc) in TARGET_PT:
                        fs = fitted_style((dr, dc), cards[side_card])
                        pa = re.sub(r'\ss="\d+"', "", pa or "") + f' s="{fs}"'
                        if (dr, dc) == (2, 4):
                            ds = fs
                    elif (dr, dc) in style_over:
                        pa = re.sub(r'\ss="\d+"', "", pa or "") + f' s="{style_over[(dr, dc)]}"'
                    cells.append(_card_cell(r, col, dr, dc, pa, pi, cards[side_card], vendor, ds))
                else:
                    pa, pi = proto.get((dr, col), (None, None))
                    if pa is not None:
                        cells.append(f'<c r="{ref(r, col)}"{_attr_s(pa)}/>')
            ra = proto_row_attrs.get(dr, "")
            out_rows.append(f'<row r="{r}"{ra}>{"".join(cells)}</row>')
    new_sheet_data = "<sheetData>" + "".join(out_rows) + "</sheetData>"

    # ---- card offsets (for merges / CF / DV) -------------------------------
    offsets = []
    for i in range(n):
        top, left = card_origin(i)
        offsets.append((top - BLOCK_START, left - LEFT_COL))
    if not offsets:
        offsets = [(0, 0)]

    tail = _rebuild_tail(tail, offsets, cards, yellow_band, last_row, book)
    pkg.set_text("xl/styles.xml", book.to_xml())

    # ---- head: dimension + view -------------------------------------------
    head = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:{ref(last_row, LAST_COL)}"/>', head, count=1)
    head = re.sub(r'topLeftCell="[^"]*"', 'topLeftCell="A1"', head)
    head = re.sub(r'<selection [^>]*/>', '<selection activeCell="A1" sqref="A1"/>', head)

    sx = head + new_sheet_data + tail

    # ---- pictures ------------------------------------------------------------
    sx = _add_pictures(pkg, tpath, sx, cards, geo)
    pkg.set_text(tpath, sx)

    # ---- workbook: rename sheet, print area, Instructions ------------------
    _rename_sheet(pkg, tname, sheet_name, last_row)
    progs = {c.get("program") for c in cards if c.get("program")}
    if len(progs) == 1:
        _set_instructions_program(pkg, progs.pop())
    if add_data_sheet:
        _add_data_sheet(pkg, cards, proto, date_style)

    _strip_vba(pkg)
    pkg.drop_calc_chain()
    pkg.force_full_calc()
    return pkg.to_bytes()


def _strip_vba(pkg: Package) -> None:
    """Output is a plain .xlsx: the app replaces every macro the old working file had."""
    for n in [n for n in pkg.parts if n.startswith("xl/vbaProject") or n.endswith("vbaProject.bin.rels")]:
        pkg.remove(n)
    rels = pkg.text("xl/_rels/workbook.xml.rels")
    rels = re.sub(r'<Relationship[^>]*Target="vbaProject\.bin"[^>]*/>', "", rels)
    pkg.set_text("xl/_rels/workbook.xml.rels", rels)
    ct = pkg.text("[Content_Types].xml")
    ct = re.sub(r'<Override PartName="/xl/vbaProject[^"]*"[^>]*/>', "", ct)
    ct = ct.replace("application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    pkg.set_text("[Content_Types].xml", ct)


# ----------------------------------------------------------------------------
# cells
# ----------------------------------------------------------------------------
def _card_cell(r, col, dr, dc, pattrs, pinner, card, vendor, date_style) -> str:
    a1 = ref(r, col)
    s = _attr_s(pattrs or "")
    key = (dr, dc)

    if key == VENDOR_CELL:
        return f'<c r="{a1}"{s} t="str"><f>Instructions!$B$4</f><v>{_x(vendor)}</v></c>'

    if key in FIELD_CELLS:
        field = FIELD_CELLS[key]
        v = card.get(field)
        if field == "sample_date":
            serial = excel_serial(v)
            return f'<c r="{a1}" s="{date_style}"><v>{serial}</v></c>' if serial else f'<c r="{a1}" s="{date_style}"/>'
        if _is_blank(v):
            return f'<c r="{a1}"{s}/>'
        if field in NUMERIC_FIELDS:
            try:
                return f'<c r="{a1}"{s}><v>{_num(float(v))}</v></c>'
            except (TypeError, ValueError):
                pass
        return f'<c r="{a1}"{s} t="inlineStr"><is><t xml:space="preserve">{_x(v)}</t></is></c>'

    if key in IMU_CELLS:
        which, with_duty = IMU_CELLS[key]
        if which == "imu":
            cost_ref, retail_ref, duty_ref = "C9", "C10", "F10"
            val = imu_value(card.get("first_cost"), card.get("retail"), card.get("duty"), with_duty)
        else:
            cost_ref, retail_ref, duty_ref = "C15", "C16", "C17"
            val = imu_value(card.get("ly_cost"), card.get("ly_retail"), card.get("ly_duty"), with_duty)
        core = (f"({retail_ref}-(({duty_ref}*{cost_ref})+{cost_ref}))/{retail_ref}" if with_duty
                else f"({retail_ref}-{cost_ref})/{retail_ref}")
        # Guard added on purpose: the original formula shows 100% IMU when the cost is blank.
        f = f'IF(OR({cost_ref}="",{retail_ref}=""),"{FIELD_MSG}",IFERROR({core},"{FIELD_MSG}"))'
        f = _shift_proto(f, r, col, dr, dc)
        if val is None:
            return f'<c r="{a1}"{s} t="str"><f>{_x(f)}</f><v>{FIELD_MSG}</v></c>'
        return f'<c r="{a1}"{s}><f>{_x(f)}</f><v>{_num(val)}</v></c>'

    if key == BARCODE_BOX[0]:
        return f'<c r="{a1}"{s}/>'          # picture goes here (old Code128B / IMAGE formula dropped)

    if key == RETAIL_ECHO:
        f = _shift_proto("C10", r, col, dr, dc)
        v = card.get("retail")
        vv = "" if _is_blank(v) else f"<v>{_num(float(v))}</v>"
        return f'<c r="{a1}"{s}><f>{f}</f>{vv}</c>'

    if key == THUMB_CELL and card.get("thumb_png"):
        return f'<c r="{a1}"{s}/>'

    if not pinner:
        return f'<c r="{a1}"{s}/>'
    fm = re.search(r"<f>(.*?)</f>", pinner, re.S)
    if fm:
        f = _shift_proto(fm.group(1), r, col, dr, dc)
        return f'<c r="{a1}"{s}><f>{f}</f></c>'
    t = re.search(r'\st="(\w+)"', pattrs or "")
    tt = f' t="{t.group(1)}"' if t and t.group(1) in ("s", "inlineStr", "b", "n", "str") else ""
    return f'<c r="{a1}"{s}{tt}>{pinner}</c>'


def _shift_proto(formula: str, r: int, col: int, dr: int, dc: int) -> str:
    """Formula written for the prototype cell (row BLOCK_START+dr, col LEFT_COL+dc)
    -> the same formula for the target cell (r, col)."""
    return shift_formula(formula, r - (BLOCK_START + dr), col - (LEFT_COL + dc))


# ----------------------------------------------------------------------------
# sheet parts
# ----------------------------------------------------------------------------
def _split_sheet(sx: str):
    i = sx.find("<sheetData")
    m = re.search(r"<sheetData>.*?</sheetData>|<sheetData/>", sx, re.S)
    return sx[:i], m.group(0), sx[m.end():]


def _parse_rows(sheet_data: str) -> dict[int, tuple[str, str]]:
    out = {}
    for m in re.finditer(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", sheet_data, re.S):
        attrs = m.group(1)
        r = int(re.search(r'\br="(\d+)"', attrs).group(1))
        out[r] = (attrs, m.group(2) or "")
    return out




def _first_cell(rng: str):
    r1, c1, _, _ = parse_range(rng)
    return r1, c1


def _expand_sqref(sqref: str, offsets) -> tuple[str, tuple[int, int], tuple[int, int]] | None:
    """Keep the prototype-card ranges of a sqref and replicate them for every
    card. Returns (new sqref, old anchor, new anchor) or None."""
    ranges = sqref.split()
    kept = [r for r in ranges if in_proto_card(r)]
    if not kept:
        return None
    out = []
    for dr, dc in offsets:
        out.extend(shift_range(r, dr, dc) for r in kept)
    return " ".join(out), _first_cell(ranges[0]), _first_cell(kept[0])


DEFAULT_IMU_DXF = {
    "green": '<dxf><font><b/><i val="0"/><color theme="0"/></font><fill><patternFill><bgColor rgb="FF00B050"/></patternFill></fill></dxf>',
    "yellow": '<dxf><font><b/><i val="0"/></font><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>',
    "red": '<dxf><font><b/><i val="0"/><color theme="0"/></font><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>',
}


def _imu_cf(tail: str, offsets, cards, band, book: StyleBook) -> str:
    """IMU traffic light as *standard* conditional formatting with each card's
    own goal as a literal (D240 mixes subclasses on one sheet). The colours are
    taken from the template's own IMU rule. Standard CF renders identically in
    Excel, Mac Numbers, Google Sheets and LibreOffice (the template's x14 rule
    pointing at another sheet does not)."""
    dx = {}
    for blk in re.findall(r"<x14:conditionalFormatting\b.*?</x14:conditionalFormatting>", tail, re.S):
        sq = re.search(r"<xm:sqref>(.*?)</xm:sqref>", blk)
        if not sq or not any(in_proto_card(r) for r in sq.group(1).split()) or not _is_imu_block(sq.group(1)):
            continue
        for rule in re.findall(r"<x14:cfRule\b.*?</x14:cfRule>", blk, re.S):
            op = re.search(r'operator="(\w+)"', rule)
            d = re.search(r"<x14:dxf>(.*?)</x14:dxf>", rule, re.S)
            if op and d:
                kind = {"greaterThanOrEqual": "green", "greaterThan": "green", "between": "yellow",
                        "lessThanOrEqual": "red", "lessThan": "red"}.get(op.group(1))
                if kind:
                    dx[kind] = "<dxf>" + d.group(1) + "</dxf>"
        break
    ids = {k: book.add_dxf(dx.get(k, DEFAULT_IMU_DXF[k])) for k in ("green", "yellow", "red")}
    pr = [int(p) for p in re.findall(r'priority="(\d+)"', tail)]
    nxt = (max(pr) if pr else 100) + 1
    groups: dict[float, list[int]] = {}
    for i, c in enumerate(cards):
        g = c.get("goal")
        if g is not None and not (isinstance(g, float) and math.isnan(g)):
            groups.setdefault(round(float(g), 6), []).append(i)
    proto_ranges = [ref(BLOCK_START + dr, LEFT_COL + dc) for (dr, dc) in IMU_CELLS]
    out = []
    for g, idxs in groups.items():
        sq = []
        for i in idxs:
            dr0, dc0 = offsets[i]
            sq.extend(shift_range(r, dr0, dc0) for r in proto_ranges)
        lo = round(g - band, 6)
        out.append(
            f'<conditionalFormatting sqref="{" ".join(sq)}">'
            f'<cfRule type="cellIs" dxfId="{ids["green"]}" priority="{nxt}" operator="greaterThanOrEqual"><formula>{g}</formula></cfRule>'
            f'<cfRule type="cellIs" dxfId="{ids["yellow"]}" priority="{nxt + 1}" operator="between"><formula>{lo}</formula><formula>{g}</formula></cfRule>'
            f'<cfRule type="cellIs" dxfId="{ids["red"]}" priority="{nxt + 2}" operator="lessThan"><formula>{lo}</formula></cfRule>'
            "</conditionalFormatting>")
        nxt += 3
    return "".join(out)


def _rebuild_tail(tail: str, offsets, cards, band, last_row, book: StyleBook) -> str:
    imu_cf_xml = _imu_cf(tail, offsets, cards, band, book)
    # remove what we regenerate
    merges = re.search(r"<mergeCells\b.*?</mergeCells>", tail, re.S)
    merge_list = re.findall(r'<mergeCell ref="([^"]+)"/>', merges.group(0)) if merges else []
    tail = re.sub(r"<mergeCells\b.*?</mergeCells>", "", tail, flags=re.S)

    cf_blocks = re.findall(r'<conditionalFormatting\b[^>]*>.*?</conditionalFormatting>', tail, re.S)
    tail = re.sub(r'<conditionalFormatting\b[^>]*>.*?</conditionalFormatting>', "", tail, flags=re.S)

    dv_block = re.search(r"<dataValidations\b.*?</dataValidations>", tail, re.S)
    tail = re.sub(r"<dataValidations\b.*?</dataValidations>", "", tail, flags=re.S)
    tail = re.sub(r"<rowBreaks\b.*?</rowBreaks>|<rowBreaks[^>]*/>", "", tail, flags=re.S)

    # merges
    new_merges = []
    for m in merge_list:
        if in_proto_card(m):
            new_merges.extend(shift_range(m, dr, dc) for dr, dc in offsets)
    merge_xml = f'<mergeCells count="{len(new_merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in new_merges) + "</mergeCells>" if new_merges else ""

    # normal conditional formatting
    cf_xml = []
    for blk in cf_blocks:
        sq = re.search(r'sqref="([^"]+)"', blk).group(1)
        if _is_vendor_bar(sq):
            continue                      # replaced by a static fill
        ex = _expand_sqref(sq, offsets)
        if not ex:
            continue
        new_sq, (ar, ac), (kr, kc) = ex
        body = re.sub(r'^<conditionalFormatting\b[^>]*>|</conditionalFormatting>$', "", blk)
        body = re.sub(r"<formula>(.*?)</formula>",
                      lambda m: f"<formula>{shift_formula(m.group(1), kr - ar, kc - ac)}</formula>", body, flags=re.S)
        attrs = re.search(r"^<conditionalFormatting\b([^>]*)>", blk).group(1)
        attrs = re.sub(r'sqref="[^"]+"', f'sqref="{new_sq}"', attrs)
        cf_xml.append(f"<conditionalFormatting{attrs}>{body}</conditionalFormatting>")
    cf_xml.append(imu_cf_xml)       # after the template rules -> "FIELD WILL AUTO CALC" keeps precedence

    # data validations
    dv_xml = ""
    if dv_block:
        items = []
        for dv in re.findall(r"<dataValidation\b[^>]*?(?:/>|>.*?</dataValidation>)", dv_block.group(0), re.S):
            sq = re.search(r'sqref="([^"]+)"', dv).group(1)
            ex = _expand_sqref(sq, offsets)
            if ex:
                items.append(re.sub(r'sqref="[^"]+"', f'sqref="{ex[0]}"', dv, count=1))
        if items:
            dv_xml = f'<dataValidations count="{len(items)}">' + "".join(items) + "</dataValidations>"

    # x14 extension: vendor/IMU rules are replaced above; anything else is replicated
    tail = _rebuild_ext(tail, offsets, cards, band)

    # row breaks: same page size as the Target template (BLOCKS_PER_PAGE blocks)
    per = BLOCK_ROWS * BLOCKS_PER_PAGE
    brks = list(range(1 + per, last_row, per))
    rb_xml = (f'<rowBreaks count="{len(brks)}" manualBreakCount="{len(brks)}">' +
              "".join(f'<brk id="{b}" max="16383" man="1"/>' for b in brks) + "</rowBreaks>") if brks else ""

    # pageSetup: fit width only so the manual page breaks are honoured
    def ps(m):
        t = m.group(0)
        t = re.sub(r'\s(?:scale|fitToWidth|fitToHeight)="[^"]*"', "", t)
        return t.replace("<pageSetup", '<pageSetup fitToWidth="1" fitToHeight="0"', 1)
    tail = re.sub(r"<pageSetup\b[^>]*/>", ps, tail, count=1)

    tail = _insert(tail, merge_xml, ["phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks",
                                     "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks",
                                     "colBreaks", "drawing", "legacyDrawing", "tableParts", "extLst"])
    tail = _insert(tail, "".join(cf_xml), ["dataValidations", "hyperlinks", "printOptions", "pageMargins",
                                           "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "drawing",
                                           "legacyDrawing", "tableParts", "extLst"])
    tail = _insert(tail, dv_xml, ["hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter",
                                  "rowBreaks", "colBreaks", "drawing", "legacyDrawing", "tableParts", "extLst"])
    tail = _insert(tail, rb_xml, ["colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags",
                                  "drawing", "legacyDrawing", "tableParts", "extLst"])
    return tail


def _insert(tail: str, xml: str, before: list[str]) -> str:
    if not xml:
        return tail
    pos = None
    for tag in before:
        m = re.search(rf"<{tag}\b", tail)
        if m and (pos is None or m.start() < pos):
            pos = m.start()
    if pos is None:
        pos = tail.rfind("</worksheet>")
    return tail[:pos] + xml + tail[pos:]


def _rebuild_ext(tail: str, offsets, cards, band) -> str:
    m = re.search(r"<extLst>(.*?)</extLst>", tail, re.S)
    if not m:
        return tail
    exts = re.findall(r"<ext\b.*?</ext>", m.group(1), re.S)
    new_exts = []
    for ext in exts:
        if "<x14:conditionalFormattings>" in ext:
            ext = _rebuild_x14_cf(ext, offsets, cards, band)
            if ext is None:
                continue
        elif "<x14:dataValidations" in ext:
            ext = _rebuild_x14_dv(ext, offsets)
            if ext is None:
                continue
        new_exts.append(ext)
    new = "<extLst>" + "".join(new_exts) + "</extLst>" if new_exts else ""
    return tail[:m.start()] + new + tail[m.end():]


def _is_vendor_bar(sqref: str) -> bool:
    kept = [r for r in sqref.split() if in_proto_card(r)]
    return bool(kept) and all(parse_range(r)[0] == BLOCK_START and parse_range(r)[2] == BLOCK_START for r in kept)


def _is_imu_block(sqref: str) -> bool:
    imu_cells = {ref(BLOCK_START + dr, LEFT_COL + dc) for (dr, dc) in IMU_CELLS}
    for rng in sqref.split():
        r1, c1, r2, c2 = parse_range(rng)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if ref(r, c) in imu_cells:
                    return True
    return False


def _rebuild_x14_cf(ext: str, offsets, cards, band) -> str | None:
    blocks = re.findall(r"<x14:conditionalFormatting\b.*?</x14:conditionalFormatting>", ext, re.S)
    pr = [int(p) for p in re.findall(r'priority="(\d+)"', ext)]
    next_pr = (max(pr) if pr else 1000) + 1
    out = []
    imu_template = None
    for blk in blocks:
        sq = re.search(r"<xm:sqref>(.*?)</xm:sqref>", blk).group(1)
        if not any(in_proto_card(r) for r in sq.split()):
            continue
        if _is_imu_block(sq):
            imu_template = blk
            continue
        if _is_vendor_bar(sq):
            continue
        ex = _expand_sqref(sq, offsets)
        if not ex:
            continue
        new_sq, (ar, ac), (kr, kc) = ex
        b = re.sub(r"<xm:f>(.*?)</xm:f>",
                   lambda mm: f"<xm:f>{shift_formula(mm.group(1), kr - ar, kc - ac)}</xm:f>", blk, flags=re.S)
        b = re.sub(r"<xm:sqref>.*?</xm:sqref>", f"<xm:sqref>{new_sq}</xm:sqref>", b)
        b = re.sub(r'id="\{[^}]+\}"', lambda _: f'id="{{{str(uuid.uuid4()).upper()}}}"', b)
        out.append(b)

    if not out:
        return None
    head = re.match(r"<ext\b[^>]*>\s*<x14:conditionalFormattings>", ext, re.S).group(0)
    return head + "".join(out) + "</x14:conditionalFormattings></ext>"


def _rebuild_x14_dv(ext: str, offsets) -> str | None:
    items = []
    for dv in re.findall(r"<x14:dataValidation\b.*?</x14:dataValidation>", ext, re.S):
        sq = re.search(r"<xm:sqref>(.*?)</xm:sqref>", dv).group(1)
        ex = _expand_sqref(sq, offsets)
        if ex:
            items.append(re.sub(r"<xm:sqref>.*?</xm:sqref>", f"<xm:sqref>{ex[0]}</xm:sqref>", dv))
    if not items:
        return None
    head = re.match(r"<ext\b[^>]*>\s*<x14:dataValidations\b[^>]*>", ext, re.S).group(0)
    head = re.sub(r'count="\d+"', f'count="{len(items)}"', head)
    return head + "".join(items) + "</x14:dataValidations></ext>"


# ----------------------------------------------------------------------------
# pictures
# ----------------------------------------------------------------------------
def _col_left_emu(geo: Geometry, col: int) -> int:
    return sum(geo.col_px(c) for c in range(1, col)) * EMU_PER_PX


def _row_top_emu(geo: Geometry, row: int) -> int:
    return int(sum(_row_pt(geo, r) for r in range(1, row)) * EMU_PER_PT)


def _row_pt(geo: Geometry, r: int) -> float:
    if r < BLOCK_START:
        return geo.row_pt(r)
    return geo.row_pt(BLOCK_START + (r - BLOCK_START) % BLOCK_ROWS)


def _emu_to_col(geo: Geometry, x: int) -> tuple[int, int]:
    c = 1
    while True:
        w = geo.col_px(c) * EMU_PER_PX
        if x < w or c > 200:
            return c - 1, int(x)
        x -= w
        c += 1


def _emu_to_row(geo: Geometry, y: int) -> tuple[int, int]:
    r = 1
    while True:
        h = int(_row_pt(geo, r) * EMU_PER_PT)
        if y < h or r > 100000:
            return r - 1, int(y)
        y -= h
        r += 1


def _anchor(geo, top, left, box, img_w, img_h, pad_px, pid, name, rid) -> str:
    (r1, c1), (r2, c2) = box
    x0 = _col_left_emu(geo, left + c1)
    x1 = _col_left_emu(geo, left + c2 + 1)
    y0 = _row_top_emu(geo, top + r1)
    y1 = _row_top_emu(geo, top + r2 + 1)
    pad = pad_px * EMU_PER_PX
    bw, bh = (x1 - x0) - 2 * pad, (y1 - y0) - 2 * pad
    scale = min(bw / img_w, bh / img_h)
    w, h = int(img_w * scale), int(img_h * scale)
    ax, ay = x0 + (x1 - x0 - w) // 2, y0 + (y1 - y0 - h) // 2
    fc, fco = _emu_to_col(geo, ax)
    fr, fro = _emu_to_row(geo, ay)
    tc, tco = _emu_to_col(geo, ax + w)
    tr, tro = _emu_to_row(geo, ay + h)
    return (f'<xdr:twoCellAnchor editAs="oneCell">'
            f'<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>{fco}</xdr:colOff><xdr:row>{fr}</xdr:row><xdr:rowOff>{fro}</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>{tco}</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>{tro}</xdr:rowOff></xdr:to>'
            f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{pid}" name="{_x(name)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
            f'<xdr:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            f'<xdr:spPr><a:xfrm><a:off x="{ax}" y="{ay}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
            f'</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>')


def _add_pictures(pkg: Package, sheet_part: str, sx: str, cards, geo: Geometry) -> str:
    pics = []
    for i, c in enumerate(cards):
        top, left = card_origin(i)
        if c.get("barcode_png"):
            pics.append((top, left, BARCODE_BOX, c["barcode_png"], 6, f"Barcode {c.get('style', '')}"))
        if c.get("thumb_png"):
            pics.append((top, left, THUMB_BOX, c["thumb_png"], 4, f"Thumbnail {c.get('style', '')}"))
    if not pics:
        return sx
    dpath = pkg.next_name("xl/drawings/drawing{}.xml")
    media_ids = {}
    anchors = []
    pkg.ensure_default_ct("png", "image/png")
    for k, (top, left, box, png, pad, name) in enumerate(pics, start=2):
        h = hashlib.sha1(png).hexdigest()
        if h not in media_ids:
            mpath = pkg.next_name("xl/media/image{}.png")
            pkg.set_bytes(mpath, png)
            rid = pkg.add_rel(dpath, REL_IMAGE, "../media/" + mpath.split("/")[-1])
            media_ids[h] = rid
        with Image.open(io.BytesIO(png)) as im:
            iw, ih = im.size
        anchors.append(_anchor(geo, top, left, box, iw, ih, pad, k, name, media_ids[h]))
    dxml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">' + "".join(anchors) + "</xdr:wsDr>")
    pkg.set_text(dpath, dxml)
    pkg.ensure_override_ct(dpath, CT_DRAWING)
    rid = pkg.add_rel(sheet_part, REL_DRAWING, "../drawings/" + dpath.split("/")[-1])
    if 'xmlns:r="' not in sx[:2000]:
        sx = sx.replace("<worksheet ", f'<worksheet xmlns:r="{NS_R}" ', 1)
    tail_start = sx.find("</sheetData>")
    head, tail = sx[:tail_start], sx[tail_start:]
    tail = _insert(tail, f'<drawing r:id="{rid}"/>', ["legacyDrawing", "legacyDrawingHF", "drawingHF", "picture",
                                                       "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"])
    return head + tail


# ----------------------------------------------------------------------------
# workbook level
# ----------------------------------------------------------------------------
def _sheet_ref(name: str) -> str:
    return f"'{name}'" if re.search(r"[^A-Za-z0-9_]", name) else name


def _rename_sheet(pkg: Package, old: str, new: str, last_row: int) -> None:
    wb = pkg.text("xl/workbook.xml")
    sheets = re.findall(r"<sheet\b[^>]*/>", wb)
    idx = next(i for i, t in enumerate(sheets) if re.search(rf'name="{re.escape(escape(old))}"', t))
    wb = wb.replace(sheets[idx], re.sub(r'name="[^"]*"', f'name="{escape(new)}"', sheets[idx]), 1)
    pa = f"{_sheet_ref(new)}!$A$1:${'M'}${last_row}"
    pat = re.compile(rf'(<definedName name="_xlnm\.Print_Area" localSheetId="{idx}"[^>]*>)(.*?)(</definedName>)')
    if pat.search(wb):
        wb = pat.sub(lambda m: m.group(1) + escape(pa, quote=False) + m.group(3), wb)
    oldref = re.escape(_sheet_ref(old)) + "!"
    wb = re.sub(oldref, _sheet_ref(new).replace("\\", "\\\\") + "!", wb)
    pkg.set_text("xl/workbook.xml", wb)
    # formulas in other sheets that point at the old name
    for _, _, p in pkg.sheets():
        if pkg.has(p):
            t = pkg.text(p)
            if re.search(oldref, t):
                pkg.set_text(p, re.sub(oldref, _sheet_ref(new) + "!", t))
    if pkg.has("docProps/app.xml"):
        app = pkg.text("docProps/app.xml")
        app = app.replace(f"<vt:lpstr>{escape(old)}</vt:lpstr>", f"<vt:lpstr>{escape(new)}</vt:lpstr>")
        app = app.replace(f"<vt:lpstr>{escape(old)}!Print_Area</vt:lpstr>", f"<vt:lpstr>{escape(_sheet_ref(new), quote=False)}!Print_Area</vt:lpstr>")
        pkg.set_text("docProps/app.xml", app)


def _set_instructions_program(pkg: Package, program: str) -> None:
    p = pkg.sheet_path("Instructions")
    if not p:
        return
    x = pkg.text(p)
    m = re.search(r'<c r="B5"([^>]*?)(?:/>|>.*?</c>)', x, re.S)
    if not m:
        return
    s = _attr_s(m.group(1))
    x = x[:m.start()] + f'<c r="B5"{s} t="inlineStr"><is><t xml:space="preserve">{_x(program)}</t></is></c>' + x[m.end():]
    pkg.set_text(p, x)


def _add_data_sheet(pkg: Package, cards, proto, date_style) -> None:
    def ps(dr, dc):
        return _attr_s(proto.get((dr, LEFT_COL + dc), ("", None))[0])
    hdr_s = ps(2, 0)          # grey bar label style
    txt_s = ps(5, 1)
    cur_s = ps(7, 1)
    pct_s = ps(8, 4)
    imu_s = ps(9, 1)
    cols = [("#", "n", None), ("MFG STYLE #", "style", txt_s), ("PRODUCT DESCRIPTION", "description", txt_s),
            ("SUBCLASS", "program", txt_s), ("PCN ITEM ID", "pcn_item_id", txt_s), ("SPARK PID", "spark_pid", txt_s),
            ("APEX ID", "apex_id", txt_s), ("DPCI", "dpci", txt_s), ("FIRST COST", "first_cost", cur_s),
            ("COST BASIS", "cost_basis", txt_s), ("PCN RETAIL", "retail", cur_s), ("COUNTRY OF EXPORT", "country", txt_s),
            ("DUTY RATE", "duty", pct_s), ("IMU", "imu", imu_s), ("IMU w/DUTY", "imu_duty", imu_s),
            ("IMU GOAL", "goal", pct_s), ("SAMPLE DATE", "sample_date", f' s="{date_style}"'),
            ("LAST YEAR DPCI", "ly_dpci", txt_s), ("LY FIRST COST", "ly_cost", cur_s), ("LY RETAIL", "ly_retail", cur_s),
            ("LY DUTY RATE", "ly_duty", pct_s), ("COMMENTS", "comments", txt_s)]
    widths = [5, 16, 48, 18, 13, 13, 11, 14, 12, 11, 12, 18, 11, 10, 11, 10, 13, 15, 13, 12, 12, 40]
    rows = []
    hdr = "".join(f'<c r="{ref(1, j + 1)}"{hdr_s} t="inlineStr"><is><t>{_x(h)}</t></is></c>' for j, (h, _, _) in enumerate(cols))
    rows.append(f'<row r="1" ht="30" customHeight="1">{hdr}</row>')
    for i, c in enumerate(cards, start=1):
        r = i + 1
        cells = []
        for j, (_, key, s) in enumerate(cols):
            a1 = ref(r, j + 1)
            if key == "n":
                cells.append(f'<c r="{a1}"><v>{i}</v></c>')
                continue
            if key in ("imu", "imu_duty"):
                v = imu_value(c.get("first_cost"), c.get("retail"), c.get("duty"), key == "imu_duty")
            elif key == "sample_date":
                v = excel_serial(c.get("sample_date"))
            else:
                v = c.get(key)
            if _is_blank(v):
                cells.append(f'<c r="{a1}"{s or ""}/>')
            elif isinstance(v, (int, float)) and key not in ("pcn_item_id",):
                cells.append(f'<c r="{a1}"{s or ""}><v>{_num(float(v))}</v></c>')
            else:
                cells.append(f'<c r="{a1}"{s or ""} t="inlineStr"><is><t xml:space="preserve">{_x(v)}</t></is></c>')
        rows.append(f'<row r="{r}">{"".join(cells)}</row>')
    last = ref(len(cards) + 1, len(cols))
    colxml = "".join(f'<col min="{j + 1}" max="{j + 1}" width="{w}" customWidth="1"/>' for j, w in enumerate(widths))
    x = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
         f'xmlns:r="{NS_R}"><dimension ref="A1:{last}"/>'
         '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
         '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
         '<sheetFormatPr defaultRowHeight="15"/>'
         f"<cols>{colxml}</cols><sheetData>{''.join(rows)}</sheetData>"
         f'<autoFilter ref="A1:{last}"/>'
         '<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>')
    path = pkg.next_name("xl/worksheets/sheet{}.xml")
    pkg.set_text(path, x)
    pkg.ensure_override_ct(path, CT_SHEET)
    rid = pkg.add_rel("xl/workbook.xml", REL_SHEET, "worksheets/" + path.split("/")[-1])
    wb = pkg.text("xl/workbook.xml")
    ids = [int(i) for i in re.findall(r'sheetId="(\d+)"', wb)]
    existing = {n.lower() for n, _, _ in pkg.sheets()}
    data_name = "Data" if "data" not in existing else "Tag Data"
    new_sheet = f'<sheet name="{data_name}" sheetId="{max(ids) + 1}" r:id="{rid}"/>'
    # place right after the Sample Tag sheet
    st = re.search(r'<sheet\b[^>]*name="Sample Tag"[^>]*/>', wb)
    wb = wb[:st.end()] + new_sheet + wb[st.end():] if st else wb.replace("</sheets>", new_sheet + "</sheets>")
    # localSheetId of defined names after the insertion point shift by one
    pos = [m.start() for m in re.finditer(r"<sheet\b", wb)]
    data_idx = next(i for i, p in enumerate(pos) if wb[p:].startswith(new_sheet))

    def bump(m):
        k = int(m.group(1))
        return f'localSheetId="{k + 1 if k >= data_idx else k}"'
    wb = re.sub(r'localSheetId="(\d+)"', bump, wb)
    # filter database name for the Data sheet's autofilter
    dn = (f'<definedName name="_xlnm._FilterDatabase" localSheetId="{data_idx}" hidden="1">'
          f'{escape(_sheet_ref(data_name), quote=False)}!$A$1:${re.sub(r"[0-9]", "", last)}${len(cards) + 1}</definedName>')
    if "<definedNames>" in wb:
        wb = wb.replace("<definedNames>", "<definedNames>" + dn, 1)
    else:
        wb = wb.replace("</sheets>", "</sheets><definedNames>" + dn + "</definedNames>", 1)
    pkg.set_text("xl/workbook.xml", wb)
    if pkg.has("docProps/app.xml"):
        app = pkg.text("docProps/app.xml")
        app = app.replace("<vt:lpstr>Sample Tag</vt:lpstr>", f"<vt:lpstr>Sample Tag</vt:lpstr><vt:lpstr>{data_name}</vt:lpstr>", 1)
        app = re.sub(r"(<vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)",
                     lambda m: m.group(1) + str(int(m.group(2)) + 1), app)
        app = re.sub(r'(<TitlesOfParts><vt:vector size=")(\d+)', lambda m: m.group(1) + str(int(m.group(2)) + 1), app)
        pkg.set_text("docProps/app.xml", app)
