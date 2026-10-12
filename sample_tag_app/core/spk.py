"""Product thumbnails from an SPK Workspace export.

Works with both ways Excel stores pictures:
  * floating pictures anchored on a row (xl/drawings/…), and
  * "Place in cell" pictures (xl/richData/…, cell attribute vm="n").
Each picture is tied to its row, and the row is identified by whichever key
columns the sheet has (DPCI, Spark PID / Product ID, MFG Style, PCN Item ID)."""
from __future__ import annotations

import io
import posixpath
import re
from html import unescape

from openpyxl.utils import column_index_from_string
from PIL import Image

from .xlsxpkg import Package

KEY_HEADERS = {
    "dpci": ["dpci"],
    "spark_pid": ["spark pid", "product id", "pid"],
    "style": ["manufacturer style #", "mfg style #", "vendor style #", "manufacturer style", "style #", "style"],
    "pcn_item_id": ["pcn item id"],
}


def _norm(s: str) -> str:
    s = re.sub(r"\s*\*+\s*$", "", (s or "").strip())
    return re.sub(r"\s+", " ", s).lower()


def _clean(v: str) -> str:
    v = (v or "").strip()
    return v[:-2] if re.fullmatch(r"\d+\.0", v) else v


def _cell_values(pkg: Package, sheet_xml: str) -> dict[tuple[int, int], str]:
    sst = pkg.shared_strings()
    out = {}
    for m in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', sheet_xml, re.S):
        col, row, attrs, inner = m.group(1), int(m.group(2)), m.group(3), m.group(4) or ""
        v = re.search(r"<v>(.*?)</v>", inner, re.S)
        t = re.search(r't="(\w+)"', attrs)
        val = ""
        if 's' == (t.group(1) if t else "") and v:
            i = int(v.group(1))
            val = sst[i] if i < len(sst) else ""
        elif t and t.group(1) == "inlineStr":
            val = unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S)))
        elif v:
            val = unescape(v.group(1))
        if val:
            out[(row, column_index_from_string(col))] = val
    return out


def _rel_targets(pkg: Package, part: str) -> dict[str, str]:
    rp = pkg.rels_path(part)
    if not pkg.has(rp):
        return {}
    base = posixpath.dirname(part)
    out = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", pkg.text(rp)):
        rid = re.search(r'Id="([^"]+)"', tag).group(1)
        tgt = re.search(r'Target="([^"]+)"', tag).group(1)
        out[rid] = posixpath.normpath(posixpath.join(base, tgt)) if not tgt.startswith("/") else tgt.lstrip("/")
    return out


def _anchored_images(pkg: Package, sheet_part: str) -> list[tuple[int, bytes]]:
    """(1-based row, image bytes) for floating pictures."""
    out = []
    for rid, dpath in _rel_targets(pkg, sheet_part).items():
        if "/drawings/" not in dpath or not pkg.has(dpath):
            continue
        dx = pkg.text(dpath)
        drels = _rel_targets(pkg, dpath)
        p = r"(?:\w+:)?"          # namespace prefix varies by writer (xdr: or none)
        for anc in re.findall(rf"<{p}(?:twoCellAnchor|oneCellAnchor)\b.*?</{p}(?:twoCellAnchor|oneCellAnchor)>", dx, re.S):
            fr = re.search(rf"<{p}from>.*?<{p}row>(\d+)</{p}row>.*?</{p}from>", anc, re.S)
            blip = re.search(r'r:embed="([^"]+)"', anc)
            if fr and blip and blip.group(1) in drels and pkg.has(drels[blip.group(1)]):
                out.append((int(fr.group(1)) + 1, pkg.parts[drels[blip.group(1)]]))
    return out


def _incell_images(pkg: Package, sheet_xml: str) -> list[tuple[int, bytes]]:
    """(row, image) for 'Place in cell' pictures (richData)."""
    if not pkg.has("xl/metadata.xml") or not pkg.has("xl/richData/rdrichvalue.xml"):
        return []
    meta = pkg.text("xl/metadata.xml")
    # valueMetadata block i -> rc v -> futureMetadata XLRICHVALUE bk index -> rvb i
    vmb = re.search(r"<valueMetadata[^>]*>(.*?)</valueMetadata>", meta, re.S)
    vm_blocks = re.findall(r"<bk>(.*?)</bk>", vmb.group(1), re.S) if vmb else []
    fut = re.search(r'<futureMetadata name="XLRICHVALUE"[^>]*>(.*?)</futureMetadata>', meta, re.S)
    fut_rv = [int(x) for x in re.findall(r'<xlrd:rvb i="(\d+)"', fut.group(1))] if fut else []
    vm_to_rv = {}
    for i, bk in enumerate(vm_blocks, start=1):
        v = re.search(r'<rc t="\d+" v="(\d+)"', bk)
        if v and int(v.group(1)) < len(fut_rv):
            vm_to_rv[i] = fut_rv[int(v.group(1))]
    rv = pkg.text("xl/richData/rdrichvalue.xml")
    rvs = re.findall(r"<rv\b[^>]*>(.*?)</rv>", rv, re.S)
    rel_idx = []
    for r in rvs:
        vs = re.findall(r"<v[^>]*>(.*?)</v>", r)
        rel_idx.append(int(vs[0]) if vs and vs[0].isdigit() else None)
    if not pkg.has("xl/richData/richValueRel.xml"):
        return []
    rvrel = pkg.text("xl/richData/richValueRel.xml")
    rids = re.findall(r'<rel r:id="([^"]+)"', rvrel)
    targets = _rel_targets(pkg, "xl/richData/richValueRel.xml")
    out = []
    for m in re.finditer(r'<c r="[A-Z]+(\d+)"[^>]*\bvm="(\d+)"', sheet_xml):
        row, vm = int(m.group(1)), int(m.group(2))
        rvi = vm_to_rv.get(vm)
        if rvi is None or rvi >= len(rel_idx) or rel_idx[rvi] is None:
            continue
        ri = rel_idx[rvi]
        if ri < len(rids) and rids[ri] in targets and pkg.has(targets[rids[ri]]):
            out.append((row, pkg.parts[targets[rids[ri]]]))
    return out


def extract(data: bytes) -> tuple[dict[str, dict[str, bytes]], list[str]]:
    """-> ({key_type: {key_value: png_bytes}}, notes)."""
    pkg = Package(data)
    result: dict[str, dict[str, bytes]] = {k: {} for k in KEY_HEADERS}
    notes = []
    sheets = pkg.sheets()
    # Prefer a "Products" sheet; Claims sheets carry defect photos, not product shots.
    ordered = sorted(sheets, key=lambda s: (0 if "product" in s[0].lower() else 1))
    used = None
    for name, _, part in ordered:
        if "claim" in name.lower() or not pkg.has(part):
            continue
        xml = pkg.text(part)
        imgs = _anchored_images(pkg, part) + _incell_images(pkg, xml)
        if not imgs:
            continue
        vals = _cell_values(pkg, xml)
        # header row = first row (within 15) that has any key header
        hdr_row, keycols = None, {}
        for r in range(1, 16):
            found = {}
            for (rr, cc), v in vals.items():
                if rr != r:
                    continue
                nv = _norm(v)
                for k, opts in KEY_HEADERS.items():
                    if nv in opts and k not in found:
                        found[k] = cc
            if found:
                hdr_row, keycols = r, found
                break
        if not keycols:
            notes.append(f"「{name}」有圖片但找不到 DPCI / PID / Style 欄位，已略過。")
            continue
        n = 0
        for row, blob in imgs:
            if row <= hdr_row:
                continue
            png = _to_png(blob)
            if png is None:
                continue
            for k, cc in keycols.items():
                v = _clean(vals.get((row, cc), ""))
                if v:
                    result[k].setdefault(v.upper(), png)
            n += 1
        notes.append(f"SPK「{name}」：取得 {n} 張縮圖（對應欄位：{', '.join(keycols)}）。")
        used = name
        break
    if used is None:
        notes.append("SPK 檔案中沒有找到產品縮圖。")
    return result, notes


def _to_png(blob: bytes) -> bytes | None:
    """SPK pictures get the same treatment as uploaded photos: rotation fixed,
    white / transparent margins trimmed, shrunk to the thumbnail size."""
    from .images import prepare
    res = prepare(blob, trim=True)
    return res[0] if res else None


def lookup(images: dict[str, dict[str, bytes]], item: dict) -> bytes | None:
    """Match one card item: DPCI -> Spark PID -> MFG Style -> PCN Item ID."""
    for k in ("dpci", "spark_pid", "style", "pcn_item_id"):
        v = str(item.get(k) or "").strip().upper()
        if v and v in images.get(k, {}):
            return images[k][v]
    return None
