"""Read the Target template's Dropdowns sheet (programs + IMU goals, the yellow
band, vendor list and vendor header colours), compare two templates, and write
a refreshed Dropdowns sheet into a working template by XML surgery."""
from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field

import openpyxl

from .colors import parse_theme, resolve_xml_color
from .xlsxpkg import Package


def _s(v) -> str:
    return "" if v is None else str(v).strip()


def _key(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


@dataclass
class TemplateInfo:
    programs: dict[str, dict] = field(default_factory=dict)   # name -> {"imu":x, "imu_duty":y, "row":n}
    yellow_band: float | None = None
    vendors: list[str] = field(default_factory=list)
    vendor_rules: list = field(default_factory=list)   # [(priority, text, RRGGBB)] from the header CF
    vendor_default: str = ""
    program_default: str = ""
    theme: list[str] = field(default_factory=list)
    imu_header: str = "IMU Goal"
    imu_duty_header: str = "IMU+Duty Goal"

    def goal(self, program: str, basis: str = "imu") -> float | None:
        p = self.programs.get(program)
        if p is None:
            k = _key(program)
            for name, v in self.programs.items():
                if _key(name) == k:
                    p = v
                    break
        if not p:
            return None
        return p.get(basis)

    def vendor_color(self, vendor: str) -> str | None:
        """Header colour exactly as Excel shows it: the template colours the
        vendor bar with 'cell equals' rules, which compare case-insensitively
        but keep spaces ("Lucky Star " != "Lucky Star"). Highest-priority match wins."""
        v = (vendor or "").lower()
        hits = [r for r in self.vendor_rules if r[1].lower() == v]
        return min(hits)[2] if hits else None


def _find_sheet(wb, name):
    for ws in wb.worksheets:
        if ws.title.strip().lower() == name.lower():
            return ws
    return None


def parse_template(data: bytes) -> TemplateInfo:
    info = TemplateInfo()
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    pkg = Package(data)
    info.theme = parse_theme(pkg.text("xl/theme/theme1.xml") if pkg.has("xl/theme/theme1.xml") else None)

    ins = _find_sheet(wb, "Instructions")
    if ins is not None:
        info.vendor_default = "" if ins["B4"].value is None else str(ins["B4"].value)   # raw: Excel colour rules are space-sensitive
        info.program_default = _s(ins["B5"].value)

    dd = _find_sheet(wb, "Dropdowns")
    if dd is None:
        return info

    # ---- locate header cells -------------------------------------------
    prog_hdr = vend_hdr = None
    yellow = None
    for row in dd.iter_rows(min_row=1, max_row=min(dd.max_row, 40)):
        for c in row:
            v = _s(c.value).lower()
            if v == "program name" and prog_hdr is None:
                prog_hdr = c
            elif v == "vendor name" and vend_hdr is None:
                vend_hdr = c
            elif v == "yellow" and yellow is None:
                yellow = c

    # ---- programs + IMU goals -----------------------------------------
    if prog_hdr is not None:
        r0, c0 = prog_hdr.row, prog_hdr.column
        goal_cols = {}
        for dc in (1, 2, 3):
            h = _s(dd.cell(r0, c0 + dc).value)
            if "imu" in h.lower():
                if "duty" in h.lower():
                    goal_cols["imu_duty"] = c0 + dc
                    info.imu_duty_header = h
                else:
                    goal_cols["imu"] = c0 + dc
                    info.imu_header = h
        r = r0 + 1
        while r <= dd.max_row:
            name = _s(dd.cell(r, c0).value)
            if not name:
                break
            rec = {"row": r}
            for k, cc in goal_cols.items():
                try:
                    rec[k] = float(dd.cell(r, cc).value)
                except (TypeError, ValueError):
                    rec[k] = None
            info.programs[name] = rec
            r += 1

    # ---- yellow band: "=M6-0.03" next to the "Yellow" label ------------
    if yellow is not None:
        for dc in (1, 2):
            f = _s(dd.cell(yellow.row, yellow.column + dc).value)
            m = re.search(r"-\s*([0-9.]+)\s*$", f)
            if f.startswith("=") and m and "E" not in m.group(1).upper():
                try:
                    band = float(m.group(1))
                    if 0 < band < 0.5:
                        info.yellow_band = band
                        break
                except ValueError:
                    pass

    # ---- vendor list ------------------------------------------------------
    vend_rows = {}
    if vend_hdr is not None:
        for r in range(vend_hdr.row + 1, dd.max_row + 1):
            raw = dd.cell(r, vend_hdr.column).value
            v = _s(raw)
            if v and "!" not in v:
                info.vendors.append(v)
                vend_rows[r] = str(raw)

    # ---- vendor header colours from the Template sheet's CF --------------
    info.vendor_rules = _vendor_rules(pkg, info.theme, vend_rows, vend_hdr.column if vend_hdr else 4)
    return info


def _vendor_rules(pkg: Package, theme, vend_rows: dict, vend_col: int) -> list:
    out: list = []
    sheet_path = pkg.sheet_path("Template") or pkg.sheet_path("Sample Tag")
    if not sheet_path:
        return out
    x = pkg.text(sheet_path)
    styles = pkg.text("xl/styles.xml")
    dm = re.search(r"<dxfs[^>]*>(.*?)</dxfs>", styles, re.S)
    dxfs = re.findall(r"<dxf>.*?</dxf>|<dxf/>", dm.group(1), re.S) if dm else []
    col_letter = openpyxl.utils.get_column_letter(vend_col)

    # literal rules: <cfRule ... dxfId="n" operator="equal"><formula>"Lucky Star"</formula>
    m = re.search(r'<conditionalFormatting sqref="B2:F2">(.*?)</conditionalFormatting>', x, re.S)
    if m:
        for rule in re.findall(r"<cfRule\b[^>]*>.*?</cfRule>", m.group(1), re.S):
            did = re.search(r'dxfId="(\d+)"', rule)
            pr = re.search(r'priority="(\d+)"', rule)
            lit = re.search(r'<formula>"([^"]*)"</formula>', rule)
            if not (did and pr and lit and 'operator="equal"' in rule):
                continue
            i = int(did.group(1))
            if i < len(dxfs):
                bg = re.search(r"<bgColor ([^/]*)/>", dxfs[i])
                col = resolve_xml_color(bg.group(1), theme) if bg else None
                if col:
                    out.append((int(pr.group(1)), lit.group(1), col))
    # x14 rules referencing Dropdowns!$D$n
    for cf in re.findall(r"<x14:conditionalFormatting.*?</x14:conditionalFormatting>", x, re.S):
        sq = re.search(r"<xm:sqref>(.*?)</xm:sqref>", cf)
        if not sq or sq.group(1).split()[0] != "B2:F2":
            continue
        for rule in re.findall(r"<x14:cfRule.*?</x14:cfRule>", cf, re.S):
            ref = re.search(rf"<xm:f>Dropdowns!\$?{col_letter}\$?(\d+)</xm:f>", rule)
            bg = re.search(r"<bgColor ([^/]*)/>", rule)
            pr = re.search(r'priority="(\d+)"', rule)
            if ref and bg and pr:
                vendor = vend_rows.get(int(ref.group(1)))
                col = resolve_xml_color(bg.group(1), theme)
                if vendor and col:
                    out.append((int(pr.group(1)), vendor, col))
    return out


# --------------------------------------------------------------------------
# Compare + update (the quarterly "跑 Dropdown 更新")
# --------------------------------------------------------------------------
def diff_templates(old: TemplateInfo, new: TemplateInfo) -> dict:
    o, n = old.programs, new.programs
    ok = {_key(k): k for k in o}
    nk = {_key(k): k for k in n}
    added = [nk[k] for k in nk if k not in ok]
    removed = [ok[k] for k in ok if k not in nk]
    changed = []
    for k in nk:
        if k in ok:
            a, b = o[ok[k]], n[nk[k]]
            for basis in ("imu", "imu_duty"):
                if (a.get(basis) or 0) != (b.get(basis) or 0):
                    changed.append((nk[k], basis, a.get(basis), b.get(basis)))
    ov = {_key(v) for v in old.vendors}
    nv = {_key(v) for v in new.vendors}
    return {
        "programs_added": added,
        "programs_removed": removed,
        "goals_changed": changed,
        "vendors_added": [v for v in new.vendors if _key(v) not in ov],
        "vendors_removed": [v for v in old.vendors if _key(v) not in nv],
        "band_old": old.yellow_band,
        "band_new": new.yellow_band,
    }


def update_dropdowns(template_bytes: bytes, source_bytes: bytes) -> bytes:
    """Replace the TEMPLATE's Dropdowns sheet *content* with the SOURCE's,
    by XML surgery: only the Dropdowns sheet's <sheetData>/<mergeCells> and the
    strings it needs change. VBA, x14 validations/CF, images and every other
    sheet are copied byte-for-byte."""
    tpl = Package(template_bytes)
    src = Package(source_bytes)
    tp = tpl.sheet_path("Dropdowns")
    sp = src.sheet_path("Dropdowns")
    if not tp or not sp:
        raise ValueError("TEMPLATE 或 SOURCE 找不到 Dropdowns 工作表")

    src_xml = src.text(sp)
    src_sst = src.shared_strings()
    tpl_xml = tpl.text(tp)

    src_data = re.search(r"<sheetData>.*?</sheetData>|<sheetData/>", src_xml, re.S).group(0)

    # 1) shared-string cells -> inline strings, so the template's sst is untouched
    def to_inline(m):
        attrs, idx = m.group(1), int(m.group(2))
        attrs = re.sub(r'\st="s"', "", attrs)
        txt = src_sst[idx] if idx < len(src_sst) else ""
        txt = txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return f'<c{attrs} t="inlineStr"><is><t xml:space="preserve">{txt}</t></is></c>'
    data = re.sub(r'<c([^>]*?\st="s"[^>]*)><v>(\d+)</v></c>', to_inline, src_data)

    # 2) style indices: map each source cell's style to the template cell style at the
    #    same address when one exists (keeps the template's look); otherwise drop it.
    tpl_styles = dict(re.findall(r'<c r="([A-Z]+\d+)"[^>]*?\ss="(\d+)"', tpl_xml))
    def restyle(m):
        ref, rest = m.group(1), m.group(2)
        rest = re.sub(r'\ss="\d+"', "", rest)
        s = tpl_styles.get(ref)
        return f'<c r="{ref}"' + (f' s="{s}"' if s else "") + rest
    data = re.sub(r'<c r="([A-Z]+\d+)"((?:\s[^>]*?)?)(?=/?>)', restyle, data)
    # row-level styles from the source would point into the wrong styles table
    data = re.sub(r'(<row[^>]*?)\ss="\d+"(?:\s+customFormat="1")?', r"\1", data)
    # cached values of formulas are recalculated by Excel
    tpl_xml = re.sub(r"<sheetData>.*?</sheetData>|<sheetData/>", lambda _: data, tpl_xml, count=1, flags=re.S)

    # 3) dimension follows the new data
    dim = re.search(r'<dimension ref="[^"]+"/>', src_xml)
    if dim:
        tpl_xml = re.sub(r'<dimension ref="[^"]+"/>', dim.group(0), tpl_xml, count=1)

    tpl.set_text(tp, tpl_xml)
    tpl.drop_calc_chain()
    tpl.force_full_calc()
    return tpl.to_bytes()
