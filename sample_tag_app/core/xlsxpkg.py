"""Minimal .xlsx/.xlsm package helper for XML surgery: every part is kept as raw
bytes and only the parts we explicitly change are rewritten, so VBA projects,
x14 extensions, printer settings, custom XML etc. survive untouched."""
from __future__ import annotations

import io
import posixpath
import re
import zipfile
from html import unescape


class Package:
    def __init__(self, data: bytes):
        self._order: list[str] = []
        self.parts: dict[str, bytes] = {}
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for info in z.infolist():
                self._order.append(info.filename)
                self.parts[info.filename] = z.read(info.filename)

    # ---- basic access ----------------------------------------------------
    def has(self, name: str) -> bool:
        return name in self.parts

    def text(self, name: str) -> str:
        return self.parts[name].decode("utf-8")

    def set_text(self, name: str, text: str) -> None:
        if name not in self.parts:
            self._order.append(name)
        self.parts[name] = text.encode("utf-8")

    def set_bytes(self, name: str, data: bytes) -> None:
        if name not in self.parts:
            self._order.append(name)
        self.parts[name] = data

    def remove(self, name: str) -> None:
        if name in self.parts:
            del self.parts[name]
            self._order.remove(name)

    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            # [Content_Types].xml first, as Excel writes it
            names = ["[Content_Types].xml"] + [n for n in self._order if n != "[Content_Types].xml"]
            for n in names:
                if n in self.parts:
                    z.writestr(n, self.parts[n])
        return buf.getvalue()

    # ---- workbook structure ---------------------------------------------
    def sheets(self) -> list[tuple[str, str, str]]:
        """[(name, rId, part path)] in workbook order."""
        wb = self.text("xl/workbook.xml")
        rels = self.text("xl/_rels/workbook.xml.rels")
        relmap = {rid: tgt for rid, tgt in re.findall(r'<Relationship[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)}
        relmap.update({rid: tgt for tgt, rid in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)})
        out = []
        for m in re.finditer(r"<sheet\b[^>]*/>", wb):
            tag = m.group(0)
            name = unescape(re.search(r'name="([^"]*)"', tag).group(1))
            rid = re.search(r'r:id="([^"]+)"', tag).group(1)
            tgt = relmap.get(rid, "")
            path = posixpath.normpath(posixpath.join("xl", tgt)) if not tgt.startswith("/") else tgt.lstrip("/")
            out.append((name, rid, path))
        return out

    def sheet_path(self, name: str) -> str | None:
        for n, _, p in self.sheets():
            if n.strip().lower() == name.strip().lower():
                return p
        return None

    def rels_path(self, part: str) -> str:
        d, f = posixpath.split(part)
        return posixpath.join(d, "_rels", f + ".rels")

    def shared_strings(self) -> list[str]:
        if not self.has("xl/sharedStrings.xml"):
            return []
        x = self.text("xl/sharedStrings.xml")
        out = []
        for si in re.findall(r"<si>(.*?)</si>", x, re.S):
            ts = re.findall(r"<t(?:\s[^>]*)?>(.*?)</t>", si, re.S)
            # skip phonetic runs
            si_np = re.sub(r"<rPh.*?</rPh>", "", si, flags=re.S)
            ts = re.findall(r"<t(?:\s[^>]*)?>(.*?)</t>", si_np, re.S)
            out.append(unescape("".join(ts)))
        return out

    # ---- calc ---------------------------------------------------------------
    def drop_calc_chain(self) -> None:
        if self.has("xl/calcChain.xml"):
            self.remove("xl/calcChain.xml")
        ct = self.text("[Content_Types].xml")
        ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
        self.set_text("[Content_Types].xml", ct)
        rels = self.text("xl/_rels/workbook.xml.rels")
        rels = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', "", rels)
        self.set_text("xl/_rels/workbook.xml.rels", rels)

    def force_full_calc(self) -> None:
        wb = self.text("xl/workbook.xml")
        if "<calcPr" in wb:
            if "fullCalcOnLoad" not in wb:
                wb = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wb, count=1)
        else:
            wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
        self.set_text("xl/workbook.xml", wb)

    # ---- content types / relationships -----------------------------------
    def ensure_default_ct(self, ext: str, ctype: str) -> None:
        ct = self.text("[Content_Types].xml")
        if f'Extension="{ext}"' not in ct and f"Extension=\"{ext.upper()}\"" not in ct:
            ct = ct.replace("<Types ", "<Types ", 1)
            ct = re.sub(r"(<Types[^>]*>)", rf'\1<Default Extension="{ext}" ContentType="{ctype}"/>', ct, count=1)
            self.set_text("[Content_Types].xml", ct)

    def ensure_override_ct(self, part: str, ctype: str) -> None:
        ct = self.text("[Content_Types].xml")
        pn = "/" + part.lstrip("/")
        if f'PartName="{pn}"' not in ct:
            ct = ct.replace("</Types>", f'<Override PartName="{pn}" ContentType="{ctype}"/></Types>')
            self.set_text("[Content_Types].xml", ct)

    def add_rel(self, part: str, rtype: str, target: str) -> str:
        rp = self.rels_path(part)
        if self.has(rp):
            x = self.text(rp)
        else:
            x = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', x)]
        rid = f"rId{(max(ids) + 1) if ids else 1}"
        x = x.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/></Relationships>')
        self.set_text(rp, x)
        return rid

    def next_name(self, pattern: str) -> str:
        """pattern like 'xl/drawings/drawing{}.xml' -> first free number."""
        i = 1
        while self.has(pattern.format(i)):
            i += 1
        return pattern.format(i)
