"""Tiny editor for xl/styles.xml: clone a cell style (xf) with a different fill,
wrap or number format, appending only what is new. Existing indices never move,
so every untouched cell keeps its exact template look."""
from __future__ import annotations

import re


class StyleBook:
    def __init__(self, xml: str):
        self.xml = xml
        cx = re.search(r"<cellXfs count=\"\d+\">(.*?)</cellXfs>", xml, re.S)
        self.xfs = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", cx.group(1), re.S)
        fl = re.search(r"<fills count=\"\d+\">(.*?)</fills>", xml, re.S)
        self.fills = re.findall(r"<fill>.*?</fill>|<fill/>", fl.group(1), re.S) if fl else []
        nf = re.search(r"<numFmts count=\"\d+\">(.*?)</numFmts>", xml, re.S)
        self.numfmts = re.findall(r"<numFmt\b[^>]*/>", nf.group(1)) if nf else []
        dx = re.search(r"<dxfs count=\"\d+\">(.*?)</dxfs>", xml, re.S)
        self.dxfs = re.findall(r"<dxf>.*?</dxf>|<dxf/>", dx.group(1), re.S) if dx else []
        fo = re.search(r"<fonts count=\"\d+\"[^>]*>(.*?)</fonts>", xml, re.S)
        self.fonts = re.findall(r"<font>.*?</font>|<font/>", fo.group(1), re.S) if fo else []
        self._cache: dict[tuple, int] = {}

    # ---- pieces ----------------------------------------------------------------
    def numfmt_id(self, code: str) -> int:
        for f in self.numfmts:
            if f'formatCode="{code}"' in f:
                return int(re.search(r'numFmtId="(\d+)"', f).group(1))
        ids = [int(re.search(r'numFmtId="(\d+)"', f).group(1)) for f in self.numfmts]
        new_id = max(ids + [199]) + 1
        self.numfmts.append(f'<numFmt numFmtId="{new_id}" formatCode="{code}"/>')
        return new_id

    def fill_id(self, rgb: str) -> int:
        x = f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb.upper()}"/><bgColor indexed="64"/></patternFill></fill>'
        if x in self.fills:
            return self.fills.index(x)
        self.fills.append(x)
        return len(self.fills) - 1

    def font_id(self, base: int, size: float) -> int:
        """Copy font `base` with a new size (bold / colour / name kept)."""
        f = self.fonts[base] if base < len(self.fonts) else "<font/>"
        sz = f'<sz val="{size:g}"/>'
        if f == "<font/>":
            f = f"<font>{sz}</font>"
        elif "<sz " in f:
            f = re.sub(r'<sz val="[^"]*"/>', sz, f, count=1)
        else:
            f = f.replace("<font>", "<font>" + sz, 1)
        # fitted cells use Calibri: its metrics are known (Carlito twin), so the size we
        # compute is exactly what Excel draws. The template's Aptos is wider and not bundled.
        f = re.sub(r'<name val="[^"]*"/>', '<name val="Calibri"/>', f)
        f = re.sub(r'<scheme val="[^"]*"/>', "", f)
        if "<name " not in f:
            f = f.replace("</font>", '<name val="Calibri"/><family val="2"/></font>')
        if f in self.fonts:
            return self.fonts.index(f)
        self.fonts.append(f)
        return len(self.fonts) - 1

    def font_size(self, idx: int) -> float:
        m = re.search(r'fontId="(\d+)"', self.xfs[idx])
        f = self.fonts[int(m.group(1))] if m and int(m.group(1)) < len(self.fonts) else ""
        s = re.search(r'<sz val="([0-9.]+)"/>', f)
        return float(s.group(1)) if s else 11.0

    def add_dxf(self, dxf_xml: str) -> int:
        if dxf_xml in self.dxfs:
            return self.dxfs.index(dxf_xml)
        self.dxfs.append(dxf_xml)
        return len(self.dxfs) - 1

    # ---- xf clone ------------------------------------------------------------
    def clone(self, idx: int, *, fill_rgb: str | None = None, wrap: bool | None = None,
              numfmt: str | None = None, font_size: float | None = None) -> int:
        key = (idx, fill_rgb, wrap, numfmt, font_size)
        if key in self._cache:
            return self._cache[key]
        xf = self.xfs[idx]
        if fill_rgb:
            fid = self.fill_id(fill_rgb)
            xf = self._set_attr(xf, "fillId", str(fid))
            xf = self._set_attr(xf, "applyFill", "1")
        if font_size:
            m = re.search(r'fontId="(\d+)"', xf)
            xf = self._set_attr(xf, "fontId", str(self.font_id(int(m.group(1)) if m else 0, font_size)))
            xf = self._set_attr(xf, "applyFont", "1")
        if numfmt:
            nid = self.numfmt_id(numfmt)
            xf = self._set_attr(xf, "numFmtId", str(nid))
            xf = self._set_attr(xf, "applyNumberFormat", "1")
        if wrap is not None:
            if xf.endswith("/>"):
                xf = xf[:-2] + "><alignment/></xf>"
            if "<alignment" not in xf:
                xf = xf.replace("</xf>", "<alignment/></xf>", 1)
            def al(m):
                t = re.sub(r'\swrapText="[^"]*"', "", m.group(0))
                return t.replace("<alignment", f'<alignment wrapText="{1 if wrap else 0}"', 1)
            xf = re.sub(r"<alignment\b[^>]*/?>", al, xf, count=1)
            xf = self._set_attr(xf, "applyAlignment", "1")
        self.xfs.append(xf)
        self._cache[key] = len(self.xfs) - 1
        return self._cache[key]

    @staticmethod
    def _set_attr(xf: str, name: str, value: str) -> str:
        head = re.match(r"<xf\b[^>]*?(?=/?>)", xf).group(0)
        if re.search(rf'\s{name}="[^"]*"', head):
            new_head = re.sub(rf'\s{name}="[^"]*"', f' {name}="{value}"', head)
        else:
            new_head = head + f' {name}="{value}"'
        return new_head + xf[len(head):]

    # ---- write back -------------------------------------------------------------
    def to_xml(self) -> str:
        x = self.xml
        x = re.sub(r"<cellXfs count=\"\d+\">.*?</cellXfs>",
                   lambda _: f'<cellXfs count="{len(self.xfs)}">' + "".join(self.xfs) + "</cellXfs>", x, count=1, flags=re.S)
        if re.search(r"<fills count=", x):
            x = re.sub(r"<fills count=\"\d+\">.*?</fills>",
                       lambda _: f'<fills count="{len(self.fills)}">' + "".join(self.fills) + "</fills>", x, count=1, flags=re.S)
        if self.fonts:
            x = re.sub(r"<fonts count=\"\d+\"([^>]*)>.*?</fonts>",
                       lambda m: f'<fonts count="{len(self.fonts)}"{m.group(1)}>' + "".join(self.fonts) + "</fonts>",
                       x, count=1, flags=re.S)
        if self.dxfs:
            block = f'<dxfs count="{len(self.dxfs)}">' + "".join(self.dxfs) + "</dxfs>"
            if re.search(r"<dxfs\b", x):
                x = re.sub(r"<dxfs count=\"\d+\">.*?</dxfs>|<dxfs count=\"0\"/>", lambda _: block, x, count=1, flags=re.S)
            else:
                x = x.replace("<tableStyles", block + "<tableStyles", 1) if "<tableStyles" in x else x.replace("</styleSheet>", block + "</styleSheet>")
        if self.numfmts:
            block = f'<numFmts count="{len(self.numfmts)}">' + "".join(self.numfmts) + "</numFmts>"
            if "<numFmts" in x:
                x = re.sub(r"<numFmts count=\"\d+\">.*?</numFmts>", lambda _: block, x, count=1, flags=re.S)
            else:
                x = re.sub(r"(<styleSheet\b[^>]*>)", lambda m: m.group(1) + block, x, count=1)
        return x
