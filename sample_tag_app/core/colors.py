"""Excel colour helpers: theme palette + tint, so the PNG renderer shows the same
colours Excel does (e.g. the grey label bars are "theme 0 darker 35%")."""
from __future__ import annotations

import colorsys
import re

# Office default palette, used if a template has no theme part.
DEFAULT_THEME = ["FFFFFF", "000000", "E7E6E6", "44546A", "5B9BD5", "ED7D31",
                 "A5A5A5", "FFC000", "4472C4", "70AD47", "0563C1", "954F72"]

# Excel legacy indexed colours (only the ones that realistically appear).
INDEXED = {0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF",
           5: "FFFF00", 6: "FF00FF", 7: "00FFFF", 8: "000000", 9: "FFFFFF",
           10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF",
           15: "00FFFF", 22: "C0C0C0", 23: "808080", 64: "000000", 65: "FFFFFF"}


def parse_theme(theme_xml: str | None) -> list[str]:
    """Return the 12 theme colours in Excel *index* order
    (0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9 accents, 10 hlink, 11 folHlink)."""
    if not theme_xml:
        return list(DEFAULT_THEME)
    m = re.search(r"<a:clrScheme.*?</a:clrScheme>", theme_xml, re.S)
    if not m:
        return list(DEFAULT_THEME)
    block = m.group(0)
    out = {}
    for tag in ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                "accent4", "accent5", "accent6", "hlink", "folHlink"]:
        t = re.search(rf"<a:{tag}>(.*?)</a:{tag}>", block, re.S)
        val = None
        if t:
            s = re.search(r'srgbClr val="([0-9A-Fa-f]{6})"', t.group(1))
            if s:
                val = s.group(1)
            else:
                s = re.search(r'lastClr="([0-9A-Fa-f]{6})"', t.group(1))
                if s:
                    val = s.group(1)
        out[tag] = (val or "000000").upper()
    return [out["lt1"], out["dk1"], out["lt2"], out["dk2"], out["accent1"],
            out["accent2"], out["accent3"], out["accent4"], out["accent5"],
            out["accent6"], out["hlink"], out["folHlink"]]


def apply_tint(hex6: str, tint: float) -> str:
    if not tint:
        return hex6.upper()
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint < 0:
        l = l * (1 + tint)
    else:
        l = l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def resolve(color, theme: list[str], default: str | None = None) -> str | None:
    """openpyxl Color -> 'RRGGBB' (or default when the colour is unset/auto)."""
    if color is None:
        return default
    try:
        ctype = color.type
    except AttributeError:
        return default
    if ctype == "rgb":
        rgb = color.rgb
        if not isinstance(rgb, str) or len(rgb) < 6:
            return default
        return rgb[-6:].upper()
    if ctype == "theme":
        idx = color.theme
        if idx is None or idx >= len(theme):
            return default
        return apply_tint(theme[idx], color.tint or 0)
    if ctype == "indexed":
        return INDEXED.get(color.indexed, default)
    return default


def resolve_xml_color(attrs: str, theme: list[str]) -> str | None:
    """Resolve a raw XML colour element's attributes (rgb=.. / theme=.. tint=..)."""
    m = re.search(r'rgb="([0-9A-Fa-f]{6,8})"', attrs)
    if m:
        return m.group(1)[-6:].upper()
    m = re.search(r'theme="(\d+)"', attrs)
    if m:
        t = re.search(r'tint="(-?[0-9.Ee+-]+)"', attrs)
        idx = int(m.group(1))
        base = theme[idx] if idx < len(theme) else "000000"
        return apply_tint(base, float(t.group(1)) if t else 0.0)
    m = re.search(r'indexed="(\d+)"', attrs)
    if m:
        return INDEXED.get(int(m.group(1)))
    return None


def hex_to_rgb(hex6: str) -> tuple[int, int, int]:
    return tuple(int(hex6[i:i + 2], 16) for i in (0, 2, 4))
