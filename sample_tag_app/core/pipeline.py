"""Glue: edited item table -> card dicts -> workbook / PNG / PDF outputs."""
from __future__ import annotations

import io
import math
import re
import zipfile
from datetime import date, datetime

import pandas as pd

from . import barcode_img, spk
from .builder import build_workbook, imu_value
from .dropdowns import TemplateInfo

CARD_KEYS = ["style", "description", "program", "pcn_item_id", "spark_pid", "apex_id", "first_cost",
             "cost_basis", "retail", "country", "duty", "sample_date", "ly_dpci", "ly_cost", "ly_retail",
             "ly_duty", "comments", "dpci", "subclass_name", "factory"]


def _clean(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, pd.Timestamp):
        return v.date()
    if isinstance(v, datetime):
        return v.date()
    if hasattr(v, "item") and not isinstance(v, (str, bytes, date)):
        try:
            return v.item()
        except Exception:
            return v
    return v


def make_cards(df: pd.DataFrame, info: TemplateInfo, imu_basis: str = "imu_duty",
               spk_images: dict | None = None, photo_map: dict | None = None) -> list[dict]:
    """photo_map = {item index: png} from uploaded pictures (hand-assigned or
    matched by file name); it wins over the SPK picture of the same item."""
    cards = []
    for idx, r in df.iterrows():
        c = {k: _clean(r.get(k)) for k in CARD_KEYS}
        for k in ("ly_cost", "ly_retail", "ly_duty"):
            if isinstance(c.get(k), str):
                s = c[k].replace("$", "").replace(",", "").strip().rstrip("%")
                try:
                    v = float(s)
                    c[k] = v / 100 if (k == "ly_duty" and v > 1) else v
                except ValueError:
                    c[k] = None
        c["goal"] = info.goal(c.get("program") or "", imu_basis)
        c["barcode_png"] = barcode_img.png_bytes(str(c["style"])) if c.get("style") else None
        c["thumb_png"] = ((photo_map or {}).get(idx)
                          or (spk.lookup(spk_images, c) if spk_images else None))
        cards.append(c)
    return cards


def imu_status(card: dict, band: float) -> str:
    v = imu_value(card.get("first_cost"), card.get("retail"), card.get("duty"), False)
    g = card.get("goal")
    if v is None or g is None:
        return ""
    if v >= g:
        return "🟢"
    if v >= g - band:
        return "🟡"
    return "🔴"


def safe_name(s: str) -> str:
    s = re.sub(r'[\\/:*?"<>|]+', "_", str(s or "").strip())
    return s[:80] or "card"


def output_stem(filename: str, code: str) -> str:
    stem = re.sub(r"\.(xlsx|xlsm|xls)$", "", filename or "", flags=re.I).strip()
    stem = re.sub(r"\s*\(\d+\)$", "", stem)
    return safe_name(f"{stem} Sample Tag" if stem else f"{code} Sample Tag")


def build_xlsx(template_bytes: bytes, cards: list[dict], info: TemplateInfo, band: float) -> bytes:
    vendor = info.vendor_default or "Lucky Star"
    return build_workbook(template_bytes, cards, vendor=vendor, yellow_band=band,
                          vendor_fill=info.vendor_color(vendor))


def png_zip(pngs: list[tuple[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        used = set()
        for name, data in pngs:
            n, k = name, 2
            while n in used:
                n = re.sub(r"(\.png)$", f"_{k}\\1", name)
                k += 1
            used.add(n)
            z.writestr(n, data)
    return buf.getvalue()
