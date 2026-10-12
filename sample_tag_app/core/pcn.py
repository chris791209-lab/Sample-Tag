"""Read a Target PCN export and pull out exactly the fields a sample tag needs.

Columns are found by *header name*, never by position: the D240 and D53 PCN
exports carry the same headers in different places (MFG Style is col O in D240,
col K in D53) and some headers carry a trailing " *"."""
from __future__ import annotations

import io
import math
import re
from datetime import date

import pandas as pd

# card field -> accepted PCN headers (normalised: lower-case, no "*", single spaces)
FIELDS: dict[str, list[str]] = {
    "style": ["manufacturer style #", "mfg style #", "vendor style #"],
    "vendor_desc": ["vendor product description"],
    "product_desc": ["product description"],
    "pcn_item_id": ["pcn item id"],
    "spark_pid": ["spark pid"],
    "apex_id": ["apex id"],
    "fca": ["fca factory city unit cost"],
    "fob": ["fob unit cost"],
    "retail": ["suggested unit retail"],
    "country": ["import country of origin"],
    "duty": ["duty rate %"],
    "dpci": ["dpci"],
    "subclass_id": ["subclass id"],
    "subclass_name": ["subclass name"],
    "factory": ["factory name"],
    "vendor": ["import vendor name"],
    "status": ["item system status"],
}

# Columns shown/edited in the app (Traditional Chinese labels live in the UI).
ITEM_COLUMNS = [
    "selected", "style", "description", "program", "pcn_item_id", "spark_pid",
    "apex_id", "first_cost", "cost_basis", "retail", "country", "duty",
    "sample_date", "ly_dpci", "ly_cost", "ly_retail", "ly_duty", "comments",
    "dpci", "subclass_name", "factory", "issues",
]


def norm_header(h) -> str:
    s = str(h or "").replace("\n", " ").strip()
    s = re.sub(r"\s*\*+\s*$", "", s)          # trailing " *"
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def _clean_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        if math.isnan(v):
            return ""
        if v.is_integer():
            return str(int(v))
    s = str(v).strip()
    if s.lower() in ("nan", "none", "nat"):
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def _num(v) -> float | None:
    s = _clean_text(v)
    if not s:
        return None
    s = s.replace("$", "").replace(",", "").strip()
    pct = s.endswith("%")
    s = s.rstrip("%").strip()
    try:
        x = float(s)
    except ValueError:
        return None
    return x / 100 if pct else x


def _duty(v) -> float | None:
    x = _num(v)
    if x is None:
        return None
    # PCN stores duty as a fraction (0.2). Guard against "20" meaning 20 %.
    return x / 100 if x > 1 else x


def program_from_filename(filename: str) -> str:
    """'2027 C5 D53 WSc - Boo Basket & Wrap.xlsx' -> 'Boo Basket & Wrap'."""
    stem = re.sub(r"\.(xlsx|xlsm|xls)$", "", filename or "", flags=re.I)
    stem = re.sub(r"\s*\(\d+\)\s*$", "", stem)       # "(1)" download suffix
    stem = re.sub(r"[_]+", " ", stem)
    if " - " in stem:
        return stem.split(" - ")[-1].strip()
    return ""


def _key(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def match_program(name: str, programs: list[str]) -> str | None:
    """Match a program name to the Dropdowns list (ignores case/spaces/&)."""
    k = _key(name)
    if not k:
        return None
    for p in programs:
        if _key(p) == k:
            return p
    for p in programs:
        pk = _key(p)
        if pk and (pk in k or k in pk):
            return p
    return None


def read_pcn(data: bytes, filename: str = "") -> tuple[pd.DataFrame, list[str]]:
    """Return one row per PCN item with the card fields, plus warnings."""
    warnings: list[str] = []
    xls = pd.ExcelFile(io.BytesIO(data))
    sheet = "Data" if "Data" in xls.sheet_names else xls.sheet_names[0]
    raw = pd.read_excel(xls, sheet_name=sheet, dtype=object)
    headers = {norm_header(c): c for c in raw.columns}

    def col(field):
        for h in FIELDS[field]:
            if h in headers:
                return raw[headers[h]]
        return pd.Series([None] * len(raw), dtype=object)

    missing = [f for f in ("style", "pcn_item_id", "retail") if not any(h in headers for h in FIELDS[f])]
    if missing:
        warnings.append("PCN 缺少欄位：" + ", ".join(FIELDS[m][0] for m in missing))

    rows = []
    for i in range(len(raw)):
        style = _clean_text(col("style").iloc[i])
        vdesc = _clean_text(col("vendor_desc").iloc[i])
        pdesc = _clean_text(col("product_desc").iloc[i])
        if not (style or vdesc or pdesc):
            continue                                   # truly empty row
        fca, fob = _num(col("fca").iloc[i]), _num(col("fob").iloc[i])
        cost, basis = (fca, "FCA") if fca is not None else ((fob, "FOB") if fob is not None else (None, ""))
        rows.append({
            "pcn_row": i + 2,
            "style": style,
            "description": vdesc,                      # fixed: Vendor Product Description only
            "pcn_item_id": _clean_text(col("pcn_item_id").iloc[i]),
            "spark_pid": _clean_text(col("spark_pid").iloc[i]),
            "apex_id": _clean_text(col("apex_id").iloc[i]),
            "first_cost": cost,
            "cost_basis": basis,
            "retail": _num(col("retail").iloc[i]),
            "country": _clean_text(col("country").iloc[i]),
            "duty": _duty(col("duty").iloc[i]),
            "dpci": _clean_text(col("dpci").iloc[i]),
            "subclass_id": _clean_text(col("subclass_id").iloc[i]),
            "subclass_name": _clean_text(col("subclass_name").iloc[i]),
            "factory": _clean_text(col("factory").iloc[i]),
            "vendor": _clean_text(col("vendor").iloc[i]),
            "status": _clean_text(col("status").iloc[i]),
        })
    df = pd.DataFrame(rows)
    for c in ("first_cost", "retail", "duty"):
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce").astype("float64")
    if df.empty:
        warnings.append("PCN 沒有可用的品項列。")
        return df, warnings

    dup = df["style"].ne("") & df["style"].duplicated(keep=False)
    if dup.any():
        warnings.append(f"MFG Style 重複 {int(df.loc[dup, 'style'].nunique())} 組（已標示，仍可各自產生卡片）。")
    nostyle = df["style"].eq("")
    if nostyle.any():
        warnings.append(f"{int(nostyle.sum())} 列沒有 MFG Style（無法產生條碼，預設不勾選）。")
    return df, warnings


def norm_subclass(v) -> str:
    """'27' / '27.0' / 'subclass 27' / 'Subclass27' -> 'Subclass 27'; '' stays ''."""
    t = _clean_text(v)
    if not t:
        return ""
    m = re.search(r"(\d+)", t)
    return f"Subclass {int(m.group(1))}" if m else t.strip()


def assign_subclass(df: pd.DataFrame) -> pd.DataFrame:
    """Card row 3 = the item's Subclass exactly as the PCN has it ("Subclass <ID>").
    The IMU target is looked up by that same Subclass (D240 and D53 alike)."""
    df = df.copy()
    df["subclass_pcn"] = df["subclass_id"].map(norm_subclass)
    df["subclass"] = df["subclass_pcn"]
    df["program"] = df["subclass"]
    return df


def flag_issues(df: pd.DataFrame, goals: dict) -> pd.DataFrame:
    """goals = {"Subclass 27": 0.714, ...} (IMU+Duty Goal actually used)."""
    """Per-row data-quality notes shown in the app (and only there)."""
    df = df.copy()
    dup = df["style"].ne("") & df["style"].duplicated(keep=False)
    notes = []
    for i, r in df.iterrows():
        n = []
        if not r["style"]:
            n.append("缺 MFG Style")
        if not r.get("description"):
            n.append("缺 Vendor Product Description")
        if dup.loc[i]:
            n.append("Style 重複")
        if r.get("first_cost") is None or pd.isna(r.get("first_cost")):
            n.append("缺成本")
        elif r.get("cost_basis") == "FOB":
            n.append("成本取 FOB")
        if r.get("retail") is None or pd.isna(r.get("retail")):
            n.append("缺零售價")
        if r.get("duty") is None or pd.isna(r.get("duty")):
            n.append("缺 Duty")
        if not r.get("subclass"):
            n.append("缺 Subclass（可在 Subclass 欄填入）")
        notes.append("、".join(n))
    df["issues"] = notes
    return df


def init_editable(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["selected"] = df["style"].ne("")
    df["sample_date"] = date.today()
    for c in ("ly_dpci", "comments"):
        df[c] = ""
    for c in ("ly_cost", "ly_retail", "ly_duty"):
        df[c] = float("nan")
    return df


def imu(cost, retail, duty=None, with_duty=False):
    try:
        if cost is None or retail in (None, 0) or pd.isna(cost) or pd.isna(retail):
            return None
        landed = cost + (cost * duty if (with_duty and duty is not None and not pd.isna(duty)) else 0)
        if with_duty and (duty is None or pd.isna(duty)):
            landed = cost          # Excel formula treats a blank duty as 0
        return (retail - landed) / retail
    except (TypeError, ZeroDivisionError):
        return None
