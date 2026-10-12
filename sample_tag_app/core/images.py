"""Product photos from loose image files or a ZIP (for early-stage samples that
have no SPK yet). Each picture is matched to items by its file name:

  1. the file name *is* a key        HWD704A.jpg, 240-43-5528.png, 8351342.jpg
  2. the file name *contains* a key   001_HWD704A_front.jpg, IMG HWD704A (2).png
     (whole-word match; the longest key wins, so HWD704A beats HWD704)

Keys tried: MFG Style #, DPCI, PCN Item ID, Spark PID, Apex ID. Pictures that
match nothing are listed in the app so they can be assigned by hand."""
from __future__ import annotations

import io
import posixpath
import re
import zipfile
from dataclasses import dataclass, field

from PIL import Image, ImageChops, ImageOps

IMAGE_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff")
KEY_FIELDS = ("style", "dpci", "pcn_item_id", "spark_pid", "apex_id")
MAX_SIDE = 700          # px; plenty for the thumbnail box, keeps the Excel small


@dataclass
class Photo:
    name: str                       # original file name (inside the ZIP: its path)
    png: bytes
    size: tuple[int, int]
    matched: list = field(default_factory=list)   # item indexes
    rule: str = ""                  # "exact" / "contains" / "manual"


def _alnum(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def prepare(blob: bytes, trim: bool = True) -> tuple[bytes, tuple[int, int]] | None:
    """Fix phone-camera rotation, trim plain white/transparent margins, shrink, -> PNG."""
    try:
        im = Image.open(io.BytesIO(blob))
        im.load()
    except Exception:
        return None
    im = ImageOps.exif_transpose(im)
    if im.mode in ("P", "LA", "L", "CMYK", "I", "F", "1"):
        im = im.convert("RGBA")
    if trim:
        im = _trim(im)
    im.thumbnail((MAX_SIDE, MAX_SIDE), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="PNG", optimize=True)
    return buf.getvalue(), im.size


def _trim(im: Image.Image, tol: int = 12) -> Image.Image:
    rgb = im.convert("RGB")
    if im.mode == "RGBA":
        # treat transparent pixels as white
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[3])
        rgb = bg
    diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, (255, 255, 255)))
    diff = diff.convert("L").point(lambda p: 255 if p > tol else 0)
    box = diff.getbbox()
    if not box:
        return im
    w, h = im.size
    pad = max(2, int(0.03 * max(box[2] - box[0], box[3] - box[1])))
    box = (max(0, box[0] - pad), max(0, box[1] - pad), min(w, box[2] + pad), min(h, box[3] + pad))
    return im.crop(box)


def load(files: list[tuple[str, bytes]], trim: bool = True) -> tuple[list[Photo], list[str]]:
    """files = [(file name, bytes)] — single images and/or ZIPs."""
    photos, notes, skipped = [], [], []
    for name, data in files:
        low = name.lower()
        if low.endswith(".zip"):
            try:
                z = zipfile.ZipFile(io.BytesIO(data))
            except zipfile.BadZipFile:
                notes.append(f"{name}：不是有效的 ZIP。")
                continue
            for info in sorted(z.infolist(), key=lambda i: i.filename):
                fn = info.filename
                base = posixpath.basename(fn)
                if info.is_dir() or fn.startswith("__MACOSX/") or base.startswith((".", "._")):
                    continue
                if not base.lower().endswith(IMAGE_EXT):
                    continue
                res = prepare(z.read(info), trim)
                if res:
                    photos.append(Photo(fn, res[0], res[1]))
                else:
                    skipped.append(fn)
        elif low.endswith(IMAGE_EXT):
            res = prepare(data, trim)
            if res:
                photos.append(Photo(name, res[0], res[1]))
            else:
                skipped.append(name)
    if skipped:
        notes.append(f"無法讀取的圖檔 {len(skipped)} 個：" + "、".join(skipped[:8]) + ("…" if len(skipped) > 8 else ""))
    return photos, notes


def match(photos: list[Photo], items) -> dict[int, bytes]:
    """items: DataFrame (index = item id). Returns {item index: png}. Mutates
    photos[*].matched / .rule for the app's report."""
    keys: dict[str, set] = {}          # alnum key -> item indexes
    raw: dict[str, str] = {}           # alnum key -> display form (for word-boundary search)
    for idx, r in items.iterrows():
        for f in KEY_FIELDS:
            v = str(r.get(f) or "").strip()
            k = _alnum(v)
            if len(k) >= 4:
                keys.setdefault(k, set()).add(idx)
                raw.setdefault(k, v.upper())
    out: dict[int, bytes] = {}
    for p in photos:
        stem = re.sub(r"\.[A-Za-z0-9]+$", "", posixpath.basename(p.name)).upper()
        sk = _alnum(stem)
        hit = None
        if sk in keys:
            hit, p.rule = keys[sk], "exact"
        else:
            best = None
            for k, disp in raw.items():
                if len(k) < 5:
                    continue
                if re.search(rf"(?<![A-Z0-9]){re.escape(disp)}(?![A-Z0-9])", stem):
                    if best is None or len(k) > len(best):
                        best = k
            if best:
                hit, p.rule = keys[best], "contains"
        if hit:
            p.matched = sorted(hit)
            for idx in hit:
                out.setdefault(idx, p.png)      # first picture (by file name) wins
    return out
