"""Sample Tag 生成器 — D240 / D53
Target 新一季範本（更新 IMU Goal）＋ PCN ➜ Sample Tag Excel（條碼、縮圖內嵌）＋ 每張 Sample Tag PNG ＋ 列印 PDF。
Run:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import re
import os
from datetime import date

import sys
from pathlib import Path

import pandas as pd
import streamlit as st

# Streamlit Cloud re-runs app.py after a GitHub update but keeps already-imported
# core.* modules in memory, so new names (e.g. norm_subclass) fail to import until
# a reboot. Drop core.* whenever any core file changed so the new code is loaded.
_CORE = Path(__file__).with_name("core")
_sig = tuple(sorted((p.name, p.stat().st_mtime_ns, p.stat().st_size) for p in _CORE.glob("*.py")))
if getattr(sys, "_sample_tag_core_sig", None) != _sig:
    for _m in [m for m in sys.modules if m == "core" or m.startswith("core.")]:
        del sys.modules[_m]
    sys._sample_tag_core_sig = _sig

from core import images, spk
from core.config import PROGRAMS
from core.dropdowns import parse_template, update_dropdowns
from core.pcn import assign_subclass, flag_issues, init_editable, norm_subclass, read_pcn
from core.pipeline import build_xlsx, make_cards, output_stem, png_zip, safe_name
from core.render import CardRenderer, cards_pdf, to_png_bytes

st.set_page_config(page_title="Sample Tag 生成器", page_icon="🏷️", layout="wide")

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------------------------------------------------------------------------
# password
# ---------------------------------------------------------------------------
def _password() -> str | None:
    try:
        if "APP_PASSWORD" in st.secrets:
            return str(st.secrets["APP_PASSWORD"])
    except Exception:
        pass
    return os.environ.get("APP_PASSWORD")


def gate() -> bool:
    pw = _password()
    if not pw:
        return True
    if st.session_state.get("auth_ok"):
        return True
    st.title("🏷️ Sample Tag 生成器")
    with st.form("login"):
        entered = st.text_input("密碼", type="password")
        if st.form_submit_button("進入"):
            if entered == pw:
                st.session_state["auth_ok"] = True
                st.rerun()
            else:
                st.error("密碼錯誤")
    return False


# ---------------------------------------------------------------------------
# cached helpers
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def cached_template_info(data: bytes):
    return parse_template(data)


@st.cache_data(show_spinner=False)
def cached_season_template(builtin: bytes, season: bytes) -> bytes:
    """Built-in layout + the new season's Dropdowns (IMU goals)."""
    return update_dropdowns(builtin, season)


@st.cache_data(show_spinner=False)
def cached_pcn(data: bytes, name: str):
    return read_pcn(data, name)


@st.cache_data(show_spinner=False)
def cached_spk(data: bytes):
    return spk.extract(data)


@st.cache_data(show_spinner=False)
def cached_photos(files: tuple, trim: bool):
    return images.load(list(files), trim)


@st.cache_resource(show_spinner=False)
def cached_renderer(data: bytes) -> CardRenderer:
    return CardRenderer(data)


def _ekey(code: str, name: str) -> str:
    """Versioned widget key: every saved edit / filter change mounts a fresh editor,
    so the next edit is never lost to a data-changed reset."""
    return f"{code}_{name}_v{st.session_state.get(f'{code}_{name}_ver', 0)}"


def _bump(code: str, name: str) -> None:
    st.session_state[f"{code}_{name}_ver"] = st.session_state.get(f"{code}_{name}_ver", 0) + 1


def _h(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()[:12]


# ---------------------------------------------------------------------------
# item table editing (edits survive filtering)
# ---------------------------------------------------------------------------
EDIT_COLS = ["selected", "style", "photo", "description", "issues", "subclass", "first_cost", "retail", "duty",
             "sample_date", "ly_dpci", "ly_cost", "ly_retail", "ly_duty", "comments",
             "pcn_item_id", "spark_pid", "apex_id", "country", "dpci", "subclass_name", "factory"]
EDITABLE = {"selected", "subclass", "sample_date", "ly_dpci", "ly_cost", "ly_retail", "ly_duty", "comments"}


def _apply_edits(code: str):
    key = _ekey(code, "editor")
    master = st.session_state[f"{code}_df"]
    shown = st.session_state.get(f"{code}_shown_index", [])
    state = st.session_state.get(key, {})
    for pos, changes in state.get("edited_rows", {}).items():
        if int(pos) >= len(shown):
            continue
        idx = shown[int(pos)]
        for col, val in changes.items():
            if col not in EDITABLE:
                continue
            if col == "subclass":
                if master.at[idx, "subclass_pcn"]:
                    continue                                          # PCN value wins
                val = norm_subclass(val)
                master.at[idx, "program"] = val
            master.at[idx, col] = val
    st.session_state[f"{code}_df"] = master
    _bump(code, "editor")


def column_config():
    return {
        "selected": st.column_config.CheckboxColumn("生成", width="small"),
        "style": st.column_config.TextColumn("MFG Style #", disabled=True),
        "photo": st.column_config.TextColumn("圖", disabled=True, width="small"),
        "description": st.column_config.TextColumn("品名", disabled=True, width="large"),
        "subclass": st.column_config.TextColumn("Subclass", width="small",
                                                help="以 PCN 為準；PCN 空白的品項可在這裡填數字（例如 27）"),
        "first_cost": st.column_config.NumberColumn("First Cost", format="dollar", disabled=True),
        "retail": st.column_config.NumberColumn("Retail", format="dollar", disabled=True),
        "duty": st.column_config.NumberColumn("Duty", format="percent", disabled=True),
        "sample_date": st.column_config.DateColumn("Sample Date", format="YYYY/MM/DD"),
        "ly_dpci": st.column_config.TextColumn("LY DPCI"),
        "ly_cost": st.column_config.NumberColumn("LY Cost", format="dollar"),
        "ly_retail": st.column_config.NumberColumn("LY Retail", format="dollar"),
        "ly_duty": st.column_config.NumberColumn("LY Duty", format="percent"),
        "comments": st.column_config.TextColumn("Comments", width="medium"),
        "issues": st.column_config.TextColumn("待確認", disabled=True, width="medium", help="PCN 缺漏或需留意的資料"),
        "pcn_item_id": st.column_config.TextColumn("PCN Item ID", disabled=True),
        "spark_pid": st.column_config.TextColumn("Spark PID", disabled=True),
        "apex_id": st.column_config.TextColumn("Apex ID", disabled=True),
        "country": st.column_config.TextColumn("Country", disabled=True),
        "dpci": st.column_config.TextColumn("DPCI", disabled=True),
        "subclass_name": st.column_config.TextColumn("Subclass Name", disabled=True),
        "factory": st.column_config.TextColumn("Factory", disabled=True),
    }


# ---------------------------------------------------------------------------
# product pictures panel
# ---------------------------------------------------------------------------
def photo_panel(code, df, sel, photos, notes, spk_images, photo_map):
    with_pic = int(sel["photo"].eq("🖼️").sum())
    auto = [p for p in photos if p.rule in ("exact", "contains")]
    unmatched = [p for p in photos if not p.matched]
    with st.expander(f"🖼️ 產品圖片：要生成的 {len(sel)} 張 Sample Tag 中有 {with_pic} 張有圖"
                     + (f"　｜　{len([p for p in unmatched if p.rule != 'manual'])} 張圖待指定" if unmatched else ""),
                     expanded=bool(unmatched)):
        for n in notes:
            st.caption(n)
        if photos:
            st.caption(f"圖檔 {len(photos)} 張：依檔名自動對應 {len(auto)} 張，未對應 {len(unmatched)} 張。"
                       "同一品項有多張圖時取檔名排序第一張；手動指定、圖檔優先於 SPK。所有圖片都會自動裁掉白邊。")
        if auto:
            rows = []
            for p in auto:
                for i in p.matched:
                    rows.append({"圖檔": p.name, "對應 MFG Style": df.at[i, "style"],
                                 "品名": df.at[i, "description"], "方式": "檔名相同" if p.rule == "exact" else "檔名包含"})
            st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch", height=min(300, 38 + 35 * len(rows)))
        if unmatched:
            st.markdown("**檔名對不到品項的圖片 — 請手動指定：**")
            opts = [-1] + [i for i in df.index if df.at[i, "style"]]
            label = {-1: "（不使用）"}
            label.update({i: f"{df.at[i, 'style']} — {str(df.at[i, 'description'])[:45]}" for i in opts if i != -1})
            for p in unmatched[:80]:
                a, b = st.columns([1, 5])
                a.image(p.png, width=90)
                b.selectbox(p.name, opts, format_func=lambda i: label.get(i, str(i)),
                            key=f"{code}_man_{_h(p.name.encode())}")
            if len(unmatched) > 80:
                st.caption(f"另有 {len(unmatched) - 80} 張未顯示；建議把檔名改成 MFG Style 再上傳。")


# ---------------------------------------------------------------------------
# Step ① Target new-season template: only its Dropdowns (season IMU goals set by
# Target's Sourcing Manager) are used; layout / formulas always come from the built-in template.
# ---------------------------------------------------------------------------
def _sub_order(x: str):
    m = re.search(r"(\d+)$", x or "")
    return (0, int(m.group(1))) if m else (1, x)


def sub_num(x) -> str:
    """'Subclass 22' -> '22' for on-screen tables; cards / Excel keep the full text."""
    m = re.fullmatch(r"Subclass (\d+)", str(x or ""))
    return m.group(1) if m else str(x or "")


def goals_for(info, df, basis) -> dict:
    """Season IMU goal per Subclass (from the Dropdowns), used for the tag colours and the Data sheet."""
    subs = sorted({x for x in df["subclass"] if x}, key=_sub_order)
    return {x: info.goal(x, basis) for x in subs}


# ---------------------------------------------------------------------------
# main page per program
# ---------------------------------------------------------------------------
def sample_tag_tab(cfg):
    code = cfg.code
    builtin_bytes = cfg.template_path.read_bytes()
    c0, c1, c2 = st.columns(3)
    with c0:
        newtpl = st.file_uploader("① Target 新一季範本（更新 IMU Goal）", type=["xlsx", "xlsm"], key=f"{code}_newtpl",
                                  help="Target Sourcing Manager 提供的新一季 Sample Tag 範本，只取用其中的 IMU Goal（Dropdowns）；"
                                       "Sample Tag 版面、公式一律以內建範本為準。未放入時沿用內建範本的 IMU Goal。")
    with c1:
        pcn_file = st.file_uploader("② PCN（必填）", type=["xlsx"], key=f"{code}_pcn",
                                    help="Target PCN 匯出檔，例如「2027 C5 D53 WSc - Boo Basket & Wrap.xlsx」")
    with c2:
        pic_files = st.file_uploader("③ 產品圖片（選填）", key=f"{code}_pics", accept_multiple_files=True,
                                     type=["xlsx", "zip"] + [e.lstrip(".") for e in images.IMAGE_EXT],
                                     help="可放：圖檔（可多選）、ZIP（內含圖檔）、或 SPK Excel。"
                                          "檔名含 MFG Style / DPCI / PCN Item ID 就會自動對應；對不到的可在下方手動指定。")

    tpl_bytes = cached_season_template(builtin_bytes, newtpl.getvalue()) if newtpl is not None else builtin_bytes
    info = cached_template_info(tpl_bytes)
    # IMU colours follow the "IMU GOALS" box on the season template's Instructions sheet
    # (goal column + yellow band the SM shows), not the template's own CF wiring.
    box = cached_template_info(newtpl.getvalue()) if newtpl is not None else info

    if pcn_file is None:
        for n in cfg.notes:
            st.info(n)
        return

    pcn_bytes = pcn_file.getvalue()
    df_raw, warns = cached_pcn(pcn_bytes, pcn_file.name)
    if df_raw.empty:
        st.error("PCN 讀不到品項。")
        return
    basis = box.box_basis or info.box_basis or "imu"
    band = box.box_band or info.box_band or cfg.default_yellow_band   # IMU light on the tag (not shown in the App)

    # ---- session table -------------------------------------------------------
    sig = (_h(pcn_bytes), pcn_file.name)
    if st.session_state.get(f"{code}_sig") != sig:
        df = init_editable(assign_subclass(df_raw))
        st.session_state[f"{code}_df"] = df
        st.session_state[f"{code}_sig"] = sig
        st.session_state.pop(f"{code}_out", None)
        _bump(code, "editor")
    df = st.session_state[f"{code}_df"]

    goals = goals_for(info, df, basis)
    df = flag_issues(df, goals)
    st.session_state[f"{code}_df"] = df
    no_sub = int((df["style"].ne("") & df["subclass"].eq("")).sum())
    if no_sub:
        st.warning(f"有 {no_sub} 個品項 PCN 沒有 Subclass，請在表格的 Subclass 欄手動填入（數字即可），"
                   "系統將自動抓取該季 IMU。可勾「只看待確認」快速找出。")

    for w in warns:
        st.warning(w)
    # ---- product pictures: SPK / ZIP / loose images ---------------------------
    spk_images, photos, photo_notes = None, [], []
    trim = True                      # always trim white / transparent margins (all sources)
    if pic_files:
        xl = [f for f in pic_files if f.name.lower().endswith(".xlsx")]
        other = tuple((f.name, f.getvalue()) for f in pic_files if not f.name.lower().endswith(".xlsx"))
        for f in xl:
            got, notes = cached_spk(f.getvalue())
            photo_notes += notes
            if spk_images is None:
                spk_images = got
            else:
                for k, v in got.items():
                    for kk, vv in v.items():
                        spk_images.setdefault(k, {}).setdefault(kk, vv)
        if other:
            photos, notes = cached_photos(other, trim)
            photo_notes += notes
    photo_map = images.match(photos, df) if photos else {}
    manual = {}
    for p in photos:
        if not p.matched:
            choice = st.session_state.get(f"{code}_man_{_h(p.name.encode())}")
            if choice is not None and choice != -1 and choice in df.index:
                manual[choice] = p.png
                p.rule = "manual"
    photo_map.update(manual)
    has_spk = (lambda r: bool(spk_images) and spk.lookup(spk_images, r) is not None)
    df["photo"] = ["🖼️" if (i in photo_map or has_spk(r)) else "" for i, r in df.iterrows()]
    st.session_state[f"{code}_df"] = df

    # ---- filters -----------------------------------------------------------------
    f1, f2, f3, f4 = st.columns([2, 2, 1, 1])
    with f1:
        subs = sorted((p for p in df["subclass"].unique() if p), key=_sub_order)
        pf = st.multiselect("篩選 Subclass", subs, format_func=sub_num, key=f"{code}_pf")
    with f2:
        q = st.text_input("搜尋（Style / 品名 / DPCI）", key=f"{code}_q")
    with f3:
        only_issue = st.checkbox("只看待確認", key=f"{code}_oi")
    mask = pd.Series(True, index=df.index)
    if pf:
        mask &= df["subclass"].isin(pf)
    if q:
        ql = q.lower()
        mask &= (df["style"].str.lower().str.contains(ql, regex=False) |
                 df["description"].str.lower().str.contains(ql, regex=False) |
                 df["dpci"].str.lower().str.contains(ql, regex=False))
    if only_issue:
        mask &= df["issues"].ne("")
    with f4:
        b1, b2 = st.columns(2)
        if b1.button("全選", key=f"{code}_all", help="勾選目前篩選結果中有 MFG Style 的列"):
            df.loc[mask & df["style"].ne(""), "selected"] = True
            _bump(code, "editor")
        if b2.button("全不選", key=f"{code}_none"):
            df.loc[mask, "selected"] = False
            _bump(code, "editor")

    shown = df.loc[mask, EDIT_COLS]
    new_index = list(shown.index)
    if st.session_state.get(f"{code}_shown_index") != new_index:
        # rows on screen changed (filter / new PCN): pending editor deltas are positional,
        # so drop them — every edit is already saved in the master table
        _bump(code, "editor")
    st.session_state[f"{code}_shown_index"] = new_index
    shown = shown.assign(subclass=shown["subclass"].map(sub_num))
    st.data_editor(shown, key=_ekey(code, "editor"), on_change=_apply_edits, args=(code,),
                   column_config=column_config(), hide_index=True, width="stretch",
                   height=min(520, 38 + 35 * max(3, len(shown))), num_rows="fixed")

    df = st.session_state[f"{code}_df"]
    # generate exactly what the filter shows: checked rows within the current filter only
    sel = df[df["selected"] & df["style"].ne("") & mask.reindex(df.index, fill_value=False)]
    n_issue = int((sel["issues"] != "").sum())
    m1, m2, m3 = st.columns(3)
    m1.metric("要生成的 Sample Tag", len(sel), help="目前篩選結果中有勾選「生成」的品項")
    m2.metric("PCN 品項", len(df))
    m3.metric("待確認的 Sample Tag", n_issue)

    if pic_files:
        photo_panel(code, df, sel, photos, photo_notes, spk_images, photo_map)

    if sel.empty:
        st.info("目前篩選結果中沒有勾選任何品項，請勾選至少一張 Sample Tag。")
        return

    renderer = cached_renderer(tpl_bytes)
    vendor = info.vendor_default or "Lucky Star"
    vendor_fill = info.vendor_color(vendor)

    with st.expander("預覽前 4 張", expanded=True):
        cards_prev = make_cards(sel.head(4), info, basis, spk_images, photo_map, goals)
        cols = st.columns(2)
        for i, c in enumerate(cards_prev):
            img = renderer.render(c, vendor=vendor, vendor_fill=vendor_fill, goal=c["goal"], band=band)
            cols[i % 2].image(img, width="stretch")

    if st.button(f"生成 {len(sel)} 張 Sample Tag", type="primary", key=f"{code}_go"):
        prog = st.progress(0.0, text="準備資料…")
        cards = make_cards(sel, info, basis, spk_images, photo_map, goals)
        prog.progress(0.05, text="生成 Excel…")
        xlsx = build_xlsx(tpl_bytes, cards, info, band)
        pngs, imgs = [], []
        for i, c in enumerate(cards):
            img = renderer.render(c, vendor=vendor, vendor_fill=vendor_fill, goal=c["goal"], band=band)
            imgs.append(img)
            pngs.append((f"{i + 1:03d}_{safe_name(c['style'])}.png", to_png_bytes(img)))
            prog.progress(0.1 + 0.75 * (i + 1) / len(cards), text=f"繪製 Sample Tag {i + 1}/{len(cards)}…")
        prog.progress(0.9, text="組成 PDF…")
        pdf = cards_pdf(imgs)
        prog.progress(1.0, text="完成")
        stem = output_stem(pcn_file.name, code)
        st.session_state[f"{code}_out"] = {
            "stem": stem, "xlsx": xlsx, "zip": png_zip(pngs), "pdf": pdf, "n": len(cards),
            "thumbs": sum(1 for c in cards if c.get("thumb_png")),
        }

    out = st.session_state.get(f"{code}_out")
    if out:
        st.success(f"完成：{out['n']} 張 Sample Tag" + (f"（含縮圖 {out['thumbs']} 張）" if out["thumbs"] else ""))
        d1, d2, d3 = st.columns(3)
        d1.download_button("⬇️ Excel（Sample Tag）", out["xlsx"], f"{out['stem']}.xlsx", XLSX_MIME,
                           key=f"{code}_dlx", width="stretch")
        d2.download_button("⬇️ 每張 Tag 獨立圖片檔（PNG）", out["zip"], f"{out['stem']} PNG.zip", "application/zip",
                           key=f"{code}_dlz", width="stretch")
        d3.download_button("⬇️ 整份列印檔（PDF）", out["pdf"], f"{out['stem']}.pdf", "application/pdf",
                           key=f"{code}_dlp", width="stretch")


# ---------------------------------------------------------------------------
def main():
    if not gate():
        return
    st.title("🏷️ Sample Tag 生成器")
    st.caption("上傳 PCN ➜ 勾選品項／篩選 Subclass ➜ 一鍵生成 Sample Tag Excel、每張 Tag 的獨立圖片檔（PNG）、整份列印檔（PDF）")
    tabs = st.tabs([p.title for p in PROGRAMS.values()])
    for tab, cfg in zip(tabs, PROGRAMS.values()):
        with tab:
            sample_tag_tab(cfg)


main()
