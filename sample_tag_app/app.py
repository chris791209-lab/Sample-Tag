"""Sample Tag 產生器 — D240 / D53
PCN ➜ Sample Tag Excel（條碼、縮圖內嵌）＋ 每張卡 PNG ＋ 列印 PDF；另含 Dropdown 更新。
Run:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import os
from datetime import date

import pandas as pd
import streamlit as st

from core import images, spk
from core.config import PROGRAMS
from core.dropdowns import diff_templates, parse_template, update_dropdowns
from core.pcn import assign_programs, flag_issues, init_editable, program_from_filename, read_pcn
from core.pipeline import build_xlsx, imu_status, make_cards, output_stem, png_zip, safe_name
from core.render import CardRenderer, cards_pdf, to_png_bytes

st.set_page_config(page_title="Sample Tag 產生器", page_icon="🏷️", layout="wide")

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLSM_MIME = "application/vnd.ms-excel.sheet.macroEnabled.12"


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
    st.title("🏷️ Sample Tag 產生器")
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


def _h(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()[:12]


# ---------------------------------------------------------------------------
# item table editing (edits survive filtering)
# ---------------------------------------------------------------------------
EDIT_COLS = ["selected", "style", "photo", "description", "issues", "program", "first_cost", "retail", "duty", "imu_flag",
             "sample_date", "ly_dpci", "ly_cost", "ly_retail", "ly_duty", "comments",
             "pcn_item_id", "spark_pid", "apex_id", "country", "dpci", "subclass_name", "factory"]
EDITABLE = {"selected", "program", "sample_date", "ly_dpci", "ly_cost", "ly_retail", "ly_duty", "comments"}


def _apply_edits(code: str):
    key = f"{code}_editor"
    master = st.session_state[f"{code}_df"]
    shown = st.session_state.get(f"{code}_shown_index", [])
    state = st.session_state.get(key, {})
    for pos, changes in state.get("edited_rows", {}).items():
        if int(pos) >= len(shown):
            continue
        idx = shown[int(pos)]
        for col, val in changes.items():
            if col in EDITABLE:
                master.at[idx, col] = val
    st.session_state[f"{code}_df"] = master


def column_config(programs: list[str]):
    return {
        "selected": st.column_config.CheckboxColumn("產生", width="small"),
        "style": st.column_config.TextColumn("MFG Style #", disabled=True),
        "photo": st.column_config.TextColumn("圖", disabled=True, width="small"),
        "description": st.column_config.TextColumn("品名", disabled=True, width="large"),
        "program": st.column_config.SelectboxColumn("Program", options=programs or None, width="medium"),
        "first_cost": st.column_config.NumberColumn("First Cost", format="dollar", disabled=True),
        "retail": st.column_config.NumberColumn("Retail", format="dollar", disabled=True),
        "duty": st.column_config.NumberColumn("Duty", format="percent", disabled=True),
        "imu_flag": st.column_config.TextColumn("IMU", disabled=True, width="small"),
        "sample_date": st.column_config.DateColumn("Sample Date", format="YYYY/MM/DD"),
        "ly_dpci": st.column_config.TextColumn("LY DPCI"),
        "ly_cost": st.column_config.NumberColumn("LY Cost", format="dollar"),
        "ly_retail": st.column_config.NumberColumn("LY Retail", format="dollar"),
        "ly_duty": st.column_config.NumberColumn("LY Duty", format="percent"),
        "comments": st.column_config.TextColumn("Comments", width="medium"),
        "issues": st.column_config.TextColumn("提醒", disabled=True, width="medium"),
        "pcn_item_id": st.column_config.TextColumn("PCN Item ID", disabled=True),
        "spark_pid": st.column_config.TextColumn("Spark PID", disabled=True),
        "apex_id": st.column_config.TextColumn("Apex ID", disabled=True),
        "country": st.column_config.TextColumn("Country", disabled=True),
        "dpci": st.column_config.TextColumn("DPCI", disabled=True),
        "subclass_name": st.column_config.TextColumn("Subclass", disabled=True),
        "factory": st.column_config.TextColumn("Factory", disabled=True),
    }


# ---------------------------------------------------------------------------
# product pictures panel
# ---------------------------------------------------------------------------
def photo_panel(code, df, photos, notes, spk_images, photo_map):
    sel = df[df["selected"] & df["style"].ne("")]
    with_pic = int(sel["photo"].eq("🖼️").sum())
    auto = [p for p in photos if p.rule in ("exact", "contains")]
    unmatched = [p for p in photos if not p.matched]
    with st.expander(f"🖼️ 產品圖片：已勾選的 {len(sel)} 張卡中有 {with_pic} 張有圖"
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
# Sample Tag tab
# ---------------------------------------------------------------------------
def sample_tag_tab(cfg):
    code = cfg.code
    c1, c2, c3 = st.columns(3)
    with c1:
        pcn_file = st.file_uploader("① PCN（必填）", type=["xlsx"], key=f"{code}_pcn",
                                    help="Target PCN 匯出檔，例如「2027 C5 D53 WSc - Boo Basket & Wrap.xlsx」")
    with c2:
        pic_files = st.file_uploader("② 產品圖片（選填）", key=f"{code}_pics", accept_multiple_files=True,
                                     type=["xlsx", "zip"] + [e.lstrip(".") for e in images.IMAGE_EXT],
                                     help="可放：圖檔（可多選）、ZIP（內含圖檔）、或 SPK Excel。"
                                          "檔名含 MFG Style / DPCI / PCN Item ID 就會自動對應；對不到的可在下方手動指定。")
    with c3:
        tpl_file = st.file_uploader("③ Target 範本（選填）", type=["xlsx", "xlsm"], key=f"{code}_tpl",
                                    help=f"不上傳就用內建範本：{cfg.template_file}")

    override = st.session_state.get(f"{code}_template_override")
    if tpl_file is not None:
        tpl_bytes, tpl_name = tpl_file.getvalue(), tpl_file.name
    elif override:
        tpl_bytes, tpl_name = override
    else:
        tpl_bytes, tpl_name = cfg.template_path.read_bytes(), cfg.template_file
    info = cached_template_info(tpl_bytes)
    programs = list(info.programs)
    st.caption(f"範本：{tpl_name}　｜　Program {len(programs)} 個　｜　IMU 黃燈區間 "
               f"{(info.yellow_band or cfg.default_yellow_band):.0%}　｜　廠商：{info.vendor_default.strip() or '—'}")

    if pcn_file is None:
        for n in cfg.notes:
            st.info(n)
        return

    pcn_bytes = pcn_file.getvalue()
    df_raw, warns = cached_pcn(pcn_bytes, pcn_file.name)
    if df_raw.empty:
        st.error("PCN 讀不到品項。")
        return

    # ---- program rule ------------------------------------------------------
    s1, s2 = st.columns([2, 1])
    with s1:
        if cfg.program_source == "filename":
            guess = program_from_filename(pcn_file.name)
            from core.pcn import match_program
            matched = match_program(guess, programs)
            opts = programs if programs else [guess]
            idx = opts.index(matched) if matched in opts else 0
            prog_choice = st.selectbox("Program Name（依 PCN 檔名帶出，可改）", opts, index=idx, key=f"{code}_prog")
            if not matched:
                st.warning(f"檔名中的「{guess or '（無）'}」對不到 Dropdowns 的 Program，請確認上方選擇。")
            override_prog = prog_choice
        else:
            choice = st.selectbox("Program Name", ["依 Subclass ID 自動"] + programs, key=f"{code}_prog")
            override_prog = None if choice.startswith("依") else choice
    with s2:
        basis_label = st.radio("IMU 燈號目標", [info.imu_header, info.imu_duty_header], index=1, horizontal=True,
                               key=f"{code}_basis",
                               help="IMU 紅黃綠燈對照 Dropdowns 的哪一欄目標。預設 IMU+Duty Goal。")
        basis = "imu" if basis_label == info.imu_header else "imu_duty"
    band = info.yellow_band or cfg.default_yellow_band

    # ---- session table -------------------------------------------------------
    sig = (_h(pcn_bytes), _h(tpl_bytes), override_prog or "")
    if st.session_state.get(f"{code}_sig") != sig:
        df = assign_programs(df_raw, cfg, programs, pcn_file.name, override_prog)
        df = flag_issues(df, info.programs)
        df = init_editable(df)
        st.session_state[f"{code}_df"] = df
        st.session_state[f"{code}_sig"] = sig
        st.session_state.pop(f"{code}_out", None)
    df = st.session_state[f"{code}_df"]
    df = flag_issues(df, info.programs)
    df["imu_flag"] = [imu_status({"first_cost": r.first_cost, "retail": r.retail, "duty": r.duty,
                                  "goal": info.goal(r.program, basis)}, band) for r in df.itertuples()]
    st.session_state[f"{code}_df"] = df

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
        progs_present = sorted(p for p in df["program"].unique() if p)
        pf = st.multiselect("篩選 Program", progs_present, key=f"{code}_pf")
    with f2:
        q = st.text_input("搜尋（Style / 品名 / DPCI）", key=f"{code}_q")
    with f3:
        only_issue = st.checkbox("只看有提醒的", key=f"{code}_oi")
    mask = pd.Series(True, index=df.index)
    if pf:
        mask &= df["program"].isin(pf)
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
            st.session_state.pop(f"{code}_editor", None)
        if b2.button("全不選", key=f"{code}_none"):
            df.loc[mask, "selected"] = False
            st.session_state.pop(f"{code}_editor", None)

    shown = df.loc[mask, EDIT_COLS]
    new_index = list(shown.index)
    if st.session_state.get(f"{code}_shown_index") != new_index:
        # rows on screen changed (filter / new PCN): pending editor deltas are positional,
        # so drop them — every edit is already saved in the master table
        st.session_state.pop(f"{code}_editor", None)
    st.session_state[f"{code}_shown_index"] = new_index
    st.data_editor(shown, key=f"{code}_editor", on_change=_apply_edits, args=(code,),
                   column_config=column_config(programs), hide_index=True, width="stretch",
                   height=min(520, 38 + 35 * max(3, len(shown))), num_rows="fixed")

    df = st.session_state[f"{code}_df"]
    sel = df[df["selected"] & df["style"].ne("")]
    n_issue = int((sel["issues"] != "").sum())
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("要產生的卡片", len(sel))
    m2.metric("PCN 品項", len(df))
    m3.metric("有提醒的卡片", n_issue)
    m4.metric("Excel 區塊（每塊 2 張）", (len(sel) + 1) // 2)

    if pic_files:
        photo_panel(code, df, photos, photo_notes, spk_images, photo_map)

    if sel.empty:
        st.info("請勾選至少一張卡片。")
        return

    renderer = cached_renderer(tpl_bytes)
    vendor = info.vendor_default or "Lucky Star"
    vendor_fill = info.vendor_color(vendor)

    with st.expander("預覽前 4 張卡片", expanded=True):
        cards_prev = make_cards(sel.head(4), info, basis, spk_images, photo_map)
        cols = st.columns(2)
        for i, c in enumerate(cards_prev):
            img = renderer.render(c, vendor=vendor, vendor_fill=vendor_fill, goal=c["goal"], band=band)
            cols[i % 2].image(img, width="stretch")

    if st.button(f"產生 {len(sel)} 張 Sample Tag", type="primary", key=f"{code}_go"):
        prog = st.progress(0.0, text="準備資料…")
        cards = make_cards(sel, info, basis, spk_images, photo_map)
        prog.progress(0.05, text="產生 Excel…")
        xlsx = build_xlsx(tpl_bytes, cards, info, band)
        pngs, imgs = [], []
        for i, c in enumerate(cards):
            img = renderer.render(c, vendor=vendor, vendor_fill=vendor_fill, goal=c["goal"], band=band)
            imgs.append(img)
            pngs.append((f"{i + 1:03d}_{safe_name(c['style'])}.png", to_png_bytes(img)))
            prog.progress(0.1 + 0.75 * (i + 1) / len(cards), text=f"繪製卡片 {i + 1}/{len(cards)}…")
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
        st.success(f"完成：{out['n']} 張卡片" + (f"（含縮圖 {out['thumbs']} 張）" if out["thumbs"] else ""))
        d1, d2, d3 = st.columns(3)
        d1.download_button("⬇️ Excel（Sample Tag）", out["xlsx"], f"{out['stem']}.xlsx", XLSX_MIME,
                           key=f"{code}_dlx", width="stretch")
        d2.download_button("⬇️ 卡片圖片 PNG（ZIP）", out["zip"], f"{out['stem']} PNG.zip", "application/zip",
                           key=f"{code}_dlz", width="stretch")
        d3.download_button("⬇️ 列印 PDF（A4，每頁 8 張）", out["pdf"], f"{out['stem']}.pdf", "application/pdf",
                           key=f"{code}_dlp", width="stretch")


# ---------------------------------------------------------------------------
# Dropdown tab
# ---------------------------------------------------------------------------
def _goal_table(info) -> pd.DataFrame:
    return pd.DataFrame([{"Program": k, info.imu_header: v.get("imu"), info.imu_duty_header: v.get("imu_duty")}
                         for k, v in info.programs.items()])


def dropdown_tab(cfg):
    code = cfg.code
    st.markdown("把 Target 新一季範本（SOURCE）的 **Dropdowns**（Program、IMU 目標、廠商清單）更新到你的工作檔（TEMPLATE）。"
                "只改 Dropdowns 工作表內容，其他工作表、巨集、格式都原封不動。")
    c1, c2 = st.columns(2)
    with c1:
        t = st.file_uploader("TEMPLATE（要被更新的工作檔）", type=["xlsx", "xlsm"], key=f"{code}_dd_t")
    with c2:
        s = st.file_uploader("SOURCE（Target 新一季範本）", type=["xlsx", "xlsm"], key=f"{code}_dd_s")

    builtin = cached_template_info(cfg.template_path.read_bytes())
    with st.expander(f"App 內建 {code} 範本目前的 IMU 目標（{cfg.template_file}）"):
        st.dataframe(_goal_table(builtin), hide_index=True, width="stretch",
                     column_config={builtin.imu_header: st.column_config.NumberColumn(format="percent"),
                                    builtin.imu_duty_header: st.column_config.NumberColumn(format="percent")})
        st.caption(f"黃燈區間 {(builtin.yellow_band or cfg.default_yellow_band):.0%}；廠商 {len(builtin.vendors)} 家。")

    if s is None:
        return
    src = s.getvalue()
    sinfo = cached_template_info(src)
    base_info = cached_template_info(t.getvalue()) if t is not None else builtin
    d = diff_templates(base_info, sinfo)

    st.subheader("差異")
    a, b, c = st.columns(3)
    a.metric("新增 Program", len(d["programs_added"]))
    b.metric("移除 Program", len(d["programs_removed"]))
    c.metric("IMU 目標變動", len(d["goals_changed"]))
    if d["programs_added"]:
        st.write("新增：", "、".join(d["programs_added"]))
    if d["programs_removed"]:
        st.write("移除：", "、".join(d["programs_removed"]))
    if d["goals_changed"]:
        names = {"imu": sinfo.imu_header, "imu_duty": sinfo.imu_duty_header}
        chg = pd.DataFrame([(p, names.get(k, k), a, b) for p, k, a, b in d["goals_changed"]],
                           columns=["Program", "欄位", "原本", "更新後"])
        st.dataframe(chg, hide_index=True, width="stretch",
                     column_config={"原本": st.column_config.NumberColumn(format="percent"),
                                    "更新後": st.column_config.NumberColumn(format="percent")})
    if d["vendors_added"] or d["vendors_removed"]:
        st.write("廠商新增：", "、".join(d["vendors_added"]) or "—", "　｜　廠商移除：", "、".join(d["vendors_removed"]) or "—")
    if d["band_old"] != d["band_new"]:
        st.write(f"黃燈區間：{d['band_old']} → {d['band_new']}")
    if not any([d["programs_added"], d["programs_removed"], d["goals_changed"], d["vendors_added"],
                d["vendors_removed"]]) and d["band_old"] == d["band_new"]:
        st.success("兩份 Dropdowns 內容相同，不需要更新。")

    st.dataframe(_goal_table(sinfo), hide_index=True, width="stretch",
                 column_config={sinfo.imu_header: st.column_config.NumberColumn(format="percent"),
                                sinfo.imu_duty_header: st.column_config.NumberColumn(format="percent")})

    b1, b2 = st.columns(2)
    if t is not None:
        if b1.button("更新 TEMPLATE 的 Dropdowns", type="primary", key=f"{code}_dd_go"):
            try:
                out = update_dropdowns(t.getvalue(), src)
                st.session_state[f"{code}_dd_out"] = (t.name, out)
            except Exception as e:
                st.error(f"更新失敗：{e}")
        res = st.session_state.get(f"{code}_dd_out")
        if res and res[0] == t.name:
            name, out = res
            stem, ext = os.path.splitext(name)
            b1.download_button("⬇️ 下載更新後的 TEMPLATE", out, f"{stem}_Dropdowns更新{ext}",
                               XLSM_MIME if ext.lower() == ".xlsm" else XLSX_MIME, key=f"{code}_dd_dl")
    if b2.button("用這份 SOURCE 產生本次 Sample Tag", key=f"{code}_dd_use",
                 help="本次登入期間，『產生 Sample Tag』改用這份範本（不影響內建範本）"):
        st.session_state[f"{code}_template_override"] = (src, s.name)
        st.success(f"已切換：{s.name}")


# ---------------------------------------------------------------------------
def main():
    if not gate():
        return
    st.title("🏷️ Sample Tag 產生器")
    st.caption("上傳 PCN ➜ 勾選品項 ➜ 一鍵產生 Sample Tag Excel、每張卡的 PNG 與列印 PDF。條碼為內嵌圖片，不需安裝字型、不需連網。")
    tabs = st.tabs([p.title for p in PROGRAMS.values()])
    for tab, cfg in zip(tabs, PROGRAMS.values()):
        with tab:
            t1, t2 = st.tabs(["產生 Sample Tag", "Dropdown 更新"])
            with t1:
                sample_tag_tab(cfg)
            with t2:
                dropdown_tab(cfg)


main()
