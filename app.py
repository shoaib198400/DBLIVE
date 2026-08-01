import base64
import hashlib
import inspect
import pandas as pd
import html
import os
from io import BytesIO
from datetime import datetime
import plotly.express as px
import streamlit as st
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def render_location_visit_observation_detail() -> None:
    """Show raw CAPA-level observation/recommendation detail for a selected audit."""
    audit_no   = st.session_state.get("lv_obs_audit_no", "")
    plant      = st.session_state.get("lv_obs_plant", "")
    plant_desc = st.session_state.get("lv_obs_plant_desc", "")
    zone       = st.session_state.get("lv_obs_zone", "")
    fy         = st.session_state.get("lv_obs_fy", "")
    quarter    = st.session_state.get("lv_obs_quarter", "")

    _bcols = st.columns([1, 5])
    if _bcols[0].button("⬅ Back to Audit List", key="btn_back_obs_detail", use_container_width=True):
        st.session_state["lv_sub_page"] = "summary"
        st.rerun()

    st.markdown(
        f"<div class='sec-title'>&#128203; Observation Detail &mdash; "
        f"{html.escape(str(plant_desc))} ({html.escape(str(plant))})</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='background:#f0f7ff;border-left:3px solid #003087;border-radius:5px;"
        f"padding:6px 14px;margin:4px 0 10px 0;font-size:13px;'>"
        f"<b>Zone:</b> {html.escape(str(zone))} &nbsp;&#124;&nbsp; "
        f"<b>Period:</b> {quarter} FY{fy} &nbsp;&#124;&nbsp; "
        f"<b>Audit No.:</b> {html.escape(str(audit_no))}</div>",
        unsafe_allow_html=True,
    )

    if not os.path.exists(LOCATION_VISIT_PATH):
        st.warning("Location Visit file not found.")
        return

    try:
        _ext = os.path.splitext(LOCATION_VISIT_PATH)[1].lower()
        try:
            df_raw = pd.read_excel(LOCATION_VISIT_PATH, engine="xlrd" if _ext == ".xls" else "openpyxl")
        except Exception:
            df_raw = pd.read_excel(LOCATION_VISIT_PATH, engine="openpyxl" if _ext == ".xls" else "xlrd")
        df_raw.columns = df_raw.columns.astype(str).str.strip()
    except Exception as e:
        st.error(f"Could not load raw data: {e}")
        return

    _plant_col = next((c for c in ["Planning Plant", "Plant Code"] if c in df_raw.columns), None)
    _audit_col = next((c for c in ["Audit Number", "AuditNumber"] if c in df_raw.columns), None)
    _capa_col  = next((c for c in ["CAPA Status", "Capa Status"] if c in df_raw.columns), None)

    if not _audit_col:
        st.warning("Cannot find Audit Number column in raw file.")
        return

    mask = df_raw[_audit_col].astype(str).str.strip() == str(audit_no).strip()
    if _plant_col:
        mask &= (
            df_raw[_plant_col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            == str(plant).strip()
        )
    df_audit = df_raw[mask].copy()

    if df_audit.empty:
        st.warning(f"No detail rows found for Audit {audit_no} / Plant {plant}.")
        return

    # ── Colored KPI cards ────────────────────────────────────────────────────
    _CLOSED = {"Closed", "Completed"}
    _OPEN   = {"Open", "Reopened"}
    if _capa_col:
        statuses = df_audit[_capa_col].astype(str).str.strip()
        valid    = statuses[statuses.isin(_CLOSED | _OPEN)]
        total_c  = len(valid)
        closed_c = int(statuses.isin(_CLOSED).sum())
        open_c   = int(statuses.isin(_OPEN).sum())
    else:
        total_c = closed_c = open_c = 0
    comp_pct  = (closed_c / total_c * 100) if total_c > 0 else 0.0
    comp_icon = "✅" if comp_pct >= 75 else ("⚠️" if comp_pct >= 50 else "❌")
    _cc = "#2e7d32" if comp_pct >= 75 else ("#F57C00" if comp_pct >= 50 else "#c62828")

    _obs_kpi_html = f"""
    <div style="display:flex;gap:10px;margin:6px 0 14px 0;flex-wrap:wrap;">
      <div style="flex:1;min-width:120px;background:linear-gradient(135deg,#003087,#0057A8);border-radius:9px;padding:14px 10px;text-align:center;color:#fff;box-shadow:0 3px 10px rgba(0,48,135,0.25);">
        <div style="font-size:28px;font-weight:800;">{total_c}</div>
        <div style="font-size:12px;font-weight:600;margin-top:4px;opacity:.9;">&#128203; Total Recommendations</div>
      </div>
      <div style="flex:1;min-width:120px;background:linear-gradient(135deg,#c62828,#d32f2f);border-radius:9px;padding:14px 10px;text-align:center;color:#fff;box-shadow:0 3px 10px rgba(198,40,40,0.25);">
        <div style="font-size:28px;font-weight:800;">{open_c}</div>
        <div style="font-size:12px;font-weight:600;margin-top:4px;opacity:.9;">&#128308; Open / Pending</div>
      </div>
      <div style="flex:1;min-width:120px;background:linear-gradient(135deg,#2e7d32,#388e3c);border-radius:9px;padding:14px 10px;text-align:center;color:#fff;box-shadow:0 3px 10px rgba(46,125,50,0.25);">
        <div style="font-size:28px;font-weight:800;">{closed_c}</div>
        <div style="font-size:12px;font-weight:600;margin-top:4px;opacity:.9;">&#128994; Closed / Complied</div>
      </div>
      <div style="flex:1;min-width:120px;background:linear-gradient(135deg,{_cc},{_cc}cc);border-radius:9px;padding:14px 10px;text-align:center;color:#fff;box-shadow:0 3px 10px rgba(0,0,0,0.18);">
        <div style="font-size:28px;font-weight:800;">{comp_pct:.1f}%</div>
        <div style="font-size:12px;font-weight:600;margin-top:4px;opacity:.9;">{comp_icon} Compliance</div>
      </div>
    </div>"""
    st.markdown(_obs_kpi_html, unsafe_allow_html=True)

    # ── Select display columns (remove Risk Category, Severity, Revised Target Date, Root Cause) ──
    _obs_cols = [c for c in [
        "CAPA Number", "Observation Comment", "Recommendation",
        "CAPA Status", "Target Date",
        "Action Taken Current Status",
        "Action Closed by Name", "Action Closed Date",
    ] if c in df_audit.columns]
    if not _obs_cols:
        _obs_cols = [c for c in df_audit.columns
                     if c not in {"Risk Category","Severity Code","Revised Target Date","Root Cause Analysis"}][:10]

    df_obs = df_audit[_obs_cols].reset_index(drop=True)

    if "CAPA Status" in df_obs.columns:
        def _fmt_s(v):
            v = str(v).strip()
            if v in ("Closed", "Completed"): return f"✅ {v}"
            if v == "Open":                  return f"🔴 {v}"
            if v == "Reopened":              return f"⚠️ {v}"
            return v
        df_obs["CAPA Status"] = df_obs["CAPA Status"].apply(_fmt_s)

    # ── Custom HTML table: headers centered, data left+top aligned ────────────
    _LONG_COLS = {"Observation Comment", "Recommendation", "Action Taken Current Status"}
    _hth = "background:#003087;color:#fff;font-weight:700;text-align:center;padding:8px 10px;font-size:12px;border-right:1px solid #1a5fb4;white-space:nowrap;"

    hdr_html = f"<th style='{_hth}'>S.No</th>"
    for c in df_obs.columns:
        hdr_html += f"<th style='{_hth}'>{html.escape(str(c))}</th>"

    rows_html = ""
    for i, (_, row) in enumerate(df_obs.iterrows()):
        _bg = "#ffffff" if i % 2 == 0 else "#f0f7ff"
        row_html = f"<tr style='background:{_bg};vertical-align:top;'>"
        row_html += f"<td style='text-align:center;font-weight:700;padding:7px 8px;font-size:12px;border-bottom:1px solid #e2eaf4;border-right:1px solid #e2eaf4;white-space:nowrap;color:#003087;'>{i+1}</td>"
        for col, val in zip(df_obs.columns, row):
            val_str = "" if pd.isna(val) else str(val)
            if col in _LONG_COLS:
                row_html += (
                    f"<td style='text-align:left;vertical-align:top;padding:7px 10px;font-size:12px;"
                    f"border-bottom:1px solid #e2eaf4;border-right:1px solid #e2eaf4;"
                    f"min-width:180px;max-width:320px;word-wrap:break-word;white-space:pre-wrap;'>"
                    f"{html.escape(val_str)}</td>"
                )
            else:
                row_html += (
                    f"<td style='text-align:left;vertical-align:top;padding:7px 10px;font-size:12px;"
                    f"border-bottom:1px solid #e2eaf4;border-right:1px solid #e2eaf4;white-space:nowrap;'>"
                    f"{html.escape(val_str)}</td>"
                )
        row_html += "</tr>"
        rows_html += row_html

    obs_table_html = (
        f"<p style='font-weight:700;color:#003087;font-size:13px;margin:6px 0 4px 0;'>"
        f"&#128196; {len(df_obs)} Recommendation(s) for this Audit</p>"
        f"<div style='overflow:auto;max-height:650px;border-radius:8px;"
        f"border:1px solid #d5e2f3;background:#fff;box-shadow:0 2px 8px #e0e0e0;margin:4px 0 10px 0;'>"
        f"<table style='border-collapse:collapse;width:100%;font-size:12px;'>"
        f"<thead><tr>{hdr_html}</tr></thead>"
        f"<tbody>{rows_html}</tbody>"
        f"</table></div>"
    )
    st.markdown(obs_table_html, unsafe_allow_html=True)

    _download_excel_button(
        label="⬇  Download Observation Detail  (.xlsx)",
        file_prefix=f"obs_{str(audit_no).replace('/', '_').replace(' ', '_')}",
        sheets={"Observations": df_audit[_obs_cols].reset_index(drop=True)},
        key="dl_obs_detail",
    )


def render_location_visit_details(df: pd.DataFrame) -> None:
    """
    Level-1 drill-down: Location Visit & Compliance Analysis.
    df: aggregated audit-level data (one row per plant+audit), with FY and Quarter columns.
    """
    if st.session_state.get("lv_sub_page") == "obs_detail":
        render_location_visit_observation_detail()
        return

    st.markdown(
        "<div class='sec-title'>&#128205; Location Visit &amp; Compliance Analysis</div>",
        unsafe_allow_html=True,
    )
    _bcols = st.columns([1, 5])
    if _bcols[0].button("⬅ Back to Dashboard", key="btn_back_location_visit", use_container_width=True):
        st.session_state["location_visit_page"] = "main"
        st.session_state["lv_sub_page"] = "summary"
        st.session_state["selected_tile"] = None
        st.session_state["dummy_tank_clicked"] = False
        st.session_state["pl_unblock_clicked"] = False
        st.session_state["tank_turns_page"] = "main"
        st.rerun()

    if df is None or df.empty:
        st.warning("No data available. Check file upload or data source.")
        return

    if "FY" not in df.columns or "Quarter" not in df.columns:
        def _fy_q(dt):
            if pd.isna(dt): return ("Unknown", "Unknown")
            m, y = dt.month, dt.year
            if m >= 4:
                return (str(y)[2:] + "-" + str(y+1)[2:], "Q" + str(((m-4)//3)+1))
            return (str(y-1)[2:] + "-" + str(y)[2:], "Q4")
        _dates = pd.to_datetime(df.get("Audit Start Date"), errors="coerce", dayfirst=True)
        _fqs = _dates.apply(_fy_q)
        df = df.copy()
        df["FY"]      = _fqs.apply(lambda x: x[0])
        df["Quarter"] = _fqs.apply(lambda x: x[1])

    # ── Filters ───────────────────────────────────────────────────────────────
    st.markdown("<div class='sec-title'>&#128269; Filters</div>", unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns(4)
    _zones = sorted(df["Zone"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
    sel_zone = f1.selectbox("Zone", ["All"] + _zones, key="lv_f_zone")
    _dz = df if sel_zone == "All" else df[df["Zone"].astype(str).str.strip() == sel_zone]
    _plants = sorted(_dz["Plant Desc."].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
    sel_plant = f2.selectbox("Location", ["All"] + _plants, key="lv_f_plant")
    _fy_opts = sorted(df["FY"].fillna("Unknown").astype(str).unique().tolist(), reverse=True)
    sel_fy = f3.selectbox("Financial Year", ["All"] + _fy_opts, key="lv_f_fy")
    sel_qtr = f4.selectbox("Quarter", ["All", "Q1", "Q2", "Q3", "Q4"], key="lv_f_qtr")

    dv = df.copy()
    if sel_zone  != "All": dv = dv[dv["Zone"].astype(str).str.strip() == sel_zone]
    if sel_plant != "All": dv = dv[dv["Plant Desc."].astype(str).str.strip() == sel_plant]
    if sel_fy    != "All": dv = dv[dv["FY"].astype(str) == sel_fy]
    if sel_qtr   != "All": dv = dv[dv["Quarter"].astype(str) == sel_qtr]

    if dv.empty:
        st.info("No data matches the selected filters.")
        return

    def _n(s): return pd.to_numeric(s, errors="coerce").fillna(0)

    total_r  = int(_n(dv["TotalRecomms"]).sum())
    closed_r = int(_n(dv["ClosedRecomms"]).sum())
    open_r   = int(_n(dv["OpenRecomms"]).sum())
    audits   = len(dv)
    comp_pct = (closed_r / total_r * 100) if total_r > 0 else 0.0
    comp_icon = "✅" if comp_pct >= 75 else ("⚠️" if comp_pct >= 50 else "❌")

    def _comp_color(pct):
        return "#2e7d32" if pct >= 75 else ("#F57C00" if pct >= 50 else "#c62828")

    _comp_clr = _comp_color(comp_pct)

    # ── Section index (quick navigation) ─────────────────────────────────────
    _lnk = ("color:#003087;font-weight:600;font-size:12px;text-decoration:none;"
             "background:#ddeeff;padding:3px 10px;border-radius:12px;white-space:nowrap;")
    _toc_html = (
        "<div style='background:#f0f7ff;border-left:4px solid #003087;border-radius:6px;"
        "padding:8px 14px;margin:6px 0 12px 0;'>"
        "<span style='font-weight:700;color:#003087;font-size:12px;'>&#128196; Quick Navigation:&nbsp;</span>"
        f"<a href='#lv-kpi' style='{_lnk}'>&#128202; KPI Summary</a>&nbsp;"
        f"<a href='#lv-zone' style='{_lnk}'>&#128506; Zone Summary</a>&nbsp;"
        f"<a href='#lv-quarter' style='{_lnk}'>&#128197; Quarter Summary</a>&nbsp;"
        f"<a href='#lv-audit' style='{_lnk}'>&#128203; Audit Detail</a>&nbsp;"
        f"<a href='#lv-perf-zone' style='{_lnk}'>&#128202; Zone Analysis</a>&nbsp;"
        f"<a href='#lv-perf-loc' style='{_lnk}'>&#128205; Location Analysis</a>&nbsp;"
        f"<a href='#lv-rank-zone' style='{_lnk}'>&#127941; Zone Ranking</a>&nbsp;"
        f"<a href='#lv-rank-loc' style='{_lnk}'>&#127941; Location Ranking</a>&nbsp;"
        f"<a href='#lv-missing' style='{_lnk}'>&#9888; Missing Locations</a>"
        "</div>"
    )
    st.markdown(_toc_html, unsafe_allow_html=True)

    # ── 5-card KPI summary ────────────────────────────────────────────────────
    st.markdown("<a id='lv-kpi'></a><div class='sec-title'>&#128202; Overall Summary</div>", unsafe_allow_html=True)
    _cards_html = f"""
    <div style="display:flex;gap:10px;margin:6px 0 12px 0;flex-wrap:wrap;">
      <div style="flex:1;min-width:110px;background:linear-gradient(135deg,#003087,#0057A8);border-radius:9px;padding:13px 8px;text-align:center;color:#fff;box-shadow:0 3px 10px rgba(0,48,135,0.2);">
        <div style="font-size:26px;font-weight:800;">{audits:,}</div>
        <div style="font-size:12px;font-weight:600;margin-top:4px;opacity:.9;">&#128506; Audited Locations</div>
      </div>
      <div style="flex:1;min-width:110px;background:linear-gradient(135deg,#37474F,#546E7A);border-radius:9px;padding:13px 8px;text-align:center;color:#fff;box-shadow:0 3px 10px rgba(0,0,0,0.15);">
        <div style="font-size:26px;font-weight:800;">{total_r:,}</div>
        <div style="font-size:12px;font-weight:600;margin-top:4px;opacity:.9;">&#128203; Total Recommendations</div>
      </div>
      <div style="flex:1;min-width:110px;background:linear-gradient(135deg,#c62828,#d32f2f);border-radius:9px;padding:13px 8px;text-align:center;color:#fff;box-shadow:0 3px 10px rgba(198,40,40,0.22);">
        <div style="font-size:26px;font-weight:800;">{open_r:,}</div>
        <div style="font-size:12px;font-weight:600;margin-top:4px;opacity:.9;">&#128308; Open Recommendations</div>
      </div>
      <div style="flex:1;min-width:110px;background:linear-gradient(135deg,#2e7d32,#388e3c);border-radius:9px;padding:13px 8px;text-align:center;color:#fff;box-shadow:0 3px 10px rgba(46,125,50,0.22);">
        <div style="font-size:26px;font-weight:800;">{closed_r:,}</div>
        <div style="font-size:12px;font-weight:600;margin-top:4px;opacity:.9;">&#128994; Closed / Complied</div>
      </div>
      <div style="flex:1;min-width:110px;background:linear-gradient(135deg,{_comp_clr},{_comp_clr}cc);border-radius:9px;padding:13px 8px;text-align:center;color:#fff;box-shadow:0 3px 10px rgba(0,0,0,0.18);">
        <div style="font-size:26px;font-weight:800;">{comp_pct:.1f}%</div>
        <div style="font-size:12px;font-weight:600;margin-top:4px;opacity:.9;">{comp_icon} Overall Compliance</div>
      </div>
    </div>"""
    st.markdown(_cards_html, unsafe_allow_html=True)

    # ── Shared helpers ────────────────────────────────────────────────────────
    _ZONE_COLORS = {
        "COZ": "#E65100", "NOZ": "#003087", "SOZ": "#2e7d32",
        "WOZ": "#AD1457", "EOZ": "#6A1B9A", "Unmapped": "#546E7A",
    }
    def _zone_badge(z):
        z_up = str(z).strip().upper()
        clr = _ZONE_COLORS.get(z_up, "#546E7A")
        return f"<span style='background:{clr};color:#fff;font-weight:700;padding:2px 8px;border-radius:10px;font-size:11px;letter-spacing:.4px;'>{html.escape(z_up)}</span>"

    def _comp_badge(pct):
        clr = _comp_color(pct)
        return f"<span style='background:{clr};color:#fff;font-weight:700;padding:2px 8px;border-radius:10px;font-size:11px;'>{pct:.1f}%</span>"

    def _progress_bar(pct):
        fill = min(max(pct, 0), 100)
        clr = _comp_color(pct)
        return (
            f"<div style='background:#e0e0e0;border-radius:6px;height:11px;width:100%;min-width:80px;'>"
            f"<div style='background:{clr};width:{fill:.0f}%;height:11px;border-radius:6px;'></div></div>"
        )

    _th = "padding:7px 10px;font-weight:700;border-bottom:2px solid #d5e2f3;font-size:12px;"

    def _rank_badge(rank):
        if rank == 1:
            return "<span style='background:linear-gradient(135deg,#DAA520,#FFD700);color:#5a3200;font-weight:900;padding:2px 9px;border-radius:6px;font-size:12px;'>&#11088; 1st</span>"
        if rank == 2:
            return "<span style='background:linear-gradient(135deg,#9E9E9E,#BDBDBD);color:#1a1a1a;font-weight:900;padding:2px 9px;border-radius:6px;font-size:12px;'>&#9733; 2nd</span>"
        if rank == 3:
            return "<span style='background:linear-gradient(135deg,#8D4E1B,#CD7F32);color:#fff;font-weight:900;padding:2px 9px;border-radius:6px;font-size:12px;'>&#9733; 3rd</span>"
        return f"<span style='color:#555;font-size:12px;font-weight:600;'>{rank}</span>"

    # ── Zone-wise Summary ─────────────────────────────────────────────────────
    st.markdown("<a id='lv-zone'></a>", unsafe_allow_html=True)
    st.markdown("<div class='sec-title'>&#128506; Zone-wise Summary</div>", unsafe_allow_html=True)
    _zg = dv.assign(_t=_n(dv["TotalRecomms"]), _c=_n(dv["ClosedRecomms"]), _o=_n(dv["OpenRecomms"]))
    zone_tbl = (
        _zg.groupby("Zone", as_index=False)
        .agg(Locations=("Planning Plant","count"), Total=("_t","sum"), Open=("_o","sum"), Closed=("_c","sum"))
    )
    zone_tbl["_comp"] = zone_tbl.apply(lambda r: r.Closed/r.Total*100 if r.Total > 0 else 0.0, axis=1)
    zone_tbl = zone_tbl.sort_values("_comp", ascending=False).reset_index(drop=True)

    _zone_rows = ""
    for i, r in zone_tbl.iterrows():
        _bg = "#ffffff" if i % 2 == 0 else "#f7fafd"
        _zone_rows += (
            f"<tr style='background:{_bg};'>"
            f"<td style='padding:6px 10px;'>{_zone_badge(r.Zone)}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:600;'>{int(r.Locations)}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:600;'>{int(r.Total):,}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:700;color:#c62828;font-size:13px;'>{int(r.Open):,}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:700;color:#2e7d32;font-size:13px;'>{int(r.Closed):,}</td>"
            f"<td style='padding:6px 10px;text-align:center;'>{_comp_badge(r._comp)}</td>"
            f"<td style='padding:6px 12px;min-width:90px;'>{_progress_bar(r._comp)}</td>"
            f"</tr>"
        )
    # Grand total row
    _gt_comp = closed_r / total_r * 100 if total_r > 0 else 0.0
    _zone_rows += (
        f"<tr style='background:#e8f0fe;font-weight:800;border-top:2px solid #003087;'>"
        f"<td style='padding:7px 10px;color:#003087;font-weight:800;font-size:12px;'>&#9646; TOTAL / OVERALL</td>"
        f"<td style='padding:7px 10px;text-align:center;color:#003087;'>{audits}</td>"
        f"<td style='padding:7px 10px;text-align:center;color:#003087;'>{total_r:,}</td>"
        f"<td style='padding:7px 10px;text-align:center;color:#c62828;'>{open_r:,}</td>"
        f"<td style='padding:7px 10px;text-align:center;color:#2e7d32;'>{closed_r:,}</td>"
        f"<td style='padding:7px 10px;text-align:center;'>{_comp_badge(_gt_comp)}</td>"
        f"<td style='padding:7px 12px;min-width:90px;'>{_progress_bar(_gt_comp)}</td>"
        f"</tr>"
    )
    _zone_html = (
        "<div style='overflow-x:auto;margin:6px 0 14px 0;'>"
        "<table style='border-collapse:collapse;width:100%;font-size:12px;'>"
        f"<thead><tr style='background:#eaf2fb;'>"
        f"<th style='{_th}text-align:left;color:#003087;'>Zone</th>"
        f"<th style='{_th}text-align:center;color:#003087;'>Locations</th>"
        f"<th style='{_th}text-align:center;color:#003087;'>Total Recomms</th>"
        f"<th style='{_th}text-align:center;color:#c62828;'>Open</th>"
        f"<th style='{_th}text-align:center;color:#2e7d32;'>Closed</th>"
        f"<th style='{_th}text-align:center;color:#003087;'>Compliance %</th>"
        f"<th style='{_th}text-align:left;color:#003087;'>Progress</th>"
        f"</tr></thead>"
        f"<tbody>{_zone_rows}</tbody>"
        "</table></div>"
    )
    st.markdown(_zone_html, unsafe_allow_html=True)

    # ── Quarter-wise Summary ──────────────────────────────────────────────────
    st.markdown("<a id='lv-quarter'></a>", unsafe_allow_html=True)
    st.markdown("<div class='sec-title'>&#128197; Quarter-wise Summary</div>", unsafe_allow_html=True)
    _qg = dv.assign(_t=_n(dv["TotalRecomms"]), _c=_n(dv["ClosedRecomms"]), _o=_n(dv["OpenRecomms"]))
    qtr_tbl = (
        _qg.groupby(["FY", "Quarter"], as_index=False)
        .agg(Locations=("Planning Plant","count"), Total=("_t","sum"), Open=("_o","sum"), Closed=("_c","sum"))
        .sort_values(["FY","Quarter"])
        .reset_index(drop=True)
    )
    qtr_tbl["Compliance %"] = qtr_tbl.apply(
        lambda r: f"{r.Closed/r.Total*100:.1f}%" if r.Total > 0 else "N/A", axis=1
    )
    _render_html_table(qtr_tbl, max_height=300)

    # ── Audit Detail Table ────────────────────────────────────────────────────
    st.markdown("<a id='lv-audit'></a>", unsafe_allow_html=True)
    st.markdown("<div class='sec-title'>&#128203; Audit Detail by Location</div>", unsafe_allow_html=True)
    st.caption("Sorted by compliance (highest first). Select an audit and click 'View Observations' for full details.")

    # Build audit table: exclude FY and Quarter (already in filter), add S.No, sort by compliance desc
    _req = ["Planning Plant", "Plant Desc.", "Zone",
            "Audit Number", "Audit Start Date", "Audit End Date",
            "TotalRecomms", "OpenRecomms", "ClosedRecomms"]
    audit_tbl = dv[[c for c in _req if c in dv.columns]].copy()
    audit_tbl["_comp_num"] = audit_tbl.apply(
        lambda r: _n(pd.Series([r["ClosedRecomms"]])).sum() / max(_n(pd.Series([r["TotalRecomms"]])).sum(), 1) * 100,
        axis=1,
    )
    audit_tbl["Compliance %"] = audit_tbl["_comp_num"].apply(lambda x: f"{x:.1f}%")
    audit_tbl = audit_tbl.sort_values("_comp_num", ascending=False).drop(columns=["_comp_num"]).reset_index(drop=True)
    audit_tbl.insert(0, "S.No", range(1, len(audit_tbl)+1))
    # Keep FY and Quarter hidden in display but available for navigation
    audit_tbl_display = audit_tbl.copy()
    _render_html_table(audit_tbl_display, max_height=420)

    # Rebuild audit_tbl_with_fy for navigation (merge FY/Quarter back from dv)
    _audit_keys = [
        f"{row['Audit Number']} | {row['Plant Desc.']} ({row['Planning Plant']})"
        for _, row in audit_tbl_display.iterrows()
    ]
    obs_c1, obs_c2 = st.columns([4, 1])
    sel_audit_key = obs_c1.selectbox(
        "Select an audit to view observations:",
        _audit_keys if _audit_keys else ["(no audits)"],
        key="lv_sel_audit",
    )
    if obs_c2.button("&#128203; View Observations", key="btn_view_obs", use_container_width=True):
        if _audit_keys:
            _idx = _audit_keys.index(sel_audit_key)
            _row = audit_tbl_display.iloc[_idx]
            # Get FY/Quarter from dv by matching Plant+Audit
            _match = dv[
                (dv["Planning Plant"].astype(str) == str(_row["Planning Plant"])) &
                (dv["Audit Number"].astype(str) == str(_row["Audit Number"]))
            ]
            _fy_val  = _match["FY"].iloc[0]  if not _match.empty and "FY"      in _match.columns else ""
            _qtr_val = _match["Quarter"].iloc[0] if not _match.empty and "Quarter" in _match.columns else ""
            st.session_state["lv_obs_audit_no"]   = _row["Audit Number"]
            st.session_state["lv_obs_plant"]      = _row["Planning Plant"]
            st.session_state["lv_obs_plant_desc"] = _row["Plant Desc."]
            st.session_state["lv_obs_zone"]       = _row["Zone"]
            st.session_state["lv_obs_fy"]         = _fy_val
            st.session_state["lv_obs_quarter"]    = _qtr_val
            st.session_state["lv_sub_page"]       = "obs_detail"
            st.rerun()

    # ── Performance Analysis — Zone Chart (full width) ────────────────────────
    st.markdown("<a id='lv-perf-zone'></a>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sec-title'>&#128202; Performance Analysis &mdash; Zone-wise Compliance</div>",
        unsafe_allow_html=True,
    )
    _zone_c = zone_tbl.copy()
    _zone_c["Zone_Label"] = _zone_c["Zone"].astype(str).str.upper()
    fig_zone = px.bar(
        _zone_c, x="Zone_Label", y="_comp",
        color="_comp", color_continuous_scale=["#c62828","#f57c00","#2e7d32"],
        range_color=[0, 100],
        text=_zone_c["_comp"].apply(lambda x: f"{x:.1f}%"),
        title="Compliance % by Zone  (75% dashed line = target)",
        labels={"_comp": "Compliance %", "Zone_Label": "Zone"},
    )
    fig_zone.add_hline(
        y=75, line_dash="dash", line_color="#003087", line_width=2,
        annotation_text="<b>75% Target</b>", annotation_position="top right",
        annotation_font_color="#003087", annotation_font_size=13,
    )
    fig_zone.update_traces(textposition="outside", textfont_size=13, textfont_color="#1a1a1a")
    fig_zone.update_layout(
        height=420, coloraxis_showscale=False,
        margin=dict(l=30, r=30, t=60, b=60),
        yaxis=dict(title="Compliance %", range=[0, 120], tickfont=dict(size=12, color="#1a1a1a")),
        xaxis=dict(tickfont=dict(size=13, color="#111111", family="Arial Black")),
        title_font=dict(size=14, color="#003087"),
        plot_bgcolor="#f8fafd", paper_bgcolor="#f8fafd",
    )
    st.plotly_chart(fig_zone, use_container_width=True)

    # ── Performance Analysis — Location Chart (full width) ────────────────────
    st.markdown("<a id='lv-perf-loc'></a>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sec-title'>&#128205; Performance Analysis &mdash; Location-wise Compliance</div>",
        unsafe_allow_html=True,
    )
    _lg = dv.assign(_t=_n(dv["TotalRecomms"]), _c=_n(dv["ClosedRecomms"]), _o=_n(dv["OpenRecomms"]))
    loc_chart = (
        _lg.groupby(["Planning Plant","Plant Desc.","Zone"], as_index=False)
        .agg(Total=("_t","sum"), Closed=("_c","sum"), Open=("_o","sum"))
    )
    loc_chart["_pct"] = loc_chart.apply(
        lambda r: r.Closed/r.Total*100 if r.Total > 0 else 0.0, axis=1
    )
    loc_chart["Label"] = loc_chart["Plant Desc."].astype(str).str[:28]
    loc_chart = loc_chart.sort_values("_pct")
    _loc_h = max(500, len(loc_chart) * 28)
    fig_loc = px.bar(
        loc_chart, x="_pct", y="Label",
        orientation="h",
        color="_pct", color_continuous_scale=["#c62828","#f57c00","#2e7d32"],
        range_color=[0, 100],
        text=loc_chart["_pct"].apply(lambda x: f"{x:.1f}%"),
        title="Location-wise Compliance %  (75% dashed line = target)",
        labels={"_pct": "Compliance %", "Label": ""},
    )
    fig_loc.add_vline(
        x=75, line_dash="dash", line_color="#003087", line_width=2,
        annotation_text="<b>75%</b>", annotation_position="top right",
        annotation_font_color="#003087", annotation_font_size=13,
    )
    fig_loc.update_traces(textposition="outside", textfont_size=11)
    fig_loc.update_layout(
        height=_loc_h, coloraxis_showscale=False,
        margin=dict(l=20, r=60, t=60, b=30),
        xaxis=dict(range=[0, 125], tickfont=dict(size=12, color="#1a1a1a")),
        yaxis=dict(tickfont=dict(size=11, color="#111111")),
        title_font=dict(size=14, color="#003087"),
        plot_bgcolor="#f8fafd", paper_bgcolor="#f8fafd",
    )
    st.plotly_chart(fig_loc, use_container_width=True)

    # ── Zone-wise Compliance Ranking ──────────────────────────────────────────
    st.markdown("<a id='lv-rank-zone'></a>", unsafe_allow_html=True)
    st.markdown("<div class='sec-title'>&#127941; Zone-wise Compliance Ranking</div>", unsafe_allow_html=True)
    _zr = zone_tbl.copy().sort_values("_comp", ascending=False).reset_index(drop=True)
    _zrank_rows = ""
    for i, r in _zr.iterrows():
        rank = i + 1
        _bg = "#fffde7" if rank == 1 else ("#f3f3f3" if rank == 2 else ("#fdf3ee" if rank == 3 else ("#ffffff" if rank % 2 == 1 else "#f7fafd")))
        _zrank_rows += (
            f"<tr style='background:{_bg};'>"
            f"<td style='padding:6px 10px;text-align:center;'>{_rank_badge(rank)}</td>"
            f"<td style='padding:6px 10px;'>{_zone_badge(r.Zone)}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:600;'>{int(r.Locations)}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:700;color:#c62828;'>{int(r.Open):,}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:700;color:#2e7d32;'>{int(r.Closed):,}</td>"
            f"<td style='padding:6px 10px;text-align:center;'>{_comp_badge(r._comp)}</td>"
            f"<td style='padding:6px 12px;min-width:90px;'>{_progress_bar(r._comp)}</td>"
            f"</tr>"
        )
    _zrank_html = (
        "<div style='overflow-x:auto;margin:6px 0 16px 0;'>"
        "<table style='border-collapse:collapse;width:100%;font-size:12px;'>"
        f"<thead><tr style='background:#003087;'>"
        f"<th style='{_th}text-align:center;color:#fff;'>Rank</th>"
        f"<th style='{_th}text-align:left;color:#fff;'>Zone</th>"
        f"<th style='{_th}text-align:center;color:#fff;'>Locations</th>"
        f"<th style='{_th}text-align:center;color:#ffb3b3;'>Open</th>"
        f"<th style='{_th}text-align:center;color:#b3ffb3;'>Closed</th>"
        f"<th style='{_th}text-align:center;color:#fff;'>Compliance %</th>"
        f"<th style='{_th}text-align:left;color:#fff;'>Progress</th>"
        f"</tr></thead><tbody>{_zrank_rows}</tbody></table></div>"
    )
    st.markdown(_zrank_html, unsafe_allow_html=True)

    # ── Location-wise Compliance Ranking ──────────────────────────────────────
    st.markdown("<a id='lv-rank-loc'></a>", unsafe_allow_html=True)
    st.markdown("<div class='sec-title'>&#127941; Location-wise Compliance Ranking</div>", unsafe_allow_html=True)
    _lr = loc_chart.copy().sort_values("_pct", ascending=False).reset_index(drop=True)
    _lrank_rows = ""
    for i, r in _lr.iterrows():
        rank = i + 1
        _bg = "#fffde7" if rank == 1 else ("#f3f3f3" if rank == 2 else ("#fdf3ee" if rank == 3 else ("#ffffff" if rank % 2 == 1 else "#f7fafd")))
        _lrank_rows += (
            f"<tr style='background:{_bg};'>"
            f"<td style='padding:6px 10px;text-align:center;'>{_rank_badge(rank)}</td>"
            f"<td style='padding:6px 10px;font-weight:600;'>{html.escape(str(r['Plant Desc.']))}</td>"
            f"<td style='padding:6px 10px;text-align:center;'>{_zone_badge(r.Zone)}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:600;'>{int(r.Total):,}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:700;color:#c62828;'>{int(r.Open):,}</td>"
            f"<td style='padding:6px 10px;text-align:center;font-weight:700;color:#2e7d32;'>{int(r.Closed):,}</td>"
            f"<td style='padding:6px 10px;text-align:center;'>{_comp_badge(r._pct)}</td>"
            f"</tr>"
        )
    _lrank_html = (
        "<div style='overflow-x:auto;margin:6px 0 16px 0;max-height:500px;overflow-y:auto;'>"
        "<table style='border-collapse:collapse;width:100%;font-size:12px;'>"
        f"<thead><tr style='background:#003087;position:sticky;top:0;z-index:1;'>"
        f"<th style='{_th}text-align:center;color:#fff;'>Rank</th>"
        f"<th style='{_th}text-align:left;color:#fff;'>Location</th>"
        f"<th style='{_th}text-align:center;color:#fff;'>Zone</th>"
        f"<th style='{_th}text-align:center;color:#fff;'>Total</th>"
        f"<th style='{_th}text-align:center;color:#ffb3b3;'>Open</th>"
        f"<th style='{_th}text-align:center;color:#b3ffb3;'>Closed</th>"
        f"<th style='{_th}text-align:center;color:#fff;'>Compliance %</th>"
        f"</tr></thead><tbody>{_lrank_rows}</tbody></table></div>"
    )
    st.markdown(_lrank_html, unsafe_allow_html=True)

    # ── Missing Locations ─────────────────────────────────────────────────────
    st.markdown("<a id='lv-missing'></a>", unsafe_allow_html=True)
    st.markdown("<div class='sec-title'>&#9888; Locations Not Yet Audited (in scope)</div>", unsafe_allow_html=True)
    _pm = load_plant_master()[["Plant Code","Plant Name","Zone Name"]].copy()
    _pm["Plant Code"] = _pm["Plant Code"].astype(str).str.strip().str.replace(r"\.0$","",regex=True)
    _src_codes = set(df["Planning Plant"].astype(str).str.strip().str.replace(r"\.0$","",regex=True).unique())
    if sel_zone != "All":
        _pm = _pm[_pm["Zone Name"].astype(str) == sel_zone]
    _missing = (
        _pm[~_pm["Plant Code"].isin(_src_codes)]
        .rename(columns={"Zone Name":"Zone","Plant Code":"Planning Plant","Plant Name":"Location"})
        [["Zone","Planning Plant","Location"]]
        .sort_values(["Zone","Planning Plant"])
        .reset_index(drop=True)
    )
    if _missing.empty:
        st.success("All locations in scope have at least one audit on record.")
    else:
        _render_html_table(_missing, max_height=320)

    st.markdown("---")

    # ── Download ──────────────────────────────────────────────────────────────
    _download_excel_button(
        label="⬇  Download Drill-Down Data  (.xlsx)",
        file_prefix="location_visit_drilldown",
        sheets={
            "Audit Detail":     audit_tbl.drop(columns=["S.No"], errors="ignore").reset_index(drop=True),
            "Zone Summary":     zone_tbl.drop(columns=["_comp"], errors="ignore").reset_index(drop=True),
            "Quarter Summary":  qtr_tbl.reset_index(drop=True),
        },
        key="dl_loc_visit_drill",
    )

import base64
import pandas as pd
import html
import os
from io import BytesIO
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION — File Paths & Color Palette
# ─────────────────────────────────────────────────────────────────────────────

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
MASTER_DIR = os.path.join(BASE_DIR, "MAster")

# Master file paths
PLANT_MASTER_PATH = os.path.join(MASTER_DIR, "PlantMaster.xlsx")
ZONE_MASTER_PATH  = os.path.join(MASTER_DIR, "Zonewise MaiID Master.xlsx")

# Brand image paths (Title banner + logo)
TITLE_IMG_PATH = os.path.join(MASTER_DIR, "Title.png")
LOGO_IMG_PATH  = os.path.join(MASTER_DIR, "Master Logo.jpg")
SIDE_PANEL_LOGO_PATH = os.path.join(MASTER_DIR, "Side Panel Logo.png")

# Default data file paths (fallback when no file is uploaded)
REPORTS_DIR         = os.path.join(BASE_DIR, "Reports")
PENDING_DC_PATH     = os.path.join(REPORTS_DIR, "PENDING_DC_SOD.xlsx")
OPEN_DELIVERY_PATH  = os.path.join(REPORTS_DIR, "OPEN_DELIVERY.xls")
OPEN_INTRANSIT_PATH = os.path.join(REPORTS_DIR, "OPEN_INTRANSIT_SOD.xls")
OPEN_SO_PATH        = os.path.join(REPORTS_DIR, "OPEN_SALES_ORDER.xls")
PEND_INV_PATH       = os.path.join(REPORTS_DIR, "PENDING_INVOICES_SOD.xls")
SHORT_SALES_PATH    = os.path.join(REPORTS_DIR, "SOD_OPEN_SHORTAGES_SALES.xls")
SHORT_STO_PATH      = os.path.join(REPORTS_DIR, "SOD_OPEN_SHORTAGES_STO.xls")
TANK_RECO_PATH      = os.path.join(REPORTS_DIR, "TANK_RECO_REPORT.xls")
LOCAL_LOCATION_VISIT_PATH = os.path.join(REPORTS_DIR, "LOCATION_VISIT.xls")
EXTERNAL_LOCATION_VISIT_PATH = r"D:\SHOAIB\VS CODE PROJECTS\EXCEPTION SNAPSHOT DASHBOARD\Reports\LOCATION_VISIT.xls"
LOCATION_VISIT_PATH = (
    LOCAL_LOCATION_VISIT_PATH
    if os.path.exists(LOCAL_LOCATION_VISIT_PATH)
    else EXTERNAL_LOCATION_VISIT_PATH
    if os.path.exists(EXTERNAL_LOCATION_VISIT_PATH)
    else LOCAL_LOCATION_VISIT_PATH
)
DUMMY_TANK_PATH     = os.path.join(REPORTS_DIR, "DUMMY TANK STOCK.xls")
PIPELINE_STOCK_PATH = os.path.join(REPORTS_DIR, "PIPELINE STOCK.xls")       
TANK_TURNS_PATH     = os.path.join(REPORTS_DIR, "Tank Turn.xlsx")
# HPCL Corporate Color Palette
C = {
    "primary"    : "#003087",
    "secondary"  : "#0057A8",
    "accent"     : "#FF6600",
    "light_blue" : "#E8F0FE",
    "white"      : "#FFFFFF",
    "bg"         : "#F4F6FA",
    "text_muted" : "#6C757D",
    "border"     : "#D0DDEF",
    "success"    : "#28A745",
    "warning"    : "#E6A817",
    "danger"     : "#C82333",
    "shadow"     : "rgba(0,48,135,0.12)",
}

# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT PAGE CONFIG  (must be the very first Streamlit call)
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SOD Exception Dashboard",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Force Streamlit to use port 8502
import sys
if hasattr(sys, 'argv'):
    sys.argv += ["--server.port=8502"]

# ─────────────────────────────────────────────────────────────────────────────
# IMAGE LOADER  (cached — reads brand images once per session)
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _load_img_b64(path: str) -> str:
    """Return a base64 data-URI for embedding an image in HTML. Cached."""
    ext  = os.path.splitext(path)[1].lower().lstrip(".")
    mime = "jpeg" if ext == "jpg" else ext
    with open(path, "rb") as fh:
        b64 = base64.b64encode(fh.read()).decode()
    return f"data:image/{mime};base64,{b64}"


# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL CSS INJECTION
# ─────────────────────────────────────────────────────────────────────────────

def inject_css() -> None:
    """Inject all custom CSS for the HPCL corporate theme."""
    st.markdown(f"""
    <style>
    /* ══ FORCE LIGHT THEME ════════════════════════════════════════════ */
    /* 1. Pseudo-element white canvas — sits below all content, cannot be
          overridden by Streamlit's React theme engine */
    body::before {{
        content: '';
        position: fixed;
        inset: 0;
        background-color: {C['bg']};
        z-index: -9999;
        pointer-events: none;
    }}
    /* 2. Tell the browser this page is light — defeats prefers-color-scheme:dark */
    html {{
        color-scheme: light only !important;
        background-color: {C['bg']} !important;
    }}
    body {{
        background-color: {C['bg']} !important;
    }}
    /* 2. Override Streamlit's CSS custom properties in every possible scope */
    :root,
    :root[data-theme],
    [data-theme="dark"],
    [data-theme="light"] {{
        --background-color:                {C['bg']}  !important;
        --secondary-background-color:      #F0F2F6    !important;
        --text-color:                      #262730    !important;
        color-scheme: light only !important;
    }}
    /* 3. Neutralise dark-mode media query Streamlit ships */
    @media (prefers-color-scheme: dark) {{
        html, body {{
            background-color: {C['bg']} !important;
            color: #262730 !important;
            color-scheme: light only !important;
        }}
        [data-testid="stApp"],
        [data-testid="stAppViewContainer"],
        [data-testid="stAppViewBlockContainer"],
        [data-testid="stMain"],
        [data-testid="stMainBlockContainer"],
        [data-testid="block-container"],
        #root, .stApp, .main, .block-container, section.main {{
            background-color: {C['bg']} !important;
            color: #262730 !important;
        }}
        [data-baseweb="select"] > div,
        [data-baseweb="input"]  > div,
        [data-baseweb="textarea"],
        input, textarea, select {{
            background-color: #FFFFFF !important;
            color: #262730 !important;
        }}
    }}
    /* 4. Always-on rules for all Streamlit containers (v1.20 – v1.45+) */
    html, body, #root,
    .stApp, div.stApp,
    [data-testid="stApp"],
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewBlockContainer"],
    [data-testid="stMain"],
    [data-testid="stMainBlockContainer"],
    [data-testid="block-container"],
    section.main, .main, .block-container {{
        background-color: {C['bg']} !important;
        color: #262730 !important;
    }}
    /* 5. Inputs / dropdowns — always white */
    [data-baseweb="select"] > div,
    [data-baseweb="input"]  > div,
    [data-baseweb="textarea"],
    [data-testid="stSelectbox"] div[class*="container"],
    .stTextInput input, .stNumberInput input,
    .stSelectbox div[data-baseweb="select"] > div {{
        background-color: #FFFFFF !important;
        color: #262730 !important;
    }}
    [data-testid="metric-container"] {{
        background-color: #FFFFFF !important;
    }}

    /* ── Base ─────────────────────────────────────────── */
    html, body, [class*="css"] {{
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 16px;
    }}
    :root {{
        --hpcl-main-top-shift: 0rem;
        --hpcl-sidebar-top-shift: 0rem;
    }}
    [data-testid="stHeader"] {{
        display: none !important;
    }}
    [data-testid="stDecoration"] {{
        display: none !important;
    }}
    [data-testid="stToolbar"] {{
        display: none !important;
    }}
    [data-testid="stSidebarCollapsedControl"] {{
        display: none !important;
    }}
    [data-testid="stAppViewContainer"] > .main,
    [data-testid="stAppViewContainer"] > .main > div,
    [data-testid="stMain"],
    [data-testid="stMainBlockContainer"],
    .main-container,
    .page-container,
    .content-wrapper,
    .container,
    .container-fluid {{
        margin-top: 0 !important;
        padding-top: 0 !important;
    }}
    .main .block-container {{
        padding-top: 0 !important;
        padding-bottom: 1rem;
        padding-left: 0.75rem !important;
        padding-right: 0.75rem !important;
        max-width: 100%;
    }}
    .main .block-container > div:first-child {{
        margin-top: 0 !important;
        padding-top: 0 !important;
    }}

    /* ── Full-Width Title Banner (≈ 2 inches / 192 px tall) ── */
    .dashboard-header-shell {{
        margin-top: calc(-1 * var(--hpcl-main-top-shift)) !important;
        padding-top: 0 !important;
        position: relative;
        z-index: 2;
    }}
    .hpcl-banner-wrap {{
        margin: 0 -1rem !important;
        width: calc(100% + 2rem);
        height: 66px;
        line-height: 0;
        overflow: hidden;
        position: relative;
        background: linear-gradient(135deg, #003087 0%, #0057A8 100%);
    }}
    .hpcl-banner-fg {{
        position: relative;
        z-index: 1;
        width: 100%;
        height: 100%;
        object-fit: contain;
        object-position: center;
        display: block;
        padding: 0 10px 2px 10px;
        filter: drop-shadow(0 0 1px rgba(0, 48, 135, 0.9))
                drop-shadow(0 0 2px rgba(0, 48, 135, 0.65));
    }}
    .banner-date {{
        position: absolute;
        bottom: 5px;
        right: 16px;
        color: rgba(255,255,255,0.92);
        font-size: 11px;
        font-weight: 700;
        background: rgba(0,0,0,0.32);
        padding: 2px 9px;
        border-radius: 4px;
        z-index: 5;
        letter-spacing: 0.3px;
    }}

    /* ── Info Strip below banner ──────────────────────── */
    .dash-header {{
        background: linear-gradient(135deg, {C['primary']} 0%, {C['secondary']} 100%);
        color: white;
        padding: 8px 20px;
        margin: 0 -1rem 14px -1rem;
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        box-shadow: 0 3px 10px rgba(0,48,135,0.28);
    }}
    .dash-header-main {{
        width: 100%;
        text-align: center;
    }}
    .dash-header-title {{
        font-size: 22px !important;
        font-weight: 800;
        letter-spacing: 0.02em;
        margin: 0;
        line-height: 1.2;
    }}
    .dash-header-sub {{
        font-size: 13px;
        opacity: 0.88;
        margin: 3px 0 0 0;
    }}

    /* ── KPI Cards ─────────────────────────────────────── */
    .kpi-wrap {{
        background: {C['white']};
        border-radius: 10px;
        padding: 11px 14px 10px 14px;
        border-left: 4px solid {C['primary']};
        box-shadow: 0 2px 8px {C['shadow']};
        position: relative;
        overflow: hidden;
        transition: transform 0.18s ease, box-shadow 0.18s ease;
        min-height: 90px;
    }}
    .kpi-wrap:hover {{
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0,48,135,0.18);
    }}
    .kpi-wrap::after {{
        content: '';
        position: absolute;
        top: 0; right: 0;
        width: 44px; height: 44px;
        background: {C['light_blue']};
        border-radius: 0 10px 0 44px;
    }}
    .kpi-icon {{
        position: absolute;
        top: 10px; right: 12px;
        font-size: 1.2rem;
        opacity: 0.55;
        z-index: 1;
    }}
    .kpi-label {{
        font-size: 11px;
        font-weight: 800;
        color: #46515F;
        text-transform: uppercase;
        letter-spacing: 0.07em;
        margin-bottom: 5px;
        line-height: 1.25;
    }}
    .kpi-value {{
        font-size: 1.55rem;
        font-weight: 800;
        color: {C['primary']};
        line-height: 1;
        margin-bottom: 4px;
    }}
    .kpi-detail {{
        font-size: 11px;
        font-weight: 600;
        color: #24405A;
        line-height: 1.3;
    }}
    .kpi-wrap.c-danger  {{ border-left-color: {C['danger']};   }}
    .kpi-wrap.c-warning {{ border-left-color: {C['warning']};  }}
    .kpi-wrap.c-success {{ border-left-color: {C['success']};  }}
    .kpi-wrap.c-orange  {{ border-left-color: {C['accent']};   }}
    .kpi-wrap.c-muted   {{ border-left-color: #AAAAAA; }}
    .kpi-wrap.c-muted .kpi-label  {{ color: #46515F !important; font-weight: 800 !important; opacity: 1 !important; }}
    .kpi-wrap.c-muted .kpi-value  {{ opacity: 0.55; color: #8A96A8; }}
    .kpi-wrap.c-muted .kpi-icon   {{ opacity: 0.30; }}
    .kpi-wrap.c-muted .kpi-detail {{ opacity: 0.75; color: #6B7A8D; }}

    /* ── Section Titles ──────────────────────────────────── */
    .sec-title {{
        font-size: 15px;
        font-weight: 700;
        color: {C['primary']};
        padding: 5px 0 4px 0;
        border-bottom: 2px solid {C['light_blue']};
        margin: 8px 0 10px 0;
    }}
    .pro-table-wrap {{
        max-height: 520px;
        overflow: auto;
        border: 1px solid {C['border']};
        border-radius: 12px;
        background: {C['white']};
        box-shadow: 0 3px 14px rgba(0,48,135,0.08);
    }}
    .pro-table {{
        width: 100%;
        border-collapse: collapse;
        table-layout: auto;
        min-width: 480px;
    }}
    .pro-table thead th {{
        position: sticky;
        top: 0;
        z-index: 1;
        background: linear-gradient(135deg, {C['primary']} 0%, {C['secondary']} 100%);
        color: #FFFFFF;
        font-size: 15px;
        font-weight: 700;
        text-align: center;
        padding: 13px 16px;
        border-bottom: 3px solid #FFD700;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        white-space: nowrap;
    }}
    .pro-table thead th:first-child {{
        border-radius: 0;
        text-align: left;
    }}
    .pro-table tbody td {{
        font-size: 14px;
        font-weight: 500;
        color: #1B3552;
        text-align: center;
        padding: 11px 16px;
        border-bottom: 1px solid #E2EAF4;
        word-wrap: break-word;
    }}
    .pro-table tbody td:first-child {{
        text-align: left;
        font-weight: 600;
        color: #003087;
    }}
    .pro-table tbody tr:nth-child(odd) {{
        background: #FFFFFF;
    }}
    .pro-table tbody tr:nth-child(even) {{
        background: #F4F8FF;
    }}
    .pro-table tbody tr:hover {{
        background: #DCF0FF;
        transition: background 0.15s ease;
    }}
    .streamlit-expanderHeader {{
        font-size: 20px !important;
        font-weight: 700 !important;
        color: {C['primary']} !important;
    }}

    /* ── Detail Header ──────────────────────────────────── */
    .detail-hdr {{
        background: {C['light_blue']};
        border-left: 6px solid {C['primary']};
        padding: 16px 22px;
        border-radius: 8px;
        margin-bottom: 20px;
    }}
    .detail-hdr h3 {{
        color: {C['primary']};
        margin: 0;
        font-size: 24px;
        font-weight: 700;
    }}
    .detail-hdr p {{
        margin: 7px 0 0;
        font-size: 20px;
        color: {C['text_muted']};
    }}

    /* ── Sidebar ────────────────────────────────────────── */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {C['primary']} 0%, #001A5C 100%);
        display: block !important;
        visibility: visible !important;
        transform: none !important;
        width: 16rem !important;
        min-width: 16rem !important;
    }}
    [data-testid="stSidebar"][aria-expanded="false"] {{
        transform: translateX(0) !important;
        display: block !important;
    }}
    [data-testid="stSidebarContent"],
    [data-testid="stSidebarUserContent"] {{
        padding-top: 0 !important;
        margin-top: 0 !important;
    }}
    [data-testid="stSidebar"] > div:first-child,
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:first-child,
    .sidebar,
    .sidebar-header,
    .logo-container {{
        margin-top: 0 !important;
        padding-top: 0 !important;
    }}
    .sidebar-branding {{
        text-align: center;
        padding: 0 0 4px 0 !important;
        margin-top: calc(-1 * var(--hpcl-sidebar-top-shift)) !important;
        position: relative;
        z-index: 2;
    }}
    .sidebar-branding img {{
        margin-top: 0 !important;
    }}
    @media (max-width: 700px) {{
        :root {{
            --hpcl-main-top-shift: 0rem;
            --hpcl-sidebar-top-shift: 0rem;
        }}
    }}
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div,
    [data-testid="stSidebar"] label {{
        color: #DDEAFF !important;
        font-size: 12px !important;
    }}
    [data-testid="stSidebar"] .stMultiSelect label {{
        color: #FFFFFF !important;
        font-size: 11px !important;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.07em;
    }}
    /* ── File uploader drop zone ────────────────────────── */
    [data-testid="stSidebar"] [data-testid="stFileUploadDropzone"] {{
        background: rgba(255,255,255,0.95) !important;
        border: 2px dashed #5B9BD5 !important;
        border-radius: 8px !important;
    }}
    [data-testid="stSidebar"] .stFileUploader span,
    [data-testid="stSidebar"] .stFileUploader div,
    [data-testid="stSidebar"] .stFileUploader p,
    [data-testid="stSidebar"] .stFileUploader small,
    [data-testid="stSidebar"] .stFileUploader section span,
    [data-testid="stSidebar"] .stFileUploader section div {{
        color: #111111 !important;
        font-size: 13px !important;
        font-weight: 600 !important;
    }}
    [data-testid="stSidebar"] .stFileUploader label {{
        color: #FFFFFF !important;
        font-size: 14px !important;
        font-weight: 800 !important;
        text-transform: uppercase;
        letter-spacing: 0.07em;
    }}
    [data-testid="stSidebar"] .stFileUploader button {{
        background: #1B3552 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        font-size: 13px !important;
    }}
    [data-testid="stSidebar"] hr {{
        border-color: rgba(255,255,255,0.18) !important;
    }}
    .sb-nav-lbl {{
        font-size: 10px !important;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: #7AABF0 !important;
        margin: 10px 0 4px 0;
    }}
    .sb-critical-box {{
        background: linear-gradient(180deg, rgba(255,255,255,0.12), rgba(255,255,255,0.05));
        border: 1px solid rgba(122,171,240,0.35);
        border-radius: 10px;
        padding: 10px 12px;
        margin: 0 0 10px 0;
    }}
    .sb-critical-title {{
        color: #FFFFFF;
        font-size: 14px;
        font-weight: 700;
        margin: 0 0 4px 0;
    }}
    .sb-critical-subtitle {{
        color: rgba(255,255,255,0.78);
        font-size: 12px;
        line-height: 1.4;
        margin: 0;
    }}

    /* ── Sidebar multiselect — glass-card style ─────────── */
    [data-testid="stSidebar"] [data-baseweb="select"] {{
        background: rgba(255,255,255,0.10) !important;
        border: 1px solid rgba(255,255,255,0.20) !important;
        border-radius: 10px !important;
        transition: border-color 0.2s ease, background 0.2s ease !important;
    }}
    [data-testid="stSidebar"] [data-baseweb="select"]:hover {{
        background: rgba(255,255,255,0.16) !important;
        border-color: rgba(255,255,255,0.38) !important;
    }}
    [data-testid="stSidebar"] [data-baseweb="select"] > div {{
        background: transparent !important;
        color: #FFFFFF !important;
        font-size: 13px !important;
        border: none !important;
    }}
    [data-testid="stSidebar"] [data-baseweb="tag"] {{
        background: rgba(255,200,50,0.22) !important;
        border: 1px solid rgba(255,200,50,0.45) !important;
        border-radius: 20px !important;
        color: #FFE066 !important;
        font-size: 12px !important;
    }}
    /* Sidebar buttons — glass-card nav style */
    [data-testid="stSidebar"] [data-testid="stButton"] > button {{
        width: calc(100% - 8px) !important;
        display: flex !important;
        align-items: center !important;
        justify-content: flex-start !important;
        background: rgba(255,255,255,0.07) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
        border-radius: 10px !important;
        margin: 3px 4px !important;
        padding: 10px 14px !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        letter-spacing: 0.02em !important;
        box-shadow: none !important;
        transition: background 0.18s ease, transform 0.12s ease, border-color 0.18s ease, box-shadow 0.18s ease !important;
    }}
    [data-testid="stSidebar"] [data-testid="stButton"] > button:hover {{
        background: rgba(255,255,255,0.16) !important;
        border-color: rgba(255,255,255,0.28) !important;
        transform: translateX(3px) !important;
        box-shadow: 0 3px 10px rgba(0,0,0,0.22) !important;
        color: #FFFFFF !important;
    }}
    [data-testid="stSidebar"] [data-testid="stButton"] > button:active {{
        transform: translateX(1px) !important;
        box-shadow: 0 1px 4px rgba(0,0,0,0.18) !important;
    }}

    /* ── Filter Badges ──────────────────────────────────── */
    .fbadge {{
        display: inline-block;
        background: {C['light_blue']};
        color: {C['primary']};
        border-radius: 20px;
        padding: 4px 14px;
        font-size: 15px;
        font-weight: 600;
        margin: 2px 3px;
    }}

    /* ── Buttons ────────────────────────────────────────── */
    div.stButton > button {{
        background: {C['primary']};
        color: white;
        border: none;
        border-radius: 6px;
        padding: 5px 14px;
        font-weight: 600;
        font-size: 12px;
        transition: background 0.2s;
    }}
    div.stButton > button:hover {{
        background: {C['secondary']};
        color: white;
        border: none;
    }}
    div[data-testid="stDownloadButton"] > button {{
        background: {C['accent']};
        color: white;
        border: none;
        border-radius: 6px;
        font-weight: 600;
        font-size: 12px;
    }}
    div[data-testid="stDownloadButton"] > button:hover {{
        background: #E05500;
        color: white;
    }}

    /* ── Tab overrides ──────────────────────────────────── */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 4px;
        background: {C['light_blue']};
        border-radius: 8px;
        padding: 4px;
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 6px;
        font-size: 15px;
        font-weight: 600;
        padding: 6px 16px;
    }}
    .stTabs [aria-selected="true"] {{
        background: {C['primary']} !important;
        color: white !important;
    }}

    /* ── Streamlit native st.metric labels ──────────────── */
    [data-testid="stMetricLabel"] {{
        font-size: 20px !important;
        font-weight: 600 !important;
        color: {C['text_muted']} !important;
    }}
    [data-testid="stMetricValue"] {{
        font-size: 32px !important;
        font-weight: 800 !important;
        color: {C['primary']} !important;
    }}
    [data-testid="stMetricDelta"] {{
        font-size: 15px !important;
    }}

    /* ── Streamlit dataframe tables (drill-down pages) ── */
    [data-testid="stDataFrame"] [role="columnheader"] {{
        background: linear-gradient(135deg, {C['primary']} 0%, {C['secondary']} 100%) !important;
        color: #FFFFFF !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.04em !important;
        border-bottom: 2px solid #FFD700 !important;
    }}
    [data-testid="stDataFrame"] [role="gridcell"] {{
        font-size: 13px !important;
        font-weight: 500 !important;
        color: #1B3552 !important;
    }}
    [data-testid="stDataFrame"] {{
        border-radius: 10px !important;
        overflow: hidden !important;
        border: 1px solid #D5E2F3 !important;
    }}
    </style>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING  (cached where possible)
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_plant_master() -> pd.DataFrame:
    """
    Load PlantMaster.xlsx from disk (cached).
    Only returns rows where Active == 'Yes'.
    Optimized: Only load required columns.
    """
    df = pd.read_excel(PLANT_MASTER_PATH, dtype={"Plant Code": str}, engine="openpyxl")
    df.columns = df.columns.astype(str).str.replace("\n", " ", regex=False).str.strip()

    normalized = {
        " ".join(col.lower().split()): col
        for col in df.columns
    }

    zone_aliases = [
        "zone name",
        "new name of zone as per cfd minutes",
        "new name of zone",
        "zone",
    ]
    zone_col = next((normalized[key] for key in zone_aliases if key in normalized), None)

    required_cols = {
        "plant code": "Plant Code",
        "plant name": "Plant Name",
    }
    rename_map = {}
    missing_required = []
    for key, target in required_cols.items():
        source = normalized.get(key)
        if source is None:
            missing_required.append(target)
        else:
            rename_map[source] = target

    if zone_col is None:
        missing_required.append("Zone Name")
    else:
        rename_map[zone_col] = "Zone Name"

    if missing_required:
        raise ValueError(
            "PlantMaster missing required column(s): " + ", ".join(missing_required)
        )

    if "active" in normalized:
        rename_map[normalized["active"]] = "Active"

    df = df.rename(columns=rename_map)
    keep_cols = ["Plant Code", "Plant Name", "Zone Name"] + (["Active"] if "Active" in df.columns else [])
    df = df[keep_cols].copy()

    df["Plant Code"] = df["Plant Code"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df["Plant Name"] = df["Plant Name"].astype(str).str.strip()
    df["Zone Name"] = df["Zone Name"].astype(str).str.strip()
    df["Zone Name"] = df["Zone Name"].replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(subset=["Zone Name"])

    if "Active" in df.columns:
        df = df[df["Active"].astype(str).str.strip().str.lower() == "yes"]
    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_zone_master() -> pd.DataFrame:
    """Load Zonewise MaiID Master.xlsx from disk (cached). Optimized: Only load required columns."""
    usecols = ["Zone Name"]
    df = pd.read_excel(ZONE_MASTER_PATH, engine="openpyxl", usecols=usecols)
    df.columns = df.columns.str.strip()
    df["Zone Name"] = df["Zone Name"].astype(str).str.strip()
    return df.reset_index(drop=True)


def _get_file_cache_token(path: str) -> tuple[int, int] | None:
    """Return a lightweight file signature so Streamlit cache refreshes after file updates."""
    try:
        stat = os.stat(path)
        return stat.st_mtime_ns, stat.st_size
    except OSError:
        return None


@st.cache_data(show_spinner=False)
def _load_excel_from_path(path: str, cache_buster: tuple[int, int] | None = None) -> pd.DataFrame:
    """Internal helper: load any Excel file from a disk path (cached)."""
    del cache_buster  # used only to invalidate Streamlit cache when file changes
    ext = os.path.splitext(path)[1].lower()
    df = _read_excel_flexible(path, ext_hint=ext)
    df.columns = df.columns.str.strip().str.upper()
    return df


def _read_excel_flexible(source, ext_hint: str = "") -> pd.DataFrame:
    """Read Excel with engine fallback for mislabeled .xls/.xlsx files."""
    preferred_engine = "xlrd" if ext_hint.lower() == ".xls" else "openpyxl"
    try:
        return pd.read_excel(source, engine=preferred_engine)
    except Exception as exc:
        msg = str(exc).lower()
        fallback_engine = None

        if preferred_engine == "xlrd" and (
            "xlsx file; not supported" in msg
            or "zip" in msg
        ):
            fallback_engine = "openpyxl"
        elif preferred_engine == "openpyxl" and (
            "old .xls" in msg
            or "not a zip file" in msg
            or "file format cannot be determined" in msg
        ):
            fallback_engine = "xlrd"

        if fallback_engine is None:
            raise

        if hasattr(source, "seek"):
            try:
                source.seek(0)
            except Exception:
                pass

        return pd.read_excel(source, engine=fallback_engine)


def load_pending_dc(source) -> pd.DataFrame:
    """
    Load the Pending DC data file.

    source: str path (cached disk load) OR UploadedFile (live, not cached).
    Returns DataFrame with UPPER-stripped column names.
    """
    try:
        if isinstance(source, str):
            return _load_excel_from_path(source, cache_buster=_get_file_cache_token(source))
        name   = getattr(source, "name", "file.xlsx")
        ext    = os.path.splitext(name)[1].lower()
        df = _read_excel_flexible(source, ext_hint=ext)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading Pending DC file: {exc}")
        return pd.DataFrame()


def load_open_delivery(source) -> pd.DataFrame:
    """
    Load the Open Delivery data file.

    source: str path (cached disk load) OR UploadedFile (live, not cached).
    Returns DataFrame with UPPER-stripped column names.
    """
    try:
        if isinstance(source, str):
            return _load_excel_from_path(source, cache_buster=_get_file_cache_token(source))
        name   = getattr(source, "name", "file.xlsx")
        ext    = os.path.splitext(name)[1].lower()
        df = _read_excel_flexible(source, ext_hint=ext)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading Open Delivery file: {exc}")
        return pd.DataFrame()


def load_open_intransit(source) -> pd.DataFrame:
    """
    Load the Open In-Transit data file.

    source: str path (cached disk load) OR UploadedFile (live, not cached).
    Returns DataFrame with UPPER-stripped column names.
    """
    try:
        if isinstance(source, str):
            return _load_excel_from_path(source, cache_buster=_get_file_cache_token(source))
        name   = getattr(source, "name", "file.xlsx")
        ext    = os.path.splitext(name)[1].lower()
        df = _read_excel_flexible(source, ext_hint=ext)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading Open In-Transit file: {exc}")
        return pd.DataFrame()


def load_open_sales_orders(source) -> pd.DataFrame:
    """
    Load the Open Sales Orders data file.

    source: str path (cached disk load) OR UploadedFile (live, not cached).
    Returns DataFrame with UPPER-stripped column names.
    """
    try:
        if isinstance(source, str):
            return _load_excel_from_path(source, cache_buster=_get_file_cache_token(source))
        name   = getattr(source, "name", "file.xlsx")
        ext    = os.path.splitext(name)[1].lower()
        df = _read_excel_flexible(source, ext_hint=ext)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading Open Sales Orders file: {exc}")
        return pd.DataFrame()


def load_pending_invoices(source) -> pd.DataFrame:
    """
    Load the Pending Invoices data file.

    source: str path (cached disk load) OR UploadedFile (live, not cached).
    Returns DataFrame with UPPER-stripped column names.
    """
    try:
        if isinstance(source, str):
            return _load_excel_from_path(source, cache_buster=_get_file_cache_token(source))
        name   = getattr(source, "name", "file.xlsx")
        ext    = os.path.splitext(name)[1].lower()
        df = _read_excel_flexible(source, ext_hint=ext)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading Pending Invoices file: {exc}")
        return pd.DataFrame()


def load_tank_reco(source) -> pd.DataFrame:
    """
    Load the Abnormal Variations in SAP data file.

    source: str path (cached disk load) OR UploadedFile (live, not cached).
    Returns DataFrame with UPPER-stripped column names.
    """
    try:
        if isinstance(source, str):
            return _load_excel_from_path(source, cache_buster=_get_file_cache_token(source))
        name   = getattr(source, "name", "file.xlsx")
        ext    = os.path.splitext(name)[1].lower()
        df = _read_excel_flexible(source, ext_hint=ext)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading Abnormal Variations in SAP file: {exc}")
        return pd.DataFrame()


def load_open_shortages_sales(source) -> pd.DataFrame:
    """
    Load the OPEN SHORTAGES - Ltrs (Sales) data file.

    source: str path (cached disk load) OR UploadedFile (live, not cached).
    Returns DataFrame with UPPER-stripped column names.
    """
    try:
        if isinstance(source, str):
            return _load_excel_from_path(source, cache_buster=_get_file_cache_token(source))
        name   = getattr(source, "name", "file.xlsx")
        ext    = os.path.splitext(name)[1].lower()
        df = _read_excel_flexible(source, ext_hint=ext)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading OPEN SHORTAGES - Ltrs (Sales) file: {exc}")
        return pd.DataFrame()


def load_open_shortages_sto(source) -> pd.DataFrame:
    """
    Load the OPEN SHORTAGES - Ltrs (STO) data file.

    source: str path (cached disk load) OR UploadedFile (live, not cached).
    Returns DataFrame with UPPER-stripped column names.
    """
    try:
        if isinstance(source, str):
            return _load_excel_from_path(source, cache_buster=_get_file_cache_token(source))
        name   = getattr(source, "name", "file.xlsx")
        ext    = os.path.splitext(name)[1].lower()
        df = _read_excel_flexible(source, ext_hint=ext)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading OPEN SHORTAGES - Ltrs (STO) file: {exc}")
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_dummy_tank_stock(path: str, cache_buster: tuple[int, int] | None = None) -> pd.DataFrame:
    """Load Dummy Tank Stock report from disk and normalize critical columns."""
    try:
        del cache_buster  # used only to invalidate Streamlit cache when file changes
        ext = os.path.splitext(path)[1].lower()
        df = _read_excel_flexible(path, ext_hint=ext)
        df.columns = df.columns.astype(str).str.strip()

        # Normalize known report headers that can vary by case/spelling.
        canonical_cols = {
            "plant": "Plant",
            "material": "Material",
            "storage location": "Storage Location",
            "base unit of measure": "Base Unit of Measure",
            "unrestricted": "Unrestricted",
            "zone": "Zone",
        }
        rename_map = {}
        for col in df.columns:
            key = " ".join(str(col).strip().lower().split())
            if key in canonical_cols:
                rename_map[col] = canonical_cols[key]
        if rename_map:
            df = df.rename(columns=rename_map)

        text_cols = ["Plant", "Material", "Storage Location", "Base Unit of Measure", "Zone"]
        for col in text_cols:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()

        if "Unrestricted" in df.columns:
            df["Unrestricted"] = pd.to_numeric(df["Unrestricted"], errors="coerce").fillna(0)

        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_tank_turns(path: str, cache_buster: tuple[int, int] | None = None) -> pd.DataFrame:
    """Load Tank Turn report from disk and normalize critical columns."""
    try:
        del cache_buster  # used only to invalidate Streamlit cache when file changes
        ext = os.path.splitext(path)[1].lower()
        df = _read_excel_flexible(path, ext_hint=ext)
        df.columns = df.columns.astype(str).str.strip()

        canonical_cols = {
            "zone": "Zone",
            "plant": "Plant",
            "plant name": "Plant Name",
            "tank": "Tank",
            "unique ref id": "Unique Ref Id",
            "material": "Material",
            "material description": "Material Description",
            "tank capacity": "Tank Capacity",
            "dispatches": "Dispatches",
            "turn": "Turn",
            "tank type": "Tank Type",
            "tank status": "Tank Status",
            "opening stock": "Opening Stock",
            "receipts": "Receipts",
            "closing stock": "Closing Stock",
        }
        rename_map = {}
        for col in df.columns:
            key = " ".join(str(col).strip().lower().split())
            if key in canonical_cols:
                rename_map[col] = canonical_cols[key]
        if rename_map:
            df = df.rename(columns=rename_map)

        text_cols = ["Zone", "Plant", "Plant Name", "Tank", "Unique Ref Id",
                     "Material", "Material Description", "Tank Type", "Tank Status"]
        for col in text_cols:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()

        for num_col in ["Tank Capacity", "Dispatches", "Turn",
                        "Opening Stock", "Receipts", "Closing Stock"]:
            if num_col in df.columns:
                df[num_col] = pd.to_numeric(df[num_col], errors="coerce").fillna(0)

        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_location_visit(path: str, cache_buster: tuple[int, int] | None = None) -> pd.DataFrame:
    """Load Location Visit report from disk and normalize critical columns.

    The revised workbook may place the usable data on a different sheet and may
    also rename headers slightly. This loader scans a few sheet/header
    combinations and standardizes the expected business columns.
    """
    try:
        del cache_buster  # used only to invalidate Streamlit cache when file changes

        import re

        def _norm_header(value: object) -> str:
            return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())

        alias_groups = {
            "Zone": ["zone", "zone name", "sbu zone", "sbuzone", "sbu_zone"],
            "Planning Plant": [
                "planning plant", "planningplant", "planning plant code",
                "planningplantcode", "plant code", "plant"
            ],
            "Plant Desc.": [
                "plant desc.", "plant desc", "plant description",
                "plant name", "location", "location name"
            ],
            "Audit Number": ["audit number", "audit no", "audit no.", "auditnumber", "audit id"],
            "Audit Start Date": ["audit start date", "audit start", "auditstartdate", "visit start date"],
            "Audit End Date": ["audit end date", "audit end", "auditenddate", "visit end date"],
            "TotalRecomms": ["totalrecomms", "total recomms", "total recommendations", "total recommendation"],
            "ClosedRecomms": ["closedrecomms", "closed recomms", "closed recommendations", "closed recommendation"],
            "OpenRecomms": ["openrecomms", "open recomms", "open recommendations", "open recommendation"],
        }
        alias_map = {
            _norm_header(alias): canonical
            for canonical, aliases in alias_groups.items()
            for alias in aliases
        }

        def _standardize_candidate(frame: pd.DataFrame) -> tuple[pd.DataFrame, set[str]]:
            candidate = frame.copy()
            candidate.columns = candidate.columns.astype(str).str.strip()
            rename_map: dict[str, str] = {}
            matched: set[str] = set()
            for col in candidate.columns:
                canonical = alias_map.get(_norm_header(col))
                if canonical and canonical not in matched:
                    rename_map[col] = canonical
                    matched.add(canonical)
            if rename_map:
                candidate = candidate.rename(columns=rename_map)
            return candidate, matched

        best_df = pd.DataFrame()
        best_match_count = -1

        try:
            workbook = pd.ExcelFile(path)
            candidate_sheets = workbook.sheet_names or [0]
        except Exception:
            candidate_sheets = [0]

        for sheet in candidate_sheets:
            for header_row in range(0, 4):
                try:
                    raw_df = pd.read_excel(path, sheet_name=sheet, header=header_row)
                except Exception:
                    continue
                if raw_df is None or raw_df.empty:
                    continue
                raw_df = raw_df.dropna(how="all").copy()
                if raw_df.empty:
                    continue

                candidate_df, matched = _standardize_candidate(raw_df)
                match_count = len(matched)
                if match_count > best_match_count:
                    best_df = candidate_df
                    best_match_count = match_count

                if match_count >= 6:
                    break
            if best_match_count >= 6:
                break

        if best_df.empty:
            ext = os.path.splitext(path)[1].lower()
            best_df = _read_excel_flexible(path, ext_hint=ext)
            if best_df is None or best_df.empty:
                return pd.DataFrame()
            best_df.columns = best_df.columns.astype(str).str.strip()
            best_df, _ = _standardize_candidate(best_df)

        df = best_df.copy()

        # ── Detect raw CAPA format and aggregate into one row per plant+audit ──
        # Raw format: one row per recommendation, has "CAPA Status" column.
        # Pre-aggregated format: one row per audit with numeric TotalRecomms etc.
        # This mirrors the Sr. Manager Inspection Dashboard's isRaw / aggregateFromCAPARows logic.
        _capa_col = next(
            (c for c in df.columns if c.strip() in ("CAPA Status", "Capa Status")), None
        )
        if _capa_col is not None:
            _CLOSED = {"Closed", "Completed"}
            _OPEN   = {"Open", "Reopened"}
            # Group by Plant+Audit only (matches Sr. Manager Inspection Dashboard logic)
            # Extra descriptor columns (Zone, dates) are taken as first non-null value per group
            _agg_keys = [c for c in ["Planning Plant", "Audit Number"] if c in df.columns]
            _extra    = [c for c in ["Plant Desc.", "Zone", "Audit Start Date", "Audit End Date"] if c in df.columns]
            if _agg_keys:
                _records = []
                for _vals, _grp in df.groupby(_agg_keys, dropna=False):
                    _statuses = _grp[_capa_col].astype(str).str.strip()
                    # Non-blank statuses (include all; "nan" from NaN is excluded as it means no data)
                    _non_blank = _statuses[_statuses != ""]
                    _valid     = _non_blank[_non_blank.str.lower() != "nan"]
                    _row = dict(zip(_agg_keys, _vals if isinstance(_vals, tuple) else [_vals]))
                    for _ec in _extra:
                        _nonnull = _grp[_ec].dropna()
                        _row[_ec] = _nonnull.iloc[0] if not _nonnull.empty else ""
                    # Include audit even if all CAPA rows are blank (TotalRecomms=0)
                    # Matches Sr. Manager logic which counts all Plant+Audit combos
                    _row["TotalRecomms"]  = len(_valid)
                    _row["ClosedRecomms"] = int(_valid.isin(_CLOSED).sum())
                    _row["OpenRecomms"]   = int(_valid.isin(_OPEN).sum())
                    _records.append(_row)
                df = pd.DataFrame(_records)
                # Remap SWZ → COZ (matches Sr. Manager zone normalisation)
                if "Zone" in df.columns:
                    df["Zone"] = df["Zone"].replace("SWZ", "COZ")

        for col in ["Planning Plant", "Plant Desc.", "Audit Number", "Zone"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()
                if col == "Planning Plant":
                    df[col] = df[col].str.replace(r"\.0$", "", regex=True)

        for col in ["TotalRecomms", "ClosedRecomms", "OpenRecomms"]:
            if col in df.columns:
                # Convert to proper numeric; if already aggregated integers they pass through cleanly
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

        for col in ["Audit Start Date", "Audit End Date"]:
            if col in df.columns:
                dt = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
                df[col] = dt.dt.strftime("%d/%m/%Y").fillna("")

        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_pipeline_stock(path: str, cache_buster: tuple[int, int] | None = None) -> pd.DataFrame:
    """Load Pipeline Stock report and normalize key columns."""
    try:
        del cache_buster  # used only to invalidate Streamlit cache when file changes
        ext = os.path.splitext(path)[1].lower()
        df = _read_excel_flexible(path, ext_hint=ext)
        df.columns = df.columns.astype(str).str.strip()

        canonical_cols = {
            "material": "Material",
            "plant": "Plant",
            "storage location": "Storage location",
            "base unit of measure": "Base Unit of Measure",
            "unrestricted": "Unrestricted",
            "blocked": "Blocked",
            "zone": "Zone",
        }
        rename_map = {}
        for col in df.columns:
            key = " ".join(str(col).strip().lower().split())
            if key in canonical_cols:
                rename_map[col] = canonical_cols[key]
        if rename_map:
            df = df.rename(columns=rename_map)

        text_cols = ["Material", "Plant", "Storage location", "Base Unit of Measure", "Zone"]
        for col in text_cols:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()

        for num_col in ["Unrestricted", "Blocked"]:
            if num_col in df.columns:
                df[num_col] = pd.to_numeric(df[num_col], errors="coerce").fillna(0)

        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def _filter_strictly_mapped_rows(df: pd.DataFrame, source_code_col: str = "") -> tuple[pd.DataFrame, list]:
    """Keep only rows mapped to PlantMaster Zone+Plant; return filtered rows and excluded source codes."""
    if df is None or df.empty:
        return pd.DataFrame(), []

    if "Plant Name" not in df.columns or "Zone Name" not in df.columns:
        return df.copy(), []

    plant_series = df["Plant Name"].astype(str).str.strip()
    zone_series = df["Zone Name"].astype(str).str.strip()

    valid_mask = (
        df["Plant Name"].notna()
        & df["Zone Name"].notna()
        & (plant_series != "")
        & (zone_series != "")
        & (~plant_series.str.lower().isin(["nan", "none"]))
        & (~zone_series.str.lower().isin(["nan", "none"]))
    )

    excluded_codes = []
    if source_code_col and source_code_col in df.columns:
        excluded_codes = (
            df.loc[~valid_mask, source_code_col]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )

    return df.loc[valid_mask].copy(), excluded_codes


def process_pending_dc(
    df_dc        : pd.DataFrame,
    df_plant     : pd.DataFrame,
    zone_filter  : list = None,
    plant_filter : list = None,
    as_of_date   = None,
) -> dict:
    """
    Process raw Pending DC data into aggregated exception metrics.

    1. De-dup on (SENDING PLANT, SHIPMENT): each unique shipment = 1 pending DC.
    2. Left-join with PlantMaster to get Plant Name & Zone Name.
    3. Apply optional sidebar filters.
    4. Aggregate at plant level and zone level.
    """
    EMPTY = {
        "total_count" : 0,
        "summary_df"  : pd.DataFrame(),
        "zone_summary": pd.DataFrame(),
        "detail_df"   : pd.DataFrame(),
        "unmatched"   : [],
    }

    if df_dc is None or df_dc.empty:
        return EMPTY

    required = {"SENDING PLANT", "SHIPMENT"}
    missing  = required - set(df_dc.columns)
    if missing:
        st.warning(f"⚠️ Pending DC file is missing expected columns: {missing}")
        return EMPTY

    # As-of-date filter: use first available date column in the source file
    # BILLING DATE is the primary date in the actual Pending DC SAP export.
    _dc_date_candidates = ["BILLING DATE", "DOCUMENT DATE", "DISPATCH DATE",
                           "PLANNED GI DATE", "PLANNED GOODS ISSUE DATE",
                           "SHIPMENT DATE", "DELIVERY DATE"]
    if as_of_date is not None:
        _aod = pd.Timestamp(as_of_date)
        _dc_date_col = next((c for c in _dc_date_candidates if c in df_dc.columns), None)
        if _dc_date_col:
            _parsed = pd.to_datetime(df_dc[_dc_date_col], errors="coerce", dayfirst=True)
            df_dc = df_dc[_parsed.isna() | (_parsed <= _aod)].copy()

    dc_unique = df_dc.drop_duplicates(subset=["SENDING PLANT", "SHIPMENT"]).copy()
    dc_unique["SENDING PLANT"] = dc_unique["SENDING PLANT"].astype(str).str.strip()
    dc_unique["SHIPMENT"]      = dc_unique["SHIPMENT"].astype(str).str.strip()

    # Step 1b: de-duplicate FOR DISPLAY — keep all SHIPMENT+MATERIAL combos (true unique records)
    _detail_dedup_cols = ["SENDING PLANT", "SHIPMENT"]
    if "MATERIAL" in df_dc.columns:
        _detail_dedup_cols.append("MATERIAL")
    df_detail = df_dc.drop_duplicates(subset=_detail_dedup_cols).copy()
    df_detail["SENDING PLANT"] = df_detail["SENDING PLANT"].astype(str).str.strip()
    df_detail["SHIPMENT"]      = df_detail["SHIPMENT"].astype(str).str.strip()

    # Step 2: map to PlantMaster
    plant_map = (
        df_plant[["Plant Code", "Plant Name", "Zone Name"]]
        .copy()
        .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip()})
    )
    merged = dc_unique.merge(
        plant_map,
        left_on  = "SENDING PLANT",
        right_on = "Plant Code",
        how      = "left",
    )
    merged, _ = _filter_strictly_mapped_rows(merged, "SENDING PLANT")

    # Step 2b: merge detail (material-level) with PlantMaster for display
    detail_merged = df_detail.merge(
        plant_map,
        left_on  = "SENDING PLANT",
        right_on = "Plant Code",
        how      = "left",
    )
    detail_merged, _ = _filter_strictly_mapped_rows(detail_merged, "SENDING PLANT")

    # Step 3: filters
    if zone_filter:
        merged        = merged[merged["Zone Name"].isin(zone_filter)]
        detail_merged = detail_merged[detail_merged["Zone Name"].isin(zone_filter)]
    if plant_filter:
        merged        = merged[merged["Plant Name"].isin(plant_filter)]
        detail_merged = detail_merged[detail_merged["Plant Name"].isin(plant_filter)]

    if merged.empty:
        return EMPTY

    # Step 4: aggregate
    summary_df = (
        merged
        .groupby(["Zone Name", "Plant Name"], sort=True)
        .agg(pending_dc=("SHIPMENT", "nunique"))
        .reset_index()
        .rename(columns={"pending_dc": "Pending DC Count"})
        .sort_values(["Zone Name", "Pending DC Count"], ascending=[True, False])
    )

    zone_summary = (
        summary_df
        .groupby("Zone Name")
        .agg(
            Plants     = ("Plant Name",      "nunique"),
            pending_dc = ("Pending DC Count", "sum"),
        )
        .reset_index()
        .rename(columns={"pending_dc": "Pending DC Count"})
        .sort_values("Pending DC Count", ascending=False)
    )

    return {
        "total_count" : int(merged["SHIPMENT"].nunique()),
        "summary_df"  : summary_df,
        "zone_summary": zone_summary,
        "detail_df"   : detail_merged,
        "unmatched"   : [],
    }


def process_open_deliveries(
    df_open      : pd.DataFrame,
    df_plant     : pd.DataFrame,
    zone_filter  : list = None,
    plant_filter : list = None,
    as_of_date   = None,
) -> dict:
    """
    Process raw Open Delivery data into KPI + drill-down outputs.

    1. Map Shipping Point/Receiving Pt with PlantMaster Plant Code.
    2. Open Delivery count = unique Delivery numbers.
    3. Apply optional sidebar filters.
    4. Build zone/plant summaries and detail with Delivery Age (Days).
    """
    EMPTY = {
        "total_count" : 0,
        "summary_df"  : pd.DataFrame(),
        "zone_summary": pd.DataFrame(),
        "detail_df"   : pd.DataFrame(),
        "unmatched"   : [],
    }

    if df_open is None or df_open.empty:
        return EMPTY

    ship_col   = "SHIPPING POINT/RECEIVING PT"
    deliv_col  = "DELIVERY"
    vol_col    = "VOLUME"
    gi_date_col = "GOODS ISSUE DATE"

    required = {ship_col, deliv_col}
    missing  = required - set(df_open.columns)
    if missing:
        st.warning(f"⚠️ Open Delivery file is missing expected columns: {missing}")
        return EMPTY

    work = df_open.copy()
    work[ship_col]  = work[ship_col].astype(str).str.strip()
    work[deliv_col] = work[deliv_col].astype(str).str.strip()
    work = work[work[deliv_col] != ""]

    # As-of-date filter: use PICKING DATE as the entry-time proxy (no future
    # placeholder dates), falling back to LOADING DATE then GOODS ISSUE DATE.
    # GOODS ISSUE DATE can have SAP month-end placeholders (e.g. 2026-12-31)
    # which would incorrectly exclude valid open deliveries.
    _od_date_candidates = ["PICKING DATE", "LOADING DATE", gi_date_col]
    _od_date_col = next((c for c in _od_date_candidates if c in work.columns), None)
    if as_of_date is not None and _od_date_col:
        _aod = pd.Timestamp(as_of_date)
        _parsed = pd.to_datetime(work[_od_date_col], errors="coerce", dayfirst=True)
        work = work[_parsed.isna() | (_parsed <= _aod)].copy()

    # Keep one row per shipping-point + delivery combination for drill-down accuracy.
    detail_base = work.drop_duplicates(subset=[ship_col, deliv_col]).copy()

    plant_map = (
        df_plant[["Plant Code", "Plant Name", "Zone Name"]]
        .copy()
        .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip()})
    )

    merged = detail_base.merge(
        plant_map,
        left_on  = ship_col,
        right_on = "Plant Code",
        how      = "left",
    )
    merged, _ = _filter_strictly_mapped_rows(merged, ship_col)

    if zone_filter:
        merged = merged[merged["Zone Name"].isin(zone_filter)]
    if plant_filter:
        merged = merged[merged["Plant Name"].isin(plant_filter)]

    if merged.empty:
        return EMPTY

    if gi_date_col in merged.columns:
        merged[gi_date_col] = pd.to_datetime(merged[gi_date_col], errors="coerce", dayfirst=True)
    else:
        merged[gi_date_col] = pd.NaT

    if vol_col in merged.columns:
        merged[vol_col] = pd.to_numeric(merged[vol_col], errors="coerce")
    else:
        merged[vol_col] = pd.NA

    today = pd.Timestamp(datetime.now().date())
    merged["DELIVERY AGE (DAYS)"] = (today - merged[gi_date_col]).dt.days
    merged.loc[merged["DELIVERY AGE (DAYS)"] < 0, "DELIVERY AGE (DAYS)"] = pd.NA

    summary_df = (
        merged
        .groupby(["Zone Name", "Plant Name"], sort=True)
        .agg(open_delivery_count=(deliv_col, "nunique"))
        .reset_index()
        .rename(columns={"open_delivery_count": "Open Delivery Count"})
        .sort_values(["Zone Name", "Open Delivery Count"], ascending=[True, False])
    )

    zone_summary = (
        summary_df
        .groupby("Zone Name")
        .agg(
            Plants        = ("Plant Name", "nunique"),
            open_delivery = ("Open Delivery Count", "sum"),
        )
        .reset_index()
        .rename(columns={"open_delivery": "Open Delivery Count"})
        .sort_values("Open Delivery Count", ascending=False)
    )

    detail_df = merged.copy().rename(
        columns={
            ship_col: "Shipping Point/Receiving Pt",
            deliv_col: "Delivery",
            vol_col: "Volume",
            gi_date_col: "Goods Issue Date",
            "DELIVERY AGE (DAYS)": "Delivery Age (Days)",
        }
    )
    if "Goods Issue Date" in detail_df.columns:
        detail_df["Goods Issue Date"] = detail_df["Goods Issue Date"].dt.strftime("%d-%m-%Y")
        detail_df["Goods Issue Date"] = detail_df["Goods Issue Date"].fillna("")

    return {
        "total_count" : int(merged[deliv_col].nunique()),
        "summary_df"  : summary_df,
        "zone_summary": zone_summary,
        "detail_df"   : detail_df,
        "unmatched"   : [],
    }


def process_open_intransit(
    df_intransit : pd.DataFrame,
    df_plant     : pd.DataFrame,
    zone_filter  : list = None,
    plant_filter : list = None,
    as_of_date   = None,
) -> dict:
    """
    Process Open In-Transit data into KPI + drill-down outputs.

    Mapping: Sending Plant -> PlantMaster Plant Code.
    KPI count: unique STO Order.
    """
    EMPTY = {
        "total_count" : 0,
        "summary_df"  : pd.DataFrame(),
        "zone_summary": pd.DataFrame(),
        "detail_df"   : pd.DataFrame(),
        "unmatched"   : [],
    }

    if df_intransit is None or df_intransit.empty:
        return EMPTY

    send_col      = "SENDING PLANT"
    sto_col       = "STO ORDER"
    recv_col      = "RECEIVING PLANT"
    disp_col      = "DISPATCH DATE"
    inco_col      = "INCO TERMS"
    delivery_col  = "DELIVERY"
    shipment_col  = "SHIPMENT"
    invoice_col   = "INVOICE"
    net_value_col = "NET VALUE"
    material_col  = "MATERIAL"
    mat_desc_col  = "MATERIAL DESCRIPTION"
    load_qty_col  = "LOAD QUANTITY"
    open_qty_col  = "OPEN QUANTITY"

    required = {send_col, sto_col}
    missing  = required - set(df_intransit.columns)
    if missing:
        st.warning(f"⚠️ Open In-Transit file is missing expected columns: {missing}")
        return EMPTY

    work = df_intransit.copy()
    work[send_col] = work[send_col].astype(str).str.strip()
    work[sto_col]  = work[sto_col].astype(str).str.strip()
    work = work[work[sto_col] != ""]

    # As-of-date filter on DISPATCH DATE
    if as_of_date is not None and disp_col in work.columns:
        _aod = pd.Timestamp(as_of_date)
        _parsed = pd.to_datetime(work[disp_col], errors="coerce", dayfirst=True)
        work = work[_parsed.isna() | (_parsed <= _aod)].copy()

    dedup_cols = [send_col, sto_col]
    for c in [delivery_col, shipment_col, invoice_col, material_col]:
        if c in work.columns:
            dedup_cols.append(c)
    detail_base = work.drop_duplicates(subset=dedup_cols).copy()

    plant_map = (
        df_plant[["Plant Code", "Plant Name", "Zone Name"]]
        .copy()
        .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip()})
    )

    merged = detail_base.merge(
        plant_map,
        left_on  = send_col,
        right_on = "Plant Code",
        how      = "left",
    )
    merged, _ = _filter_strictly_mapped_rows(merged, send_col)

    if zone_filter:
        merged = merged[merged["Zone Name"].isin(zone_filter)]
    if plant_filter:
        merged = merged[merged["Plant Name"].isin(plant_filter)]

    if merged.empty:
        return EMPTY

    if disp_col in merged.columns:
        merged[disp_col] = pd.to_datetime(merged[disp_col], errors="coerce", dayfirst=True)
    else:
        merged[disp_col] = pd.NaT

    today = pd.Timestamp(datetime.now().date())
    merged["IN-TRANSIT AGE (DAYS)"] = (today - merged[disp_col]).dt.days
    merged.loc[merged["IN-TRANSIT AGE (DAYS)"] < 0, "IN-TRANSIT AGE (DAYS)"] = pd.NA

    summary_df = (
        merged
        .groupby(["Zone Name", "Plant Name"], sort=True)
        .agg(open_intransit_count=(sto_col, "nunique"))
        .reset_index()
        .rename(columns={"open_intransit_count": "Open In-Transit STO Count"})
        .sort_values(["Zone Name", "Open In-Transit STO Count"], ascending=[True, False])
    )

    zone_summary = (
        summary_df
        .groupby("Zone Name")
        .agg(
            Plants         = ("Plant Name", "nunique"),
            open_intransit = ("Open In-Transit STO Count", "sum"),
        )
        .reset_index()
        .rename(columns={"open_intransit": "Open In-Transit STO Count"})
        .sort_values("Open In-Transit STO Count", ascending=False)
    )

    rename_map = {
        sto_col       : "STO Order",
        recv_col      : "Receiving Plant",
        disp_col      : "Dispatch Date",
        inco_col      : "Inco Terms",
        delivery_col  : "Delivery",
        shipment_col  : "Shipment",
        invoice_col   : "Invoice",
        net_value_col : "Net Value",
        material_col  : "Material",
        mat_desc_col  : "Material Description",
        load_qty_col  : "Load Quantity",
        open_qty_col  : "Open Quantity",
        "IN-TRANSIT AGE (DAYS)": "In-Transit Age (Days)",
    }
    detail_df = merged.copy().rename(columns={k: v for k, v in rename_map.items() if k in merged.columns})
    if "Dispatch Date" in detail_df.columns:
        detail_df["Dispatch Date"] = detail_df["Dispatch Date"].dt.strftime("%d-%m-%Y")
        detail_df["Dispatch Date"] = detail_df["Dispatch Date"].fillna("")

    return {
        "total_count" : int(merged[sto_col].nunique()),
        "summary_df"  : summary_df,
        "zone_summary": zone_summary,
        "detail_df"   : detail_df,
        "unmatched"   : [],
    }


def process_open_sales_orders(
    df_so        : pd.DataFrame,
    df_plant     : pd.DataFrame,
    zone_filter  : list = None,
    plant_filter : list = None,
    as_of_date   = None,
) -> dict:
    """
    Process Open Sales Orders into KPI + drill-down outputs.

    Mapping: Shipping Point/Receiving Pt -> PlantMaster Plant Code.
    KPI count: unique Sales document.
    """
    EMPTY = {
        "total_count" : 0,
        "summary_df"  : pd.DataFrame(),
        "zone_summary": pd.DataFrame(),
        "detail_df"   : pd.DataFrame(),
        "unmatched"   : [],
    }

    if df_so is None or df_so.empty:
        return EMPTY

    ship_col       = "SHIPPING POINT/RECEIVING PT"
    so_col         = "SALES DOCUMENT"
    so_type_col    = "SALES DOCUMENT TYPE"
    sold_to_col    = "SOLD-TO PARTY"
    sold_to_nm_col = "SOLD-TO PARTY NAME"
    material_col   = "MATERIAL"
    mat_desc_col   = "MATERIAL DESCRIPTION"
    ord_qty_col    = "ORDER QUANTITY (ITEM)"
    sales_unit_col = "SALES UNIT"
    doc_date_col   = "DOCUMENT DATE"
    net_val_col    = "NET VALUE (ITEM)"
    conf_qty_col   = "CONFIRMED QUANTITY (ITEM)"

    required = {ship_col, so_col}
    missing  = required - set(df_so.columns)
    if missing:
        st.warning(f"⚠️ Open Sales Orders file is missing expected columns: {missing}")
        return EMPTY

    work = df_so.copy()
    work[ship_col] = work[ship_col].astype(str).str.strip()
    work[so_col]   = work[so_col].astype(str).str.strip()
    work = work[work[so_col] != ""]

    # As-of-date filter on DOCUMENT DATE
    if as_of_date is not None and doc_date_col in work.columns:
        _aod = pd.Timestamp(as_of_date)
        _parsed = pd.to_datetime(work[doc_date_col], errors="coerce", dayfirst=True)
        work = work[_parsed.isna() | (_parsed <= _aod)].copy()

    dedup_cols = [ship_col, so_col]
    for c in [material_col, so_type_col, sold_to_col]:
        if c in work.columns:
            dedup_cols.append(c)
    detail_base = work.drop_duplicates(subset=dedup_cols).copy()

    plant_map = (
        df_plant[["Plant Code", "Plant Name", "Zone Name"]]
        .copy()
        .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip()})
    )

    merged = detail_base.merge(
        plant_map,
        left_on  = ship_col,
        right_on = "Plant Code",
        how      = "left",
    )
    merged, _ = _filter_strictly_mapped_rows(merged, ship_col)

    if zone_filter:
        merged = merged[merged["Zone Name"].isin(zone_filter)]
    if plant_filter:
        merged = merged[merged["Plant Name"].isin(plant_filter)]

    if merged.empty:
        return EMPTY

    if doc_date_col in merged.columns:
        merged[doc_date_col] = pd.to_datetime(merged[doc_date_col], errors="coerce", dayfirst=True)
    else:
        merged[doc_date_col] = pd.NaT

    today = pd.Timestamp(datetime.now().date())
    merged["SALES ORDER AGE (DAYS)"] = (today - merged[doc_date_col]).dt.days
    merged.loc[merged["SALES ORDER AGE (DAYS)"] < 0, "SALES ORDER AGE (DAYS)"] = pd.NA

    summary_df = (
        merged
        .groupby(["Zone Name", "Plant Name"], sort=True)
        .agg(open_so_count=(so_col, "nunique"))
        .reset_index()
        .rename(columns={"open_so_count": "Open Sales Order Count"})
        .sort_values(["Zone Name", "Open Sales Order Count"], ascending=[True, False])
    )

    zone_summary = (
        summary_df
        .groupby("Zone Name")
        .agg(
            Plants  = ("Plant Name", "nunique"),
            open_so = ("Open Sales Order Count", "sum"),
        )
        .reset_index()
        .rename(columns={"open_so": "Open Sales Order Count"})
        .sort_values("Open Sales Order Count", ascending=False)
    )

    rename_map = {
        so_col         : "Sales Document",
        so_type_col    : "Sales Document Type",
        sold_to_col    : "Sold-to Party",
        sold_to_nm_col : "Sold-to Party Name",
        material_col   : "Material",
        mat_desc_col   : "Material Description",
        ord_qty_col    : "Order Quantity (Item)",
        sales_unit_col : "Sales Unit",
        doc_date_col   : "Document Date",
        net_val_col    : "Net Value (Item)",
        ship_col       : "Shipping Point/Receiving Pt",
        conf_qty_col   : "Confirmed Quantity (Item)",
        "SALES ORDER AGE (DAYS)": "Sales Order Age (Days)",
    }
    detail_df = merged.copy().rename(columns={k: v for k, v in rename_map.items() if k in merged.columns})
    if "Document Date" in detail_df.columns:
        detail_df["Document Date"] = detail_df["Document Date"].dt.strftime("%d-%m-%Y")
        detail_df["Document Date"] = detail_df["Document Date"].fillna("")

    return {
        "total_count" : int(merged[so_col].nunique()),
        "summary_df"  : summary_df,
        "zone_summary": zone_summary,
        "detail_df"   : detail_df,
        "unmatched"   : [],
    }


def process_pending_invoices(
    df_inv      : pd.DataFrame,
    df_plant    : pd.DataFrame,
    zone_filter : list = None,
    plant_filter: list = None,
    as_of_date  = None,
) -> dict:
    """
    Process Pending Invoices into KPI + drill-down outputs.

    Mapping: Sending Location -> PlantMaster Plant Code.
    KPI count: unique Delivery.
    """
    EMPTY = {
        "total_count" : 0,
        "summary_df"  : pd.DataFrame(),
        "zone_summary": pd.DataFrame(),
        "detail_df"   : pd.DataFrame(),
        "unmatched"   : [],
    }

    if df_inv is None or df_inv.empty:
        return EMPTY

    send_col       = "SENDING LOCATION"
    recv_col       = "RECEIVING LOCATION"
    mot_col        = "MOT"
    po_col         = "PURCHASE ORDER"
    td_ship_col    = "TD SHIPMENT"
    delivery_col   = "DELIVERY"
    mat_doc_col    = "MATERIAL DOCUMENT"
    qty_col        = "QUANTITY"
    created_by_col = "CREATED BY"
    desc_col       = "DESCRIPTION"
    created_dt_col = "CREATED DATE"

    required = {send_col, delivery_col}
    missing  = required - set(df_inv.columns)
    if missing:
        st.warning(f"⚠️ Pending Invoices file is missing expected columns: {missing}")
        return EMPTY

    work = df_inv.copy()
    work[send_col]     = work[send_col].astype(str).str.strip()
    work[delivery_col] = work[delivery_col].astype(str).str.strip()
    work = work[work[delivery_col] != ""]

    # As-of-date filter on CREATED DATE
    if as_of_date is not None and created_dt_col in work.columns:
        _aod = pd.Timestamp(as_of_date)
        _parsed = pd.to_datetime(work[created_dt_col], errors="coerce", dayfirst=True)
        work = work[_parsed.isna() | (_parsed <= _aod)].copy()

    dedup_cols = [send_col, delivery_col]
    for c in [mat_doc_col, td_ship_col, po_col]:
        if c in work.columns:
            dedup_cols.append(c)
    detail_base = work.drop_duplicates(subset=dedup_cols).copy()

    plant_map = (
        df_plant[["Plant Code", "Plant Name", "Zone Name"]]
        .copy()
        .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip()})
    )

    merged = detail_base.merge(
        plant_map,
        left_on  = send_col,
        right_on = "Plant Code",
        how      = "left",
    )
    merged, _ = _filter_strictly_mapped_rows(merged, send_col)

    if zone_filter:
        merged = merged[merged["Zone Name"].isin(zone_filter)]
    if plant_filter:
        merged = merged[merged["Plant Name"].isin(plant_filter)]

    if merged.empty:
        return EMPTY

    if created_dt_col in merged.columns:
        merged[created_dt_col] = pd.to_datetime(merged[created_dt_col], errors="coerce", dayfirst=True)
    else:
        merged[created_dt_col] = pd.NaT

    today = pd.Timestamp(datetime.now().date())
    merged["INVOICE AGE (DAYS)"] = (today - merged[created_dt_col]).dt.days
    merged.loc[merged["INVOICE AGE (DAYS)"] < 0, "INVOICE AGE (DAYS)"] = pd.NA

    summary_df = (
        merged
        .groupby(["Zone Name", "Plant Name"], sort=True)
        .agg(pending_invoice_count=(delivery_col, "nunique"))
        .reset_index()
        .rename(columns={"pending_invoice_count": "Pending Invoice Count"})
        .sort_values(["Zone Name", "Pending Invoice Count"], ascending=[True, False])
    )

    zone_summary = (
        summary_df
        .groupby("Zone Name")
        .agg(
            Plants          = ("Plant Name", "nunique"),
            pending_invoice = ("Pending Invoice Count", "sum"),
        )
        .reset_index()
        .rename(columns={"pending_invoice": "Pending Invoice Count"})
        .sort_values("Pending Invoice Count", ascending=False)
    )

    rename_map = {
        send_col       : "Sending Location",
        recv_col       : "Receiving Location",
        mot_col        : "MOT",
        po_col         : "Purchase Order",
        td_ship_col    : "TD Shipment",
        delivery_col   : "Delivery",
        mat_doc_col    : "Material Document",
        qty_col        : "Quantity",
        created_by_col : "Created By",
        desc_col       : "Description",
        created_dt_col : "Created Date",
        "INVOICE AGE (DAYS)": "Invoice Age (Days)",
    }
    detail_df = merged.copy().rename(columns={k: v for k, v in rename_map.items() if k in merged.columns})
    if "Created Date" in detail_df.columns:
        detail_df["Created Date"] = detail_df["Created Date"].dt.strftime("%d-%m-%Y")
        detail_df["Created Date"] = detail_df["Created Date"].fillna("")

    return {
        "total_count" : int(merged[delivery_col].nunique()),
        "summary_df"  : summary_df,
        "zone_summary": zone_summary,
        "detail_df"   : detail_df,
        "unmatched"   : [],
    }


def process_tank_reco(
    df_tank      : pd.DataFrame,
    df_plant     : pd.DataFrame,
    zone_filter  : list = None,
    plant_filter : list = None,
) -> dict:
    """
    Process Tank Reco data into KPI + drill-down outputs.

    Mapping: Plant -> PlantMaster Plant Code.
    KPI count: unique Plant + Tank + Material combinations.
    """

    EMPTY = {
        "total_count" : 0,
        "summary_df"  : pd.DataFrame(),
        "zone_summary": pd.DataFrame(),
        "detail_df"   : pd.DataFrame(),
        "unmatched"   : [],
    }

    if df_tank is None or df_tank.empty:
        return EMPTY

    def _pick_col(candidates: list, required: bool = False, label: str = ""):
        for c in candidates:
            if c in df_tank.columns:
                return c
        if required:
            st.warning(
                f"⚠️ Abnormal Variations in SAP file is missing expected column for {label}: {candidates}"
            )
        return None

    plant_col    = _pick_col(["PLANT"], required=True, label="Plant")
    tank_col     = _pick_col(["TANK NO.", "TANK NO", "TANK NUMBER", "TANK"], required=True, label="Tank")
    material_col = _pick_col(["MATERIAL CODE", "MATERIAL", "MATERIAL NO", "MATERIAL NUMBER"], required=True, label="Material")

    if not all([plant_col, tank_col, material_col]):
        return EMPTY

    dip_date_col      = _pick_col(["DIP DATE"])
    dip_type_col      = _pick_col(["DIP TYPE"])
    reco_status_col   = _pick_col(["RECO STATUS"])
    reco_init_col     = _pick_col(["RECO INITIATOR"])
    physical_stock_col = _pick_col(["PHYSICAL STOCK"])
    book_dip_col      = _pick_col(["BOOK STOCK@DIP", "BOOK STOCK @ DIP"])
    book_post_col     = _pick_col(["BOOK STOCK@POSTING", "BOOK STOCK @ POSTING"])
    phy_inv_col       = _pick_col(["PHY INV DOC", "PHY. INV DOC", "PHYSICAL INV DOC"])
    gain_loss_col     = _pick_col(["GAIN/LOSS BOOKED", "GAIN LOSS BOOKED"])
    type_col          = _pick_col(["TYPE"])
    posting_date_col  = _pick_col(["POSTING DATE"])
    mat_doc_col       = _pick_col(["MATERIAL DOC NO", "MATERIAL DOC. NO", "MATERIAL DOC NO."])
    mat_doc_year_col  = _pick_col(["MATERIAL DOC. YEAR", "MATERIAL DOC YEAR"])
    reco_approver_col = _pick_col(["RECO APPROVER"])
    approval_date_col = _pick_col(["APPROVAL DATE"])
    comments_col      = _pick_col(["COMMENTS FOR ABNORMAL G/L"])
    desc_reason_col   = _pick_col(["DESC. OF REASON", "DESC OF REASON"])
    remarks_col       = _pick_col(["REMARKS FOR MANUAL DIP"])

    work = df_tank.copy()
    work[plant_col]    = work[plant_col].astype(str).str.strip()
    work[tank_col]     = work[tank_col].astype(str).str.strip()
    work[material_col] = work[material_col].astype(str).str.strip()
    work = work[
        (work[plant_col] != "")
        & (work[tank_col] != "")
        & (work[material_col] != "")
    ]

    work["TANK_RECO_KEY"] = (
        work[plant_col] + "_" + work[tank_col] + "_" + work[material_col]
    )
    detail_base = work.drop_duplicates(subset=["TANK_RECO_KEY"]).copy()

    plant_map = (
        df_plant[["Plant Code", "Plant Name", "Zone Name"]]
        .copy()
        .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip()})
    )

    merged = detail_base.merge(
        plant_map,
        left_on  = plant_col,
        right_on = "Plant Code",
        how      = "left",
    )
    merged, _ = _filter_strictly_mapped_rows(merged, plant_col)

    if zone_filter:
        merged = merged[merged["Zone Name"].isin(zone_filter)]
    if plant_filter:
        merged = merged[merged["Plant Name"].isin(plant_filter)]

    if merged.empty:
        return EMPTY

    for dt_col in [dip_date_col, posting_date_col, approval_date_col]:
        if dt_col and dt_col in merged.columns:
            merged[dt_col] = pd.to_datetime(merged[dt_col], errors="coerce", dayfirst=True)

    summary_df = (
        merged
        .groupby(["Zone Name", "Plant Name"], sort=True)
        .agg(tank_reco_count=("TANK_RECO_KEY", "nunique"))
        .reset_index()
    )
    # Support both column names for backward compatibility
    if "Tank Reco Count" not in summary_df.columns and "tank_reco_count" in summary_df.columns:
        summary_df = summary_df.rename(columns={"tank_reco_count": "Tank Reco Count"})
    summary_df = summary_df.sort_values(["Zone Name", "Tank Reco Count"], ascending=[True, False])

    zone_summary = (
        summary_df
        .groupby("Zone Name")
        .agg(
            Plants    = ("Plant Name", "nunique"),
            tank_reco = ("Tank Reco Count", "sum"),
        )
        .reset_index()
        .rename(columns={"tank_reco": "Tank Reco Count"})
        .sort_values("Tank Reco Count", ascending=False)
    )

    rename_map = {
        plant_col         : "Plant",
        tank_col          : "Tank No.",
        material_col      : "Material Code",
        dip_date_col      : "Dip Date",
        dip_type_col      : "Dip Type",
        reco_status_col   : "Reco Status",
        reco_init_col     : "Reco Initiator",
        physical_stock_col: "Physical Stock",
        book_dip_col      : "Book Stock @ Dip",
        book_post_col     : "Book Stock @ Posting",
        phy_inv_col       : "Phy Inv Doc",
        gain_loss_col     : "Gain/Loss Booked",
        type_col          : "Type",
        posting_date_col  : "Posting Date",
        mat_doc_col       : "Material Doc No",
        mat_doc_year_col  : "Material Doc Year",
        reco_approver_col : "Reco Approver",
        approval_date_col : "Approval Date",
        comments_col      : "Comments for Abnormal G/L",
        desc_reason_col   : "Description of Reason",
        remarks_col       : "Remarks for Manual Dip",
        "TANK_RECO_KEY"  : "Tank Reco Key",
        "Abnormal Variations Key": "Tank Reco Key",
    }
    detail_df = merged.copy().rename(columns={k: v for k, v in rename_map.items() if k in merged.columns or k in ["TANK_RECO_KEY", "Abnormal Variations Key"]})

    for c in ["Dip Date", "Posting Date", "Approval Date"]:
        if c in detail_df.columns:
            detail_df[c] = detail_df[c].dt.strftime("%d-%m-%Y")
            detail_df[c] = detail_df[c].fillna("")

    return {
        "total_count" : int(merged["TANK_RECO_KEY"].nunique()),
        "summary_df"  : summary_df,
        "zone_summary": zone_summary,
        "detail_df"   : detail_df,
        "unmatched"   : [],
    }


def process_open_shortages_sales(
    df_short_sales: pd.DataFrame,
    df_plant      : pd.DataFrame,
    zone_filter   : list = None,
    plant_filter  : list = None,
    as_of_date    = None,
) -> dict:
    """
    Process OPEN SHORTAGES - Ltrs (Sales) into KPI + drill-down outputs.

    Mapping: Plant -> PlantMaster Plant Code.
    KPI value: sum of Shortage Quantity (in Ltrs).
    """
    EMPTY = {
        "total_count" : 0.0,
        "summary_df"  : pd.DataFrame(),
        "zone_summary": pd.DataFrame(),
        "detail_df"   : pd.DataFrame(),
        "unmatched"   : [],
    }

    if df_short_sales is None or df_short_sales.empty:
        return EMPTY

    def _pick_col(candidates: list, required: bool = False, label: str = ""):
        for c in candidates:
            if c in df_short_sales.columns:
                return c
        if required:
            st.warning(
                f"⚠️ OPEN SHORTAGES - Ltrs (Sales) file is missing expected column for {label}: {candidates}"
            )
        return None

    plant_col = _pick_col(["PLANT"], required=True, label="Plant")
    shortage_col = _pick_col(
        ["SHORTAGE QUANTITY (IN LTRS)", "SHORTAGE QUANTITY", "SHORTAGE QTY"],
        required=True,
        label="Shortage Quantity (in Ltrs)",
    )
    created_on_col = _pick_col(["CREATED ON"], required=True, label="Created on")

    billing_doc_col = _pick_col(["BILLING DOCUMENT"])
    shipment_col = _pick_col(["SHIPMENT NUMBER"])
    sold_to_col = _pick_col(["SOLD-TO PARTY", "SOLD TO PARTY"])
    service_agent_col = _pick_col(["SERVICE AGENT"])
    sales_org_col = _pick_col(["SALES ORGANIZATION"])
    delivery_col = _pick_col(["DELIVERY"])
    material_col = _pick_col(["MATERIAL"])
    billed_qty_col = _pick_col(["BILLED QUANTITY"])
    tt_col = _pick_col(["COLUMN M", "UNNAMED: 12", "TT NUMBER"])

    if not all([plant_col, shortage_col, created_on_col]):
        return EMPTY

    work = df_short_sales.copy()
    work[plant_col] = work[plant_col].astype(str).str.strip()
    work = work[work[plant_col] != ""]

    # As-of-date filter on CREATED ON
    if as_of_date is not None:
        _aod = pd.Timestamp(as_of_date)
        _parsed = pd.to_datetime(work[created_on_col], errors="coerce", dayfirst=True)
        work = work[_parsed.isna() | (_parsed <= _aod)].copy()

    # Ensure all critical drill-down columns exist even if source header/value is blank.
    for c in [
        billing_doc_col,
        shipment_col,
        sold_to_col,
        service_agent_col,
        sales_org_col,
        delivery_col,
        material_col,
        billed_qty_col,
    ]:
        if c is None:
            continue

    work[shortage_col] = pd.to_numeric(work[shortage_col], errors="coerce").fillna(0)

    if billed_qty_col and billed_qty_col in work.columns:
        work[billed_qty_col] = pd.to_numeric(work[billed_qty_col], errors="coerce")

    work[created_on_col] = pd.to_datetime(work[created_on_col], errors="coerce", dayfirst=True)
    today = pd.Timestamp(datetime.now().date())
    work["SHORTAGE AGE (DAYS)"] = (today - work[created_on_col]).dt.days
    work.loc[work["SHORTAGE AGE (DAYS)"] < 0, "SHORTAGE AGE (DAYS)"] = pd.NA

    if tt_col and tt_col in work.columns:
        work["TT NUMBER"] = work[tt_col].astype(str)
        work.loc[work["TT NUMBER"].str.upper().eq("NAN"), "TT NUMBER"] = ""
    else:
        work["TT NUMBER"] = ""

    plant_map = (
        df_plant[["Plant Code", "Plant Name", "Zone Name"]]
        .copy()
        .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip()})
    )

    merged = work.merge(
        plant_map,
        left_on  = plant_col,
        right_on = "Plant Code",
        how      = "left",
    )
    merged, _ = _filter_strictly_mapped_rows(merged, plant_col)

    if zone_filter:
        merged = merged[merged["Zone Name"].isin(zone_filter)]
    if plant_filter:
        merged = merged[merged["Plant Name"].isin(plant_filter)]

    if merged.empty:
        return EMPTY

    summary_df = (
        merged
        .groupby(["Zone Name", "Plant Name"], sort=True)
        .agg(shortage_qty=(shortage_col, "sum"))
        .reset_index()
        .rename(columns={"shortage_qty": "Total Shortage Quantity (in Ltrs)"})
        .sort_values(["Zone Name", "Total Shortage Quantity (in Ltrs)"], ascending=[True, False])
    )

    zone_summary = (
        summary_df
        .groupby("Zone Name")
        .agg(
            Plants      = ("Plant Name", "nunique"),
            shortage_qty = ("Total Shortage Quantity (in Ltrs)", "sum"),
        )
        .reset_index()
        .rename(columns={"shortage_qty": "Total Shortage Quantity (in Ltrs)"})
        .sort_values("Total Shortage Quantity (in Ltrs)", ascending=False)
    )

    rename_map = {
        plant_col       : "Plant",
        billing_doc_col : "Billing Document",
        shipment_col    : "Shipment Number",
        sold_to_col     : "Sold-to Party",
        service_agent_col: "Service Agent",
        sales_org_col   : "Sales Organization",
        delivery_col    : "Delivery",
        material_col    : "Material",
        billed_qty_col  : "Billed Quantity",
        shortage_col    : "Shortage Quantity (in Ltrs)",
        "TT NUMBER"    : "TT Number",
        "SHORTAGE AGE (DAYS)": "Shortage Age (Days)",
        created_on_col  : "Created on",
    }
    detail_df = merged.copy().rename(columns={k: v for k, v in rename_map.items() if k in merged.columns})

    if "Created on" in detail_df.columns:
        detail_df["Created on"] = detail_df["Created on"].dt.strftime("%d-%m-%Y")
        detail_df["Created on"] = detail_df["Created on"].fillna("")

    if "Billed Quantity" in detail_df.columns:
        detail_df["Billed Quantity"] = pd.to_numeric(detail_df["Billed Quantity"], errors="coerce")
    if "Shortage Quantity (in Ltrs)" in detail_df.columns:
        detail_df["Shortage Quantity (in Ltrs)"] = pd.to_numeric(
            detail_df["Shortage Quantity (in Ltrs)"], errors="coerce"
        ).fillna(0)
    if "Shortage Age (Days)" in detail_df.columns:
        detail_df["Shortage Age (Days)"] = pd.to_numeric(detail_df["Shortage Age (Days)"], errors="coerce")

    return {
        "total_count" : float(merged[shortage_col].sum()),
        "summary_df"  : summary_df,
        "zone_summary": zone_summary,
        "detail_df"   : detail_df,
        "unmatched"   : [],
    }


def process_open_shortages_sto(
    df_short_sto : pd.DataFrame,
    df_plant     : pd.DataFrame,
    zone_filter  : list = None,
    plant_filter : list = None,
    as_of_date   = None,
) -> dict:
    """
    Process OPEN SHORTAGES - Ltrs (STO) into KPI + drill-down outputs.

    Mapping: Supplying Plant -> PlantMaster Plant Code.
    KPI value: sum of Shortage Quantity (in Ltrs).
    """
    EMPTY = {
        "total_count" : 0.0,
        "summary_df"  : pd.DataFrame(),
        "zone_summary": pd.DataFrame(),
        "detail_df"   : pd.DataFrame(),
        "unmatched"   : [],
    }

    if df_short_sto is None or df_short_sto.empty:
        return EMPTY

    def _pick_col(candidates: list, required: bool = False, label: str = ""):
        for c in candidates:
            if c in df_short_sto.columns:
                return c
        if required:
            st.warning(
                f"⚠️ OPEN SHORTAGES - Ltrs (STO) file is missing expected column for {label}: {candidates}"
            )
        return None

    supp_plant_col = _pick_col(["SUPPLYING PLANT"], required=True, label="Supplying Plant")
    shortage_col = _pick_col(
        ["SHORTAGE QUANTITY (IN LTRS)", "SHORTAGE QUANTITY"],
        required=True,
        label="Shortage quantity (in Ltrs)",
    )
    created_on_col = _pick_col(["CREATED ON"], required=True, label="Created On")

    billing_doc_col  = _pick_col(["BILLING DOCUMENT"])
    shipment_col     = _pick_col(["SHIPMENT NUMBER"])
    plant_col        = _pick_col(["PLANT"])
    service_agent_col= _pick_col(["SERVICE AGENT"])
    sales_org_col    = _pick_col(["SALES ORGANIZATION"])
    delivery_col     = _pick_col(["DELIVERY"])
    vehicle_col      = _pick_col(["VEHICLE"])
    material_col     = _pick_col(["MATERIAL"])
    billed_qty_col   = _pick_col(["BILLED QUANTITY"])
    sales_unit_col   = _pick_col(["SALES UNIT", "SALES UNIT "])
    created_by_col   = _pick_col(["CREATED BY"])

    if not all([supp_plant_col, shortage_col, created_on_col]):
        return EMPTY

    work = df_short_sto.copy()
    work[supp_plant_col] = work[supp_plant_col].astype(str).str.strip()
    work = work[work[supp_plant_col] != ""]

    work[shortage_col] = pd.to_numeric(work[shortage_col], errors="coerce").fillna(0)
    if billed_qty_col and billed_qty_col in work.columns:
        work[billed_qty_col] = pd.to_numeric(work[billed_qty_col], errors="coerce")

    work[created_on_col] = pd.to_datetime(work[created_on_col], errors="coerce", dayfirst=True)

    # As-of-date filter on CREATED ON
    if as_of_date is not None:
        _aod = pd.Timestamp(as_of_date)
        work = work[work[created_on_col].isna() | (work[created_on_col] <= _aod)].copy()

    today = pd.Timestamp(datetime.now().date())
    work["SHORTAGE AGE (DAYS)"] = (today - work[created_on_col]).dt.days
    work.loc[work["SHORTAGE AGE (DAYS)"] < 0, "SHORTAGE AGE (DAYS)"] = pd.NA

    plant_map = (
        df_plant[["Plant Code", "Plant Name", "Zone Name"]]
        .copy()
        .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip()})
    )

    merged = work.merge(
        plant_map,
        left_on  = supp_plant_col,
        right_on = "Plant Code",
        how      = "left",
    )
    merged, _ = _filter_strictly_mapped_rows(merged, supp_plant_col)

    if zone_filter:
        merged = merged[merged["Zone Name"].isin(zone_filter)]
    if plant_filter:
        merged = merged[merged["Plant Name"].isin(plant_filter)]

    if merged.empty:
        return EMPTY

    summary_df = (
        merged
        .groupby(["Zone Name", "Plant Name"], sort=True)
        .agg(shortage_qty=(shortage_col, "sum"))
        .reset_index()
        .rename(columns={"shortage_qty": "Total STO Shortage Quantity (in Ltrs)"})
        .sort_values(["Zone Name", "Total STO Shortage Quantity (in Ltrs)"], ascending=[True, False])
    )

    zone_summary = (
        summary_df
        .groupby("Zone Name")
        .agg(
            Plants      = ("Plant Name", "nunique"),
            shortage_qty = ("Total STO Shortage Quantity (in Ltrs)", "sum"),
        )
        .reset_index()
        .rename(columns={"shortage_qty": "Total STO Shortage Quantity (in Ltrs)"})
        .sort_values("Total STO Shortage Quantity (in Ltrs)", ascending=False)
    )

    rename_map = {
        supp_plant_col   : "Supplying Plant",
        billing_doc_col  : "Billing Document",
        shipment_col     : "Shipment Number",
        plant_col        : "Plant",
        service_agent_col: "Service Agent",
        sales_org_col    : "Sales Organization",
        delivery_col     : "Delivery",
        vehicle_col      : "Vehicle",
        material_col     : "Material",
        billed_qty_col   : "Billed Quantity",
        sales_unit_col   : "Sales Unit",
        shortage_col     : "Shortage Quantity (in Ltrs)",
        created_by_col   : "Created By",
        created_on_col   : "Created On",
        "SHORTAGE AGE (DAYS)": "Shortage Age (Days)",
    }
    detail_df = merged.copy().rename(columns={k: v for k, v in rename_map.items() if k in merged.columns})

    if "Created On" in detail_df.columns:
        detail_df["Created On"] = detail_df["Created On"].dt.strftime("%d-%m-%Y")
        detail_df["Created On"] = detail_df["Created On"].fillna("")

    if "Billed Quantity" in detail_df.columns:
        detail_df["Billed Quantity"] = pd.to_numeric(detail_df["Billed Quantity"], errors="coerce")
    if "Shortage Quantity (in Ltrs)" in detail_df.columns:
        detail_df["Shortage Quantity (in Ltrs)"] = pd.to_numeric(
            detail_df["Shortage Quantity (in Ltrs)"], errors="coerce"
        ).fillna(0)
    if "Shortage Age (Days)" in detail_df.columns:
        detail_df["Shortage Age (Days)"] = pd.to_numeric(detail_df["Shortage Age (Days)"], errors="coerce")

    return {
        "total_count" : float(merged[shortage_col].sum()),
        "summary_df"  : summary_df,
        "zone_summary": zone_summary,
        "detail_df"   : detail_df,
        "unmatched"   : [],
    }


# ─────────────────────────────────────────────────────────────────────────────
# UI HELPER COMPONENTS
# ─────────────────────────────────────────────────────────────────────────────

def render_header(subtitle: str = "") -> None:
    """
    Render the two-part HPCL page header:
      1. Full-width brand banner image  (≈ 2 inches / 192 px tall).
                                 Uses Master Logo.jpg if available, then Title.png, then pure-CSS.
      2. Dark-blue info strip: app title, subtitle, date/time.
    """
    date_badge = f'<span class="banner-date">{subtitle}</span>' if subtitle else ""

    # ── Part 1: Full-width brand banner ─────────────────────────────────────
    title_uri = ""
    logo_uri = ""
    try:
        if os.path.exists(TITLE_IMG_PATH):
            title_uri = _load_img_b64(TITLE_IMG_PATH)
        if os.path.exists(LOGO_IMG_PATH):
            logo_uri = _load_img_b64(LOGO_IMG_PATH)
    except Exception:
        pass

    if title_uri:
        banner_html = f"""
        <div class="hpcl-banner-wrap">
            <img class="hpcl-banner-fg" src="{title_uri}" alt="HPCL SOD Exception Dashboard" />
            {date_badge}
        </div>"""
    elif logo_uri:
        banner_html = f"""
        <div class="hpcl-banner-wrap">
            <img class="hpcl-banner-fg" src="{logo_uri}" alt="HPCL SOD Exception Dashboard" />
            {date_badge}
        </div>"""
    else:
        banner_html = f"""
        <div class="hpcl-banner-wrap" style="
                background:linear-gradient(135deg,#003087 0%,#0057A8 100%);
                height:66px;display:flex;align-items:center;
                padding:0 40px;justify-content:space-between;">
            <div style="font-size:28px;font-weight:900;color:#FFFFFF;
                        letter-spacing:.08em;">&#9981; HPCL</div>
            {date_badge}
        </div>"""

    # ── Part 2: Dark-blue info strip (title only — date is in banner) ────────
    header_html = (
        '<div class="dashboard-header-shell">'
        f'{banner_html}'
        '<div class="dash-header">'
        '<div class="dash-header-main">'
        '<p class="dash-header-title">SOD Exception Dashboard</p>'
        '</div>'
        '</div>'
        '</div>'
    )
    st.markdown(header_html, unsafe_allow_html=True)


def kpi_card(
    label      : str,
    value      : int,
    detail     : str = "",
    icon       : str = "📦",
    color_class: str = "",
    key        : str = None,
) -> bool:
    """Render a KPI tile. Returns True if the View Details button was clicked."""
    formatted   = f"{value:,}" if isinstance(value, (int, float)) else str(value)
    detail_html = f"<div class='kpi-detail'>{detail}</div>" if detail else ""
    st.markdown(f"""
    <div class="kpi-wrap {color_class}">
        <span class="kpi-icon">{icon}</span>
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{formatted}</div>
        {detail_html}
    </div>
    """, unsafe_allow_html=True)
    return st.button("📋 View Details →", key=key or f"btn_{label}",
                     width='stretch')


def render_open_delivery_tile(open_delivery_result: dict) -> bool:
    """Render Open Deliveries KPI tile using the same style as Pending DC."""
    total_open = int(open_delivery_result.get("total_count", 0) or 0)
    color_cls  = "c-success" if total_open > 0 else "c-muted"
    return kpi_card(
        label       = "Open Deliveries",
        value       = total_open,
        detail      = "Count of Open Deliveries",
        icon        = "&#128230;",
        color_class = color_cls,
        key         = "tile_open_del",
    )


def render_open_intransit_tile(open_intransit_result: dict) -> bool:
    """Render Open In-Transit KPI tile with existing card style."""
    total_intransit = int(open_intransit_result.get("total_count", 0) or 0)
    color_cls = "c-success" if total_intransit > 0 else "c-muted"
    return kpi_card(
        label       = "Open In-Transit",
        value       = total_intransit,
        detail      = "Open In-Transit STO Count",
        icon        = "&#128699;",
        color_class = color_cls,
        key         = "tile_intrans",
    )


def render_open_sales_orders_tile(open_sales_orders_result: dict) -> bool:
    """Render Open Sales Orders KPI tile with existing card style."""
    total_so = int(open_sales_orders_result.get("total_count", 0) or 0)
    color_cls = "c-success" if total_so > 0 else "c-muted"
    return kpi_card(
        label       = "Open Sales Orders",
        value       = total_so,
        detail      = "Open Sales Order Count",
        icon        = "&#128203;",
        color_class = color_cls,
        key         = "tile_open_so",
    )


def render_pending_invoices_tile(pending_invoices_result: dict) -> bool:
    """Render Pending Invoices KPI tile with existing card style."""
    total_inv = int(pending_invoices_result.get("total_count", 0) or 0)
    color_cls = "c-success" if total_inv > 0 else "c-muted"
    return kpi_card(
        label       = "Pending Invoices",
        value       = total_inv,
        detail      = "Pending Invoice Count",
        icon        = "&#129534;",
        color_class = color_cls,
        key         = "tile_pend_inv",
    )


def render_tank_reco_tile(tank_reco_result: dict) -> bool:
    """Render Tank Reco KPI tile with existing card style."""
    total_tank = int(tank_reco_result.get("total_count", 0) or 0)
    color_cls = "c-success" if total_tank > 0 else "c-muted"
    return kpi_card(
        label       = "Tank Reco",
        value       = total_tank,
        detail      = "Plant + Tank + Material Count",
        icon        = "&#128738;",
        color_class = color_cls,
        key         = "tile_tank",
    )


def _shortage_color_class(total_short: float) -> str:
    """Return tile color class based on shortage quantity severity."""
    if total_short <= 0:
        return "c-muted"
    if total_short >= 100000:
        return "c-danger"
    if total_short >= 25000:
        return "c-warning"
    return "c-success"


def render_open_shortages_sales_tile(open_short_sales_result: dict) -> bool:
    """Render SHORTAGES - Ltrs (Sales) KPI tile with existing card style."""
    total_short = float(open_short_sales_result.get("total_count", 0) or 0)
    color_cls = _shortage_color_class(total_short)
    display_value = f"{total_short:,.2f}" if abs(total_short - round(total_short)) > 1e-9 else f"{int(round(total_short)):,}"
    return kpi_card(
        label       = "SHORTAGES - Ltrs (Sales)",
        value       = display_value,
        detail      = "Total Shortage Quantity (Ltrs)",
        icon        = "&#128202;",
        color_class = color_cls,
        key         = "tile_sh_sal",
    )


def render_open_shortages_sto_tile(open_short_sto_result: dict) -> bool:
    """Render SHORTAGES - Ltrs (STO) KPI tile with existing card style."""
    total_short = float(open_short_sto_result.get("total_count", 0) or 0)
    color_cls = _shortage_color_class(total_short)
    display_value = f"{total_short:,.2f}" if abs(total_short - round(total_short)) > 1e-9 else f"{int(round(total_short)):,}"
    return kpi_card(
        label       = "SHORTAGES - Ltrs (STO)",
        value       = display_value,
        detail      = "Total STO Shortage Quantity (Ltrs)",
        icon        = "&#128202;",
        color_class = color_cls,
        key         = "tile_sh_sto",
    )


def export_to_excel(df_dict: dict) -> bytes:
    """Serialise {sheet_name: DataFrame} to Excel bytes for st.download_button."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet, df in df_dict.items():
            if df is not None and not df.empty:
                df.to_excel(writer, index=False, sheet_name=sheet[:31])
    buf.seek(0)
    return buf.getvalue()


def _download_excel_button(label: str, file_prefix: str, sheets: dict, key: str) -> None:
    """Render a standardised Excel download button for critical view pages."""
    xlsx_bytes = export_to_excel(sheets)
    st.download_button(
        label=label,
        data=xlsx_bytes,
        file_name=f"{file_prefix}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
    )


def render_professional_summary_table(summary_df: pd.DataFrame) -> None:
    """Render the plant-wise summary as a styled HTML table."""
    table_df = (
        summary_df.rename(
            columns={
                "Zone Name": "Zone",
                "Plant Name": "Plant",
                "Pending DC Count": "Pending DCs",
            }
        )
        .reset_index(drop=True)
    )

    rows_html = "".join(
        "<tr>"
        f"<td>{html.escape(str(row['Zone']))}</td>"
        f"<td>{html.escape(str(row['Plant']))}</td>"
        f"<td>{int(row['Pending DCs'])}</td>"
        "</tr>"
        for _, row in table_df.iterrows()
    )

    table_html = f"""
    <div class="pro-table-wrap">
        <table class="pro-table">
            <thead>
                <tr>
                    <th>Zone</th>
                    <th>Plant</th>
                    <th>Pending DCs</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

def _render_sidebar_system_info(
    system_info_slot,
    df_plant: pd.DataFrame,
    all_exception_plant_df: pd.DataFrame = None,
    exception_kpi_df: pd.DataFrame = None,
) -> None:
    """Render live sidebar system info with all-KPI status and mapped coverage."""
    master_plants = int(len(df_plant)) if df_plant is not None else 0
    master_zones = int(df_plant["Zone Name"].nunique()) if df_plant is not None and "Zone Name" in df_plant.columns else 0

    data_plants = 0
    data_zones = 0
    total_exceptions = 0
    if all_exception_plant_df is not None and not all_exception_plant_df.empty:
        if "Plant Name" in all_exception_plant_df.columns:
            data_plants = int(all_exception_plant_df["Plant Name"].nunique())
        if "Zone Name" in all_exception_plant_df.columns:
            data_zones = int(all_exception_plant_df["Zone Name"].nunique())
        if "Total Exceptions" in all_exception_plant_df.columns:
            total_exceptions = int(pd.to_numeric(all_exception_plant_df["Total Exceptions"], errors="coerce").fillna(0).sum())

    kpi_rows_html = ""
    if exception_kpi_df is not None and not exception_kpi_df.empty:
        for _, row in exception_kpi_df.sort_values("KPI Value", ascending=False).iterrows():
            label = html.escape(str(row.get("Exception KPI", "")))
            val = int(round(float(row.get("KPI Value", 0) or 0)))
            kpi_rows_html += (
                f'<div style="display:flex;justify-content:space-between;gap:8px;line-height:1.5;">'
                f'<span style="opacity:.92;">{label}</span><b>{val:,}</b></div>'
            )
    else:
        kpi_rows_html = '<div style="opacity:.75;">KPI module totals will appear after data load.</div>'

    info_html = f"""
    <div style="font-size:13.5px;line-height:1.9;opacity:.94;">
        &#128200; &nbsp;Total Exceptions (All KPI): <b>{total_exceptions:,}</b><br/>
        &#127981; &nbsp;Mapped Plants in Data: <b>{data_plants}</b><br/>
        &#128506; &nbsp;Mapped Zones in Data: <b>{data_zones}</b><br/>
        &#127970; &nbsp;PlantMaster Plants: <b>{master_plants}</b><br/>
        &#128205; &nbsp;PlantMaster Zones: <b>{master_zones}</b><br/>
        &#128197; &nbsp;Date: <b>{datetime.now().strftime('%d %b %Y')}</b>
    </div>
    <div style="margin-top:8px;padding:8px;border:1px solid rgba(255,255,255,.20);border-radius:8px;background:rgba(255,255,255,.04);max-height:180px;overflow:auto;">
        <div style="font-size:12px;font-weight:700;opacity:.90;margin-bottom:4px;">All KPI Module Counts</div>
        {kpi_rows_html}
    </div>
    """
    system_info_slot.markdown(info_html, unsafe_allow_html=True)

def render_sidebar(df_plant: pd.DataFrame) -> tuple:
    """Render navigation sidebar. Returns (zones, plants, uploaded_file, system_info_slot, as_of_date)."""
    with st.sidebar:
        sidebar_logo_html = '<div style="font-size:2.6rem;">&#9981;</div>'
        try:
            if os.path.exists(SIDE_PANEL_LOGO_PATH):
                logo_uri = _load_img_b64(SIDE_PANEL_LOGO_PATH)
                sidebar_logo_html = (
                    f'<img src="{logo_uri}" alt="Side Panel Logo" '
                    'style="width:100%;height:auto;display:block;margin:0 auto 6px auto;'
                    'object-fit:contain;" />'
                )
            elif os.path.exists(LOGO_IMG_PATH):
                logo_uri = _load_img_b64(LOGO_IMG_PATH)
                sidebar_logo_html = (
                    f'<img src="{logo_uri}" alt="HPCL Corporate Logo" '
                    'style="height:52px;width:auto;display:block;margin:0 auto 6px auto;'
                    'object-fit:contain;" />'
                )
        except Exception:
            pass

        st.markdown(f"""
        <div class="sidebar-branding">
            {sidebar_logo_html}
            <div style="font-size:1.2rem;font-weight:700;letter-spacing:.06em;color:#FFFFFF;">HPCL</div>
            <div style="font-size:0.75rem;opacity:.7;color:#AACCFF;">Exception Monitoring</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<hr/>", unsafe_allow_html=True)

        _sb_lv_drill = st.session_state.get("location_visit_page") == "drilldown"

        if not _sb_lv_drill:
            st.markdown('<p class="sb-nav-lbl">&#128205; Navigation Filters</p>', unsafe_allow_html=True)

            all_zones = sorted(df_plant["Zone Name"].dropna().unique().tolist())
            selected_zones = st.multiselect(
                "Zone",
                options=all_zones,
                default=[],
                placeholder="All Zones",
            )

            if selected_zones:
                avail_plants = sorted(
                    df_plant[df_plant["Zone Name"].isin(selected_zones)]
                    ["Plant Name"].dropna().unique().tolist()
                )
            else:
                avail_plants = sorted(df_plant["Plant Name"].dropna().unique().tolist())

            selected_plants = st.multiselect(
                "Plant / Location",
                options=sorted(avail_plants),
                default=[],
                placeholder="All Plants",
            )

            st.markdown("<hr/>", unsafe_allow_html=True)
        else:
            all_zones = sorted(df_plant["Zone Name"].dropna().unique().tolist())
            selected_zones = []
            avail_plants = sorted(df_plant["Plant Name"].dropna().unique().tolist())
            selected_plants = []

        st.markdown('<p class="sb-nav-lbl">&#128197; Data As-Of Date</p>', unsafe_allow_html=True)
        from datetime import timedelta as _td
        _yesterday = datetime.now().date() - _td(days=1)
        # Initialise session state on first load so the date persists across reruns
        if "as_of_date_picker" not in st.session_state:
            st.session_state["as_of_date_picker"] = _yesterday
        as_of_date = st.date_input(
            "Show data up to:",
            value=st.session_state.get("as_of_date_picker", _yesterday),
            max_value=datetime.now().date(),
            format="DD/MM/YYYY",
            key="as_of_date_picker",
            help="KPI counts will include only records on or before this date.",
        )
        st.caption(f"Filtering data up to: **{as_of_date.strftime('%d %b %Y')}**")

        st.markdown("<hr/>", unsafe_allow_html=True)

        # Use Streamlit's default sidebar collapse/expand button for robust functionality
        # No custom restore button injected

        st.markdown('<p class="sb-nav-lbl">&#8505;&#65039; System Info</p>', unsafe_allow_html=True)
        system_info_slot = st.empty()
        _render_sidebar_system_info(
            system_info_slot,
            df_plant,
            all_exception_plant_df=pd.DataFrame(),
            exception_kpi_df=pd.DataFrame(),
        )

        st.markdown("<hr/>", unsafe_allow_html=True)
        import uuid
        unique_refresh_key = f"btn_refresh_{uuid.uuid4()}"
        if st.button("&#128260; Refresh Data", width='stretch', key=unique_refresh_key):
            st.cache_data.clear()
            st.rerun()

        st.markdown("<hr/>", unsafe_allow_html=True)

        if st.button("&#9993; Send Exception Mails", width='stretch',
                     key="btn_open_mail_center"):
            st.session_state["open_mail_center"] = True
            st.rerun()

        st.markdown("<hr/>", unsafe_allow_html=True)

        st.markdown('<p class="sb-nav-lbl">&#128194; Data Upload</p>',
                    unsafe_allow_html=True)
        unique_uploader_key = f"uploader_pending_dc_{uuid.uuid4()}"
        uploaded_dc = st.file_uploader(
            "Pending DC File  (.xls / .xlsx)",
            type = ["xls", "xlsx"],
            key  = unique_uploader_key,
        )

    return selected_zones, selected_plants, uploaded_dc, system_info_slot, as_of_date


# ─────────────────────────────────────────────────────────────────────────────
# SHARED HTML TABLE RENDERER
# ─────────────────────────────────────────────────────────────────────────────

_TABLE_INSTANCE_COUNTERS = {}


def _next_table_instance_key(base_key: str) -> str:
    """Return a deterministic unique key per rendered table instance in a script run."""
    next_idx = _TABLE_INSTANCE_COUNTERS.get(base_key, 0) + 1
    _TABLE_INSTANCE_COUNTERS[base_key] = next_idx
    return f"{base_key}_{next_idx}"

def _render_html_table(df: pd.DataFrame, col_labels: dict = None, max_height: int = 500) -> None:
    """Render a DataFrame as a styled pro-table HTML element."""
    if df is None or df.empty:
        st.info("No data to display.")
        return
    display_df = df.copy()
    if col_labels:
        display_df = display_df.rename(columns=col_labels)
    caller_frame = inspect.stack()[1]
    
    # Create a hash of the dataframe content to ensure unique keys for different data
    df_content_hash = hashlib.md5(pd.util.hash_pandas_object(df, index=True).values.tobytes()).hexdigest()[:8]
    
    key_seed = "|".join(
        [
            str(caller_frame.function),
            str(caller_frame.lineno),
            df_content_hash,
            *(str(col) for col in display_df.columns),
        ]
    )
    base_table_key = f"table_view_{hashlib.md5(key_seed.encode('utf-8')).hexdigest()[:12]}"
    table_key = base_table_key
    is_maximized = st.session_state.get(f"{table_key}_maximized", False)

    controls_col1, controls_col2 = st.columns([1.2, 6])
    with controls_col1:
        toggle_label = "🗗 Restore" if is_maximized else "⛶ Maximize"
        if st.button(toggle_label, key=f"{table_key}_toggle", use_container_width=True):
            st.session_state[f"{table_key}_maximized"] = not is_maximized
            st.rerun()

    # Professional Streamlit-styled table: light bluish header, centered data
    is_drilldown_view = (
        st.session_state.get("page", "dashboard") != "dashboard"
        or st.session_state.get("selected_tile") == "location_visit"
        or st.session_state.get("location_visit_page") == "drilldown"
        or st.session_state.get("dummy_tank_clicked") is True
        or st.session_state.get("pl_unblock_clicked") is True
    )
    font_boost = -2 if is_drilldown_view else 0
    header_font = 16 + font_boost
    cell_font = 15 + font_boost
    effective_height = 1100 if is_maximized else max_height
    header_bg = "#eaf2fb"  # light bluish
    header_color = "#003087"  # deep blue for text
    cell_bg_odd = "#ffffff"
    cell_bg_even = "#f7fafd"
    headers_html = "".join(
        f"<th style='background:{header_bg};color:{header_color};font-weight:700;text-align:center;padding:10px 8px;font-size:{header_font}px;border-bottom:2px solid #d5e2f3;'>{html.escape(str(c))}</th>" for c in display_df.columns
    )
    rows_html = "".join(
        f"<tr style='background:{cell_bg_odd if i%2==0 else cell_bg_even};'>"
        + "".join(
            f"<td style='text-align:center;padding:8px 6px;font-size:{cell_font}px;border-bottom:1px solid #e2eaf4;'>{html.escape(str(v) if pd.notna(v) else '')}</td>"
            for v in row
        )
        + "</tr>"
        for i, (_, row) in enumerate(display_df.iterrows())
    )
    st.markdown(
        f'<div class="pro-table-wrap" style="max-height:{effective_height}px;overflow:auto;border-radius:8px;border:1px solid #d5e2f3;background:#fff;box-shadow:0 2px 8px #e0e0e0;">'
        f'<table class="pro-table" style="width:100%;border-collapse:collapse;font-size:{cell_font}px;">'
        f'<thead><tr>{headers_html}</tr></thead>'
        f'<tbody>{rows_html}</tbody>'
        f'</table></div>',
        unsafe_allow_html=True,
    )


def _render_streamlit_dataframe(
    df: pd.DataFrame,
    max_height: int = 420,
    hide_index: bool = True,
    use_container_width: bool = True,
) -> None:
    """Render a native Streamlit dataframe with a shared maximize toggle."""
    if df is None or df.empty:
        st.info("No data to display.")
        return

    caller_frame = inspect.stack()[1]
    
    # Create a hash of the dataframe content to ensure unique keys for different data
    df_content_hash = hashlib.md5(pd.util.hash_pandas_object(df, index=True).values.tobytes()).hexdigest()[:8]
    
    key_seed = "|".join(
        [
            str(caller_frame.function),
            str(caller_frame.lineno),
            df_content_hash,
            *(str(col) for col in df.columns),
        ]
    )
    base_table_key = f"stdf_view_{hashlib.md5(key_seed.encode('utf-8')).hexdigest()[:12]}"
    table_key = base_table_key
    is_maximized = st.session_state.get(f"{table_key}_maximized", False)

    controls_col1, controls_col2 = st.columns([1.2, 6])
    with controls_col1:
        toggle_label = "🗗 Restore" if is_maximized else "⛶ Maximize"
        if st.button(toggle_label, key=f"{table_key}_toggle", use_container_width=True):
            st.session_state[f"{table_key}_maximized"] = not is_maximized
            st.rerun()

    effective_height = 1100 if is_maximized else max_height
    st.dataframe(
        df,
        use_container_width=use_container_width,
        hide_index=hide_index,
        height=effective_height,
    )


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: MAIN DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────

def render_dashboard(
    df_plant: pd.DataFrame,
    pending_dc_result  : dict,
    open_delivery_result: dict,
    open_intransit_result: dict,
    open_sales_orders_result: dict,
    pending_invoices_result: dict,
    tank_reco_result   : dict,
    open_short_sales_result: dict,
    open_short_sto_result: dict,
    all_exception_plant_df: pd.DataFrame,
    zone_exception_summary_df: pd.DataFrame,
    zone_filter        : list,
    plant_filter       : list,
    as_of_date         = None,
) -> None:
    """Main dashboard page with KPI tiles, zone chart, and plant table."""

    _aod_label = as_of_date.strftime("%d %b %Y") if as_of_date else ""
    render_header(subtitle=f"Data as of: {_aod_label}" if _aod_label else "")
    # ── Navigation Filters (Zone & Plant) — hidden when on Location Visit drill-down ──
    _is_lv_drilldown = st.session_state.get("location_visit_page") == "drilldown"
    if not _is_lv_drilldown:
        st.markdown('<div class="sec-title">&#128205; Navigation Filters</div>', unsafe_allow_html=True)
        filters_col1, filters_col2 = st.columns([1, 1])
        zone_list = sorted(df_plant["Zone Name"].dropna().unique().tolist())
        plant_list = sorted(df_plant["Plant Name"].dropna().unique().tolist())
        with filters_col1:
            selected_zone = st.selectbox("Select Zone", ["All Zones"] + zone_list, key="zone_filter")
        with filters_col2:
            if selected_zone != "All Zones":
                filtered_plants = sorted(df_plant[df_plant["Zone Name"] == selected_zone]["Plant Name"].dropna().unique().tolist())
            else:
                filtered_plants = plant_list
            selected_plant = st.selectbox("Select Plant / Location", ["All Plants"] + filtered_plants, key="plant_filter")
    else:
        selected_zone = "All Zones"
        selected_plant = "All Plants"

    # --- LOCATION VISIT KPI LOGIC (after navigation filters) ---
    df_loc_filtered = pd.DataFrame()
    loc_visit_missing_columns = []
    loc_visit_error = ""
    kpi_location_visit = 0
    kpi_location_compliance = 0.0

    if os.path.exists(LOCATION_VISIT_PATH):
        df_loc = load_location_visit(
            LOCATION_VISIT_PATH,
            cache_buster=_get_file_cache_token(LOCATION_VISIT_PATH),
        )
        required_loc_cols = [
            "Planning Plant", "Plant Desc.", "Audit Number", "Audit Start Date", "Audit End Date",
            "TotalRecomms", "ClosedRecomms", "OpenRecomms",
        ]
        loc_visit_missing_columns = [c for c in required_loc_cols if c not in df_loc.columns]

        if loc_visit_missing_columns:
            found_cols = ", ".join(map(str, list(df_loc.columns)[:8])) if not df_loc.empty else "no readable columns found"
            loc_visit_error = (
                "Location Visit file format is not recognized. Expected columns like "
                "Planning Plant, Plant Desc., Audit Number, Audit Start Date, Audit End Date, "
                "TotalRecomms, ClosedRecomms and OpenRecomms. "
                f"Found: {found_cols}"
            )
        else:
            df_loc_work = df_loc.copy()
            df_loc_work["Planning Plant"] = (
                df_loc_work["Planning Plant"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            )

            plant_map = (
                df_plant[["Plant Code", "Plant Name", "Zone Name"]]
                .copy()
                .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)})
            )
            df_loc_work = df_loc_work.merge(
                plant_map,
                left_on="Planning Plant",
                right_on="Plant Code",
                how="left",
            )
            # Zone always comes from PlantMaster (Zone Name column)
            df_loc_work["Zone"] = (
                df_loc_work["Zone Name"]
                .fillna("Unmapped")
                .astype(str)
                .str.strip()
            )
            df_loc_work["Zone Name"] = df_loc_work["Zone"]
            df_loc_work["Plant Name"] = (
                df_loc_work["Plant Name"]
                .fillna(df_loc_work["Plant Desc."].astype(str).str.strip())
            )

            # Save all-zones copy for mail (before nav filter)
            _loc_for_mail = df_loc_work.copy()

            if selected_zone != "All Zones":
                df_loc_work = df_loc_work[df_loc_work["Zone Name"] == selected_zone]
            if selected_plant != "All Plants":
                df_loc_work = df_loc_work[df_loc_work["Plant Name"] == selected_plant]

            # Deduplicate: 1 record per Plant+FY+Quarter (latest audit) — matches Sr. Manager Inspection Dashboard
            def _get_fy_quarter(dt):
                if pd.isna(dt): return ("Unknown", "Unknown")
                m, y = dt.month, dt.year
                if m >= 4:
                    return (str(y)[2:] + "-" + str(y + 1)[2:], "Q" + str(((m - 4) // 3) + 1))
                return (str(y - 1)[2:] + "-" + str(y)[2:], "Q4")

            _dates = pd.to_datetime(
                df_loc_work["Audit Start Date"], errors="coerce", dayfirst=True
            )
            _fq = _dates.apply(_get_fy_quarter)
            df_loc_work["_FY"]       = _fq.apply(lambda x: x[0])
            df_loc_work["_Quarter"]  = _fq.apply(lambda x: x[1])
            df_loc_work["_date_sort"] = _dates
            # Dedup: per Plant+FY+Quarter keep the LATEST audit (highest date, then highest Audit Number)
            # Matches Sr. Manager Inspection Dashboard's dedupByLatestAudit logic
            df_loc_work = (
                df_loc_work
                .sort_values(
                    ["_date_sort", "Audit Number"],
                    ascending=[False, False],
                    na_position="last",
                )
                .drop_duplicates(subset=["Planning Plant", "_FY", "_Quarter"], keep="first")
                .rename(columns={"_FY": "FY", "_Quarter": "Quarter"})
                .drop(columns=["_date_sort"])
                .reset_index(drop=True)
            )

            df_loc_filtered = df_loc_work.copy()

            # Tile: default to previous quarter data
            _tnow = datetime.now()
            _tm, _ty = _tnow.month, _tnow.year
            if _tm >= 4:
                _curr_fy_t = str(_ty)[2:] + "-" + str(_ty+1)[2:]
                _curr_q_t  = ((_tm - 4) // 3) + 1
            else:
                _curr_fy_t = str(_ty-1)[2:] + "-" + str(_ty)[2:]
                _curr_q_t  = 4
            if _curr_q_t == 1:
                _prev_fy_t = str(int("20" + _curr_fy_t[:2]) - 1)[2:] + "-" + _curr_fy_t[:2]
                _prev_q_t  = "Q4"
            else:
                _prev_fy_t = _curr_fy_t
                _prev_q_t  = f"Q{_curr_q_t - 1}"
            _prev_qtr_label = f"{_prev_q_t} FY{_prev_fy_t}"

            _df_tile = df_loc_filtered[
                (df_loc_filtered["FY"] == _prev_fy_t) &
                (df_loc_filtered["Quarter"] == _prev_q_t)
            ]
            if _df_tile.empty:
                _df_tile = df_loc_filtered  # fallback: all data if no prev-qtr records

            _loc_total  = pd.to_numeric(_df_tile["TotalRecomms"],  errors="coerce").fillna(0).sum()
            _loc_closed = pd.to_numeric(_df_tile["ClosedRecomms"], errors="coerce").fillna(0).sum()
            kpi_location_visit      = len(_df_tile)   # audit record count (matches Sr. Mgr Dashboard)
            kpi_location_compliance = (_loc_closed / _loc_total) if _loc_total > 0 else 0.0
    else:
        loc_visit_error = "Location Visit file not found at Reports/LOCATION_VISIT.xls"

    # Dummy Tank KPI data prep
    dummy_tank_filtered = pd.DataFrame()
    dummy_tank_missing_columns = []
    dummy_tank_error = ""
    total_dummy_qty = 0.0

    if os.path.exists(DUMMY_TANK_PATH):
        df_dummy_tank = load_dummy_tank_stock(
            DUMMY_TANK_PATH,
            cache_buster=_get_file_cache_token(DUMMY_TANK_PATH),
        )
        required_dummy_cols = ["Plant", "Material", "Storage Location", "Base Unit of Measure", "Unrestricted"]
        dummy_tank_missing_columns = [c for c in required_dummy_cols if c not in df_dummy_tank.columns]

        if not dummy_tank_missing_columns:
            dummy_tank_filtered = df_dummy_tank[required_dummy_cols].copy()

            # Normalize plant codes to align Dummy Tank report with PlantMaster codes.
            dummy_tank_filtered["Plant"] = (
                dummy_tank_filtered["Plant"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            )

            plant_map = (
                df_plant[["Plant Code", "Plant Name", "Zone Name"]]
                .copy()
                .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)})
            )

            dummy_tank_filtered = dummy_tank_filtered.merge(
                plant_map,
                left_on="Plant",
                right_on="Plant Code",
                how="left",
            )

            dummy_tank_filtered = dummy_tank_filtered[
                ~dummy_tank_filtered["Storage Location"].isin(["DBIT", "DLUB", "DSLP"])
            ]

            # Save all-zones copy for mail (before nav filter)
            _dummy_for_mail = dummy_tank_filtered.copy()

            if selected_zone != "All Zones":
                dummy_tank_filtered = dummy_tank_filtered[dummy_tank_filtered["Zone Name"] == selected_zone]
            if selected_plant != "All Plants":
                dummy_tank_filtered = dummy_tank_filtered[dummy_tank_filtered["Plant Name"] == selected_plant]

            dummy_tank_filtered["Zone"] = dummy_tank_filtered["Zone Name"]
            dummy_tank_filtered["Unrestricted"] = pd.to_numeric(
                dummy_tank_filtered["Unrestricted"],
                errors="coerce",
            ).fillna(0)
            total_dummy_qty = float(dummy_tank_filtered["Unrestricted"].sum())
    else:
        dummy_tank_error = "Dummy Tank file not found at Reports/DUMMY TANK STOCK.xls"

    # PL Unblock KPI data prep
    pl_unblock_filtered = pd.DataFrame()
    pl_unblock_missing_columns = []
    pl_unblock_error = ""
    total_pl_unblock_qty = 0.0

    if os.path.exists(PIPELINE_STOCK_PATH):
        df_pipeline_stock = load_pipeline_stock(
            PIPELINE_STOCK_PATH,
            cache_buster=_get_file_cache_token(PIPELINE_STOCK_PATH),
        )
        required_pl_cols = ["Material", "Plant", "Storage location", "Base Unit of Measure", "Unrestricted", "Blocked"]
        pl_unblock_missing_columns = [c for c in required_pl_cols if c not in df_pipeline_stock.columns]

        if not pl_unblock_missing_columns:
            pl_unblock_filtered = df_pipeline_stock[required_pl_cols].copy()

            pl_unblock_filtered["Plant"] = (
                pl_unblock_filtered["Plant"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            )

            plant_map = (
                df_plant[["Plant Code", "Plant Name", "Zone Name"]]
                .copy()
                .assign(**{"Plant Code": lambda d: d["Plant Code"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)})
            )

            pl_unblock_filtered = pl_unblock_filtered.merge(
                plant_map,
                left_on="Plant",
                right_on="Plant Code",
                how="left",
            )

            # Save all-zones copy for mail (before nav filter)
            _pl_for_mail = pl_unblock_filtered.copy()

            if selected_zone != "All Zones":
                pl_unblock_filtered = pl_unblock_filtered[pl_unblock_filtered["Zone Name"] == selected_zone]
            if selected_plant != "All Plants":
                pl_unblock_filtered = pl_unblock_filtered[pl_unblock_filtered["Plant Name"] == selected_plant]

            pl_unblock_filtered["Zone"] = pl_unblock_filtered["Zone Name"]
            pl_unblock_filtered["Unrestricted"] = pd.to_numeric(
                pl_unblock_filtered["Unrestricted"],
                errors="coerce",
            ).fillna(0)
            pl_unblock_filtered["Blocked"] = pd.to_numeric(
                pl_unblock_filtered["Blocked"],
                errors="coerce",
            ).fillna(0)
            total_pl_unblock_qty = float(pl_unblock_filtered["Unrestricted"].sum())
    else:
        pl_unblock_error = "PL Unblock file not found at Reports/PIPELINE STOCK.xls"

    # Tank Turns KPI data prep
    tank_turns_df      = pd.DataFrame()
    tank_turns_error   = ""
    tank_turns_missing = []
    tank_turns_value   = 0.0

    if os.path.exists(TANK_TURNS_PATH):
        df_tank_turns_raw = load_tank_turns(
            TANK_TURNS_PATH,
            cache_buster=_get_file_cache_token(TANK_TURNS_PATH),
        )
        required_tt_cols  = ["Plant", "Zone", "Plant Name", "Tank", "Unique Ref Id",
                             "Material", "Material Description", "Tank Capacity",
                             "Dispatches", "Turn", "Tank Type", "Tank Status",
                             "Opening Stock", "Receipts", "Closing Stock"]
        tank_turns_missing = [c for c in required_tt_cols if c not in df_tank_turns_raw.columns]

        if not tank_turns_missing:
            # Filter to valid PlantMaster plants only
            valid_plants_set = set(
                df_plant["Plant Code"].astype(str).str.strip()
                .str.replace(r"\.0$", "", regex=True).unique()
            )
            df_tt = df_tank_turns_raw.copy()
            df_tt["Plant"] = df_tt["Plant"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            df_tt = df_tt[df_tt["Plant"].isin(valid_plants_set)]

            # Merge PlantMaster for Zone Name / Plant Name so nav filters work
            plant_map_tt = (
                df_plant[["Plant Code", "Plant Name", "Zone Name"]]
                .copy()
                .assign(**{"Plant Code": lambda d:
                           d["Plant Code"].astype(str).str.strip()
                           .str.replace(r"\.0$", "", regex=True)})
            )
            df_tt = df_tt.merge(plant_map_tt, left_on="Plant",
                                right_on="Plant Code", how="left",
                                suffixes=("", "_master"))

            # Save all-zones copy for mail (before nav filter)
            _tt_for_mail = df_tt.copy()

            # Apply navigation filters
            if selected_zone != "All Zones":
                df_tt = df_tt[df_tt["Zone Name"] == selected_zone]
            if selected_plant != "All Plants":
                df_tt = df_tt[df_tt["Plant Name_master"].fillna(df_tt.get("Plant Name", "")) == selected_plant]

            df_tt["Dispatches"]    = pd.to_numeric(df_tt["Dispatches"],    errors="coerce").fillna(0)
            df_tt["Tank Capacity"] = pd.to_numeric(df_tt["Tank Capacity"], errors="coerce").fillna(0)

            total_dispatch  = df_tt["Dispatches"].sum()
            total_capacity  = df_tt["Tank Capacity"].sum()
            tank_turns_value = (total_dispatch / total_capacity) if total_capacity != 0 else 0.0
            tank_turns_df    = df_tt
    else:
        tank_turns_error = "Tank Turns file not found at Reports/Tank Turn.xlsx"

    # Render drill-downs as standalone pages (hide dashboard tiles/charts while open)
    if st.session_state.get("location_visit_page") == "drilldown":
        if loc_visit_error:
            st.warning(loc_visit_error)
        elif loc_visit_missing_columns:
            st.warning("Missing required column(s): " + ", ".join(loc_visit_missing_columns))
        else:
            render_location_visit_details(df_loc_filtered)
        return

    if st.session_state.get("dummy_tank_clicked") is True:
        render_dummy_tank_details(
            dummy_tank_filtered,
            total_dummy_qty,
            error_message=dummy_tank_error,
            missing_columns=dummy_tank_missing_columns,
        )
        return

    if st.session_state.get("pl_unblock_clicked") is True:
        render_pl_unblock_details(
            pl_unblock_filtered,
            total_pl_unblock_qty,
            error_message=pl_unblock_error,
            missing_columns=pl_unblock_missing_columns,
        )
        return

    if st.session_state.get("tank_turns_page") == "drilldown":
        render_tank_turns_details(
            tank_turns_df,
            tank_turns_value,
            error_message=tank_turns_error,
            missing_columns=tank_turns_missing,
        )
        return

    # Active filter badges
    if zone_filter or plant_filter:
        badges = "".join(
            [f'<span class="fbadge">&#128205; {z}</span>' for z in zone_filter]
            + [f'<span class="fbadge">&#127981; {p}</span>' for p in plant_filter]
        )
        st.markdown(
            f'<div style="margin-bottom:12px;font-size:15px;">'
            f'Active Filters:&nbsp;{badges}</div>',
            unsafe_allow_html=True,
        )

    # Apply filtering to all dashboard data
    def filter_df(df):
        if selected_zone != "All Zones" and selected_plant == "All Plants":
            if "Zone Name" in df.columns:
                return df[df["Zone Name"] == selected_zone]
            else:
                return df
        elif selected_plant != "All Plants":
            if "Plant Name" in df.columns:
                return df[df["Plant Name"] == selected_plant]
            else:
                return df
        else:
            return df

    # Filter all relevant dataframes before KPI tiles
    pending_dc_result_filtered = pending_dc_result.copy()
    open_delivery_result_filtered = open_delivery_result.copy()
    open_intransit_result_filtered = open_intransit_result.copy()
    open_sales_orders_result_filtered = open_sales_orders_result.copy()
    pending_invoices_result_filtered = pending_invoices_result.copy()
    tank_reco_result_filtered = tank_reco_result.copy()
    open_short_sales_result_filtered = open_short_sales_result.copy()
    open_short_sto_result_filtered = open_short_sto_result.copy()

    # Filter summary and detail DataFrames
    for result in [pending_dc_result_filtered, open_delivery_result_filtered, open_intransit_result_filtered,
                  open_sales_orders_result_filtered, pending_invoices_result_filtered, tank_reco_result_filtered,
                  open_short_sales_result_filtered, open_short_sto_result_filtered]:
        if "summary_df" in result:
            result["summary_df"] = filter_df(result["summary_df"])
        if "zone_summary" in result:
            result["zone_summary"] = filter_df(result["zone_summary"])
        if "detail_df" in result:
            result["detail_df"] = filter_df(result["detail_df"])

    _aod_disp = as_of_date.strftime("%d %b %Y") if as_of_date else "all dates"
    st.markdown(
        f'<div class="sec-title">&#128202; Exception Parameters &#8212; Live Summary'
        f'&nbsp;&nbsp;<span style="font-size:12px;font-weight:400;color:#888;">'
        f'&#128197; Data filtered up to: <b>{_aod_disp}</b></span></div>',
        unsafe_allow_html=True,
    )

    col1, col2, col3, col4 = st.columns(4, gap="small")

    with col1:
        s_df = pending_dc_result_filtered.get("summary_df", pd.DataFrame())
        z_df = pending_dc_result_filtered.get("zone_summary", pd.DataFrame())
        # Recalculate total_dc from filtered summary_df
        total_dc = int(s_df["Pending DC Count"].sum()) if not s_df.empty and "Pending DC Count" in s_df.columns else 0
        detail_str = f"{len(z_df)} zones  |  {len(s_df)} plants affected"
        color_cls = "c-danger" if total_dc > 50 else ("c-warning" if total_dc > 20 else "")
        clicked_dc = kpi_card(
            label = "Pending DC's",
            value = total_dc,
            detail = detail_str,
            icon = "&#128666;",
            color_class = color_cls,
            key = "tile_pending_dc",
        )
        if clicked_dc:
            st.session_state["page"] = "pending_dc_details"
            st.rerun()

    with col2:
        s_df = open_delivery_result_filtered.get("summary_df", pd.DataFrame())
        total_deliveries = int(s_df["Open Delivery Count"].sum()) if not s_df.empty and "Open Delivery Count" in s_df.columns else 0
        clicked_open = render_open_delivery_tile({**open_delivery_result_filtered, "total_count": total_deliveries})
        if clicked_open:
            st.session_state["page"] = "open_delivery_details"
            st.rerun()
    with col3:
        s_df = pending_invoices_result_filtered.get("summary_df", pd.DataFrame())
        total_invoices = int(s_df["Pending Invoice Count"].sum()) if not s_df.empty and "Pending Invoice Count" in s_df.columns else 0
        clicked_pending_inv = render_pending_invoices_tile({**pending_invoices_result_filtered, "total_count": total_invoices})
        if clicked_pending_inv:
            st.session_state["page"] = "pending_invoices_details"
            st.rerun()
    with col4:
        s_df = open_sales_orders_result_filtered.get("summary_df", pd.DataFrame())
        total_so = int(s_df["Open Sales Order Count"].sum()) if not s_df.empty and "Open Sales Order Count" in s_df.columns else 0
        clicked_open_so = render_open_sales_orders_tile({**open_sales_orders_result_filtered, "total_count": total_so})
        if clicked_open_so:
            st.session_state["page"] = "open_sales_orders_details"
            st.rerun()

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

    # ── Row 2: KPI tiles ─────────────────────────────────────────────────────
    col5, col6, col7, col8 = st.columns(4, gap="small")
    with col5:
        s_df = open_intransit_result_filtered.get("summary_df", pd.DataFrame())
        total_intransit = int(s_df["Open In-Transit STO Count"].sum()) if not s_df.empty and "Open In-Transit STO Count" in s_df.columns else 0
        clicked_intransit = render_open_intransit_tile({**open_intransit_result_filtered, "total_count": total_intransit})
        if clicked_intransit:
            st.session_state["page"] = "open_intransit_details"
            st.rerun()
    with col6:
        s_df = open_short_sales_result_filtered.get("summary_df", pd.DataFrame())
        total_short_sales = int(s_df["Total Shortage Quantity (in Ltrs)"].sum()) if not s_df.empty and "Total Shortage Quantity (in Ltrs)" in s_df.columns else 0
        clicked_short_sales = render_open_shortages_sales_tile({**open_short_sales_result_filtered, "total_count": total_short_sales})
        if clicked_short_sales:
            st.session_state["page"] = "open_shortages_sales_details"
            st.rerun()
    with col7:
        s_df = open_short_sto_result_filtered.get("summary_df", pd.DataFrame())
        total_short_sto = int(s_df["Total STO Shortage Quantity (in Ltrs)"].sum()) if not s_df.empty and "Total STO Shortage Quantity (in Ltrs)" in s_df.columns else 0
        clicked_short_sto = render_open_shortages_sto_tile({**open_short_sto_result_filtered, "total_count": total_short_sto})
        if clicked_short_sto:
            st.session_state["page"] = "open_shortages_sto_details"
            st.rerun()
    with col8:
        s_df = tank_reco_result_filtered.get("summary_df", pd.DataFrame())
        total_tank_reco = int(s_df["Tank Reco Count"].sum()) if not s_df.empty and "Tank Reco Count" in s_df.columns else 0
        clicked_tank = render_tank_reco_tile({**tank_reco_result_filtered, "total_count": total_tank_reco})
        if clicked_tank:
            st.session_state["page"] = "tank_reco_details"
            st.rerun()

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # ── Row 3: New KPI tiles (Coming Soon) ───────────────────────────────
    col9, col10, col11, col12 = st.columns(4, gap="small")
    with col9:
        clicked_pl_unblock = kpi_card(
            label="PL Unblock Qty (KL)",
            value=f"{total_pl_unblock_qty / 1000:,.3f}",
            detail="Total Pipeline Unblock Quantity",
            icon="&#128295;",
            color_class="c-success" if total_pl_unblock_qty > 0 else "c-muted",
            key="pl_unblock_btn",
        )
        if clicked_pl_unblock:
            st.session_state["dummy_tank_clicked"] = False
            st.session_state["selected_tile"] = None
            st.session_state["tank_turns_page"] = "main"
            st.session_state["location_visit_page"] = "main"
            st.session_state["pl_unblock_clicked"] = True
            st.rerun()
    with col10:
        clicked_dummy_tank = kpi_card(
            label="DUMMY TANK QTY. (KL)",
            value=f"{total_dummy_qty / 1000:.3f}",
            detail="Total Dummy Tank Quantity",
            icon="&#128736;",
            color_class="c-success" if total_dummy_qty > 0 else "c-muted",
            key="dummy_tank_btn",
        )
        if clicked_dummy_tank:
            st.session_state["pl_unblock_clicked"] = False
            st.session_state["selected_tile"] = None
            st.session_state["tank_turns_page"] = "main"
            st.session_state["location_visit_page"] = "main"
            st.session_state["dummy_tank_clicked"] = True
            st.rerun()
    with col11:
        clicked_tank_turns = kpi_card(
            label="Tank Turns",
            value=f"{tank_turns_value:.2f}" if tank_turns_value > 0 else "-",
            detail="Dispatches / Tank Capacity",
            icon="&#128167;",
            color_class="c-success" if tank_turns_value > 0 else "c-muted",
            key="tank_turns_btn",
        )
        if clicked_tank_turns:
            st.session_state["dummy_tank_clicked"] = False
            st.session_state["pl_unblock_clicked"] = False
            st.session_state["selected_tile"] = None
            st.session_state["location_visit_page"] = "main"
            st.session_state["tank_turns_page"] = "drilldown"
            st.rerun()
    with col12:
        # Step 6: KPI tile UI for Location Visit (use kpi_card for alignment)
        _lv_label = _prev_qtr_label if "kpi_location_visit" in dir() and kpi_location_visit >= 0 else "Prev Qtr"
        clicked_location_visit = kpi_card(
            label="Location Visit | Compliance",
            value=f"{kpi_location_visit:,} Audits | {kpi_location_compliance * 100:.1f}%",
            detail=f"{_lv_label}  •  Closed: {int(_loc_closed):,}  •  Open: {int(_loc_total - _loc_closed):,}",
            icon="&#128205;",
            color_class="c-success" if kpi_location_visit > 0 else "c-muted",
            key="location_visit_btn"
        )
        if clicked_location_visit:
            st.session_state["pl_unblock_clicked"] = False
            st.session_state["dummy_tank_clicked"] = False
            st.session_state["tank_turns_page"] = "main"
            st.session_state["selected_tile"] = None
            st.session_state["lv_sub_page"] = "summary"
            st.session_state["location_visit_page"] = "drilldown"
            st.rerun()
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    # --- Restore bar and donut diagrams with defensive checks ---
    try:
        exception_kpi_df = _build_exception_kpi_chart_df(
            pending_dc_result_filtered,
            open_delivery_result_filtered,
            open_intransit_result_filtered,
            open_sales_orders_result_filtered,
            pending_invoices_result_filtered,
            tank_reco_result_filtered,
            open_short_sales_result_filtered,
            open_short_sto_result_filtered,
        )
        if (
            exception_kpi_df is not None
            and not exception_kpi_df.empty
            and float(exception_kpi_df["KPI Value"].sum()) > 0
        ):
            st.markdown(
                "<div class='sec-title'>&#128202; Exception Mix Across All KPI Tiles</div>",
                unsafe_allow_html=True,
            )
            st.caption("All KPI values in these charts are shown as exception record counts.")
            _render_exception_kpi_charts(exception_kpi_df)
    except Exception as exc:
        st.warning(f"Exception KPI charts could not be rendered: {exc}")

    # --- Zonewise Exception Table (All KPIs) ---
    st.markdown("<div class='sec-title'>&#128205; Zonewise Exception Summary (All KPIs)</div>", unsafe_allow_html=True)
    zone_summary_df = _build_zone_exception_summary(all_exception_plant_df)
    # Ensure consistent column order
    metric_cols = [
        "Total Exceptions", "Pending DC", "Open Delivery", "Open In-Transit", "Open Sales Order", "Pending Invoice", "Shortage Sales (Billing Docs)", "Shortage STO (Billing Docs)"
    ]
    zone_cols = [c for c in ["Zone Name", "Locations"] + metric_cols if c in zone_summary_df.columns]
    if zone_summary_df is not None and not zone_summary_df.empty:
        _render_html_table(zone_summary_df[zone_cols], max_height=420)
    else:
        st.info("No zonewise exception data available.")

    # --- Locationwise Exception Table (All KPIs) ---
    st.markdown("<div class='sec-title'>&#127981; Locationwise Exception Summary (All KPIs)</div>", unsafe_allow_html=True)
    if all_exception_plant_df is not None and not all_exception_plant_df.empty:
        loc_cols = [c for c in ["Zone Name", "Plant Name"] + metric_cols if c in all_exception_plant_df.columns]
        _render_html_table(all_exception_plant_df[loc_cols], max_height=420)
    else:
        st.info("No locationwise exception data available.")

    # ── Mail Center ──────────────────────────────────────────────────────────
    # Ensure pre-filter copies exist even when source files are absent
    _dummy_for_mail = locals().get("_dummy_for_mail", pd.DataFrame())
    _pl_for_mail    = locals().get("_pl_for_mail",    pd.DataFrame())
    _tt_for_mail    = locals().get("_tt_for_mail",    pd.DataFrame())
    _loc_for_mail   = locals().get("_loc_for_mail",   pd.DataFrame())

    _detail_dfs_for_mail = {
        "Pending DC":                    pending_dc_result.get("detail_df",          pd.DataFrame()),
        "Open Delivery":                 open_delivery_result.get("detail_df",       pd.DataFrame()),
        "Open In-Transit":               open_intransit_result.get("detail_df",      pd.DataFrame()),
        "Open Sales Order":              open_sales_orders_result.get("detail_df",   pd.DataFrame()),
        "Pending Invoice":               pending_invoices_result.get("detail_df",    pd.DataFrame()),
        "Shortage Sales (Billing Docs)": open_short_sales_result.get("detail_df",   pd.DataFrame()),
        "Shortage STO (Billing Docs)":   open_short_sto_result.get("detail_df",     pd.DataFrame()),
    }
    _zone_kpi_for_mail = _build_zone_kpi_dict(
        zone_exception_summary_df = zone_exception_summary_df,
        tank_reco_result          = tank_reco_result,
        dummy_tank_df             = _dummy_for_mail,
        pl_unblock_df             = _pl_for_mail,
        tank_turns_df_all         = _tt_for_mail,
        loc_visit_df              = _loc_for_mail,
    )
    _render_mail_center(
        all_exception_plant_df,
        _detail_dfs_for_mail,
        as_of_date    = as_of_date,
        zone_kpi_dict = _zone_kpi_for_mail,
    )

    # ── Unmatched plant warning ───────────────────────────────────────────────
    unmatched = pending_dc_result.get("unmatched", [])
    if unmatched:
        with st.expander(
            f"&#9888; {len(unmatched)} Plant Code(s) not found in PlantMaster",
            expanded=False,
        ):
            st.warning(
                "The following Sending Plant codes could not be mapped to "
                "PlantMaster. Update PlantMaster or check the data.\n\n"
                + "  |  ".join(str(c) for c in unmatched)
            )


def _build_zone_kpi_dict(
    zone_exception_summary_df: pd.DataFrame,
    tank_reco_result: dict,
    dummy_tank_df: pd.DataFrame,
    pl_unblock_df: pd.DataFrame,
    tank_turns_df_all: pd.DataFrame,
    loc_visit_df: pd.DataFrame,
) -> dict:
    """Build per-zone dict of all KPI values for the mail body tiles."""
    result: dict = {}

    # Collect all zone names from zone_exception_summary_df
    all_zones: list = []
    if zone_exception_summary_df is not None and not zone_exception_summary_df.empty:
        all_zones = sorted(zone_exception_summary_df["Zone Name"].dropna().unique().tolist())

    tank_reco_zone_df = pd.DataFrame()
    if tank_reco_result:
        tank_reco_zone_df = tank_reco_result.get("zone_summary", pd.DataFrame())

    for zone in all_zones:
        kpis: dict = {}

        # ── 7 exception counts ──────────────────────────────────────────────
        if zone_exception_summary_df is not None and not zone_exception_summary_df.empty:
            zrow = zone_exception_summary_df[zone_exception_summary_df["Zone Name"] == zone]
            if not zrow.empty:
                r = zrow.iloc[0]
                kpis["Pending DC"]             = int(pd.to_numeric(r.get("Pending DC", 0), errors="coerce") or 0)
                kpis["Open Delivery"]          = int(pd.to_numeric(r.get("Open Delivery", 0), errors="coerce") or 0)
                kpis["Open In-Transit"]        = int(pd.to_numeric(r.get("Open In-Transit", 0), errors="coerce") or 0)
                kpis["Open Sales Order"]       = int(pd.to_numeric(r.get("Open Sales Order", 0), errors="coerce") or 0)
                kpis["Pending Invoice"]        = int(pd.to_numeric(r.get("Pending Invoice", 0), errors="coerce") or 0)
                kpis["Shortage Sales (Ltrs)"]  = float(pd.to_numeric(r.get("Shortage Sales (Billing Docs)", 0), errors="coerce") or 0)
                kpis["Shortage STO (Ltrs)"]    = float(pd.to_numeric(r.get("Shortage STO (Billing Docs)", 0), errors="coerce") or 0)
                kpis["Total Exceptions"]       = int(pd.to_numeric(r.get("Total Exceptions", 0), errors="coerce") or 0)

        # ── Tank Reco ────────────────────────────────────────────────────────
        if not tank_reco_zone_df.empty and "Zone Name" in tank_reco_zone_df.columns:
            zrow = tank_reco_zone_df[tank_reco_zone_df["Zone Name"] == zone]
            kpis["Tank Reco"] = int(pd.to_numeric(zrow.get("Tank Reco Count", pd.Series([0])).sum(), errors="coerce") or 0) if not zrow.empty else 0
        else:
            kpis["Tank Reco"] = 0

        # ── Dummy Tank (KL) ──────────────────────────────────────────────────
        if dummy_tank_df is not None and not dummy_tank_df.empty and "Zone Name" in dummy_tank_df.columns:
            z_d = dummy_tank_df[dummy_tank_df["Zone Name"] == zone]
            kpis["Dummy Tank (KL)"] = round(float(pd.to_numeric(z_d["Unrestricted"], errors="coerce").fillna(0).sum()) / 1000, 3)
        else:
            kpis["Dummy Tank (KL)"] = 0.0

        # ── PL Unblock (KL) ──────────────────────────────────────────────────
        if pl_unblock_df is not None and not pl_unblock_df.empty and "Zone Name" in pl_unblock_df.columns:
            z_p = pl_unblock_df[pl_unblock_df["Zone Name"] == zone]
            kpis["PL Unblock (KL)"] = round(float(pd.to_numeric(z_p["Unrestricted"], errors="coerce").fillna(0).sum()) / 1000, 3)
        else:
            kpis["PL Unblock (KL)"] = 0.0

        # ── Tank Turns ───────────────────────────────────────────────────────
        if tank_turns_df_all is not None and not tank_turns_df_all.empty and "Zone Name" in tank_turns_df_all.columns:
            z_t = tank_turns_df_all[tank_turns_df_all["Zone Name"] == zone]
            d   = pd.to_numeric(z_t.get("Dispatches",    pd.Series()), errors="coerce").fillna(0).sum()
            c   = pd.to_numeric(z_t.get("Tank Capacity", pd.Series()), errors="coerce").fillna(0).sum()
            kpis["Tank Turns"] = round(d / c, 2) if c > 0 else 0.0
        else:
            kpis["Tank Turns"] = 0.0

        # ── Location Visit ───────────────────────────────────────────────────
        if loc_visit_df is not None and not loc_visit_df.empty and "Zone Name" in loc_visit_df.columns:
            z_l = loc_visit_df[loc_visit_df["Zone Name"] == zone]
            kpis["Locations Visited"]       = int(z_l["Planning Plant"].nunique()) if "Planning Plant" in z_l.columns else 0
            tot = pd.to_numeric(z_l.get("TotalRecomms",  pd.Series()), errors="coerce").fillna(0).sum()
            cls = pd.to_numeric(z_l.get("ClosedRecomms", pd.Series()), errors="coerce").fillna(0).sum()
            kpis["Location Compliance (%)"] = round(cls / tot * 100, 1) if tot > 0 else 0.0
        else:
            kpis["Locations Visited"]       = 0
            kpis["Location Compliance (%)"] = 0.0

        result[zone] = kpis

    return result


def _render_mail_center(
    all_exception_plant_df: pd.DataFrame,
    detail_dfs: dict = None,
    as_of_date = None,
    zone_kpi_dict: dict = None,
) -> None:
    """Mail Center — compose and send zone exception alerts via Outlook."""
    import emails as _em
    import streamlit.components.v1 as _stc

    if detail_dfs is None:
        detail_dfs = {}

    avail, avail_reason = _em.outlook_available()

    st.markdown('<div id="mail-center-anchor"></div>', unsafe_allow_html=True)
    st.markdown(
        "<div class='sec-title'>&#9993; Mail Center — Send Exception Alerts</div>",
        unsafe_allow_html=True,
    )

    if not avail:
        st.info(
            f"&#8505;&#65039; **Mail is only available when running the app locally on Windows "
            f"with Outlook open.**\n\nReason: {avail_reason}"
        )
        return

    if all_exception_plant_df is None or all_exception_plant_df.empty:
        st.warning("No exception data loaded — cannot compose mails.")
        return

    auto_open = bool(st.session_state.get("open_mail_center", False))
    if auto_open:
        st.session_state["open_mail_center"] = False

    with st.expander("&#9993; Compose & Send Exception Mails", expanded=auto_open):

        # Use the sidebar-selected as-of date; fall back to today
        as_of_label = as_of_date.strftime("%d %b %Y") if as_of_date else datetime.now().strftime("%d %b %Y")

        # ── Row 1: Zone selector + test mode ─────────────────────────────────
        col_zones, col_mode = st.columns([3, 1])
        with col_zones:
            available_zones = sorted(
                all_exception_plant_df["Zone Name"].dropna().unique().tolist()
            )
            selected_zones_mail = st.multiselect(
                "Zones to mail  (one email per zone)",
                options=available_zones,
                default=available_zones,
                key="mail_zone_select",
            )
        with col_mode:
            test_mode = st.checkbox(
                "&#128300; Test Mode", value=True, key="mail_test_mode",
                help="Routes all mail to the test address — no zone recipients contacted.",
            )
            if test_mode:
                test_email = st.text_input(
                    "Test address", value=_em.SENDER_EMAIL, key="mail_test_email",
                    label_visibility="collapsed",
                )
            else:
                test_email = ""

        # ── Row 2: Exception type checkboxes ─────────────────────────────────
        st.markdown("**Exception types to include** *(each checked type = one Excel attachment)*")
        exc_options = list(_em.EXCEPTION_LABELS.keys())
        exc_cols_ui = st.columns(len(exc_options))
        selected_exceptions = []
        for i, exc in enumerate(exc_options):
            with exc_cols_ui[i]:
                if st.checkbox(_em.EXCEPTION_LABELS[exc], value=True, key=f"mail_exc_{exc}"):
                    selected_exceptions.append(exc)

        # ── Row 3: Custom intro ───────────────────────────────────────────────
        custom_intro = st.text_area(
            "Custom intro paragraph (optional)",
            value="", height=70, key="mail_custom_intro",
        )

        # ── Preview zone ──────────────────────────────────────────────────────
        if selected_zones_mail and selected_exceptions:
            st.markdown("---")
            preview_zone = st.selectbox(
                "Preview mail for zone:", options=selected_zones_mail,
                key="mail_preview_zone",
            )
            contacts = _em.ZONE_EMAIL_MAP.get(preview_zone, {})

            # Sender / Recipient strip
            r1, r2, r3 = st.columns(3)
            with r1:
                st.markdown(f"**&#128228; From:** `{_em.SENDER_EMAIL}`")
            with r2:
                if test_mode:
                    st.markdown(f"**&#128229; To (TEST):** `{test_email or _em.SENDER_EMAIL}`")
                else:
                    st.markdown(f"**&#128229; To:** `{contacts.get('to', 'Not configured')}`")
                    cc_val = contacts.get('cc', '')
                    if cc_val:
                        st.markdown(f"**CC:** `{cc_val}`")
            with r3:
                if not test_mode:
                    st.markdown(f"**BCC:** `{'; '.join(_em.BCC_EMAILS)}`")
                st.markdown(f"**&#128197; Data as of:** `{as_of_label}`")

            # Attachments that will be generated
            preview_attachments = _em.build_zone_excel_attachments(
                preview_zone, detail_dfs, selected_exceptions
            )
            attach_names = [fname for fname, _ in preview_attachments]
            if attach_names:
                st.markdown(
                    "**&#128206; Attachments that will be sent:**  " +
                    "  |  ".join(f"`{n}`" for n in attach_names)
                )
            else:
                st.warning(f"No data found for '{preview_zone}' with the selected exceptions.")

            # HTML preview
            zone_df_preview = all_exception_plant_df[
                all_exception_plant_df["Zone Name"] == preview_zone
            ].copy()
            zone_kpis_preview = (zone_kpi_dict or {}).get(preview_zone, {})
            preview_html = _em.build_exception_email_html(
                preview_zone, as_of_label, zone_df_preview,
                selected_exceptions, custom_intro, attach_names,
                zone_kpi_dict=zone_kpis_preview,
            )
            st.markdown("**Mail body preview:**")
            _stc.html(preview_html, height=700, scrolling=True)

        # ── Send button ───────────────────────────────────────────────────────
        st.markdown("---")
        send_label = (
            f"&#128300; Send TEST to {test_email or _em.SENDER_EMAIL}"
            if test_mode
            else f"&#9993; Send to {len(selected_zones_mail)} zone(s)"
        )
        if st.button(send_label, key="mail_send_btn", type="primary"):
            if not selected_zones_mail:
                st.warning("Select at least one zone.")
            elif not selected_exceptions:
                st.warning("Select at least one exception type.")
            else:
                results = []
                prog = st.progress(0, text="Sending mails…")
                for idx, zone in enumerate(selected_zones_mail):
                    result = _em.send_exception_mail_for_zone(
                        zone_name=zone,
                        all_exception_plant_df=all_exception_plant_df,
                        detail_dfs=detail_dfs,
                        selected_exceptions=selected_exceptions,
                        as_of_date=as_of_label,
                        custom_intro=custom_intro,
                        test_mode=test_mode,
                        test_email=test_email,
                        zone_kpi_dict=(zone_kpi_dict or {}).get(zone, {}),
                    )
                    results.append((zone, result))
                    prog.progress(
                        (idx + 1) / len(selected_zones_mail),
                        text=f"{idx+1}/{len(selected_zones_mail)}: {zone}",
                    )

                prog.empty()
                ok_zones   = [(z, r) for z, r in results if r.get("ok")]
                fail_zones = [(z, r) for z, r in results if not r.get("ok")]

                for zone, r in ok_zones:
                    attach_list = ", ".join(r.get("attachments", []))
                    mode_str = "sent" if r.get("mode") == "sent" else "saved to Drafts"
                    st.success(
                        f"&#9989; **{zone}** — {mode_str}.  "
                        f"Attachments: {attach_list or '(none)'}"
                    )
                for zone, r in fail_zones:
                    st.error(f"&#10060; **{zone}**: {r.get('msg', 'Unknown error')}")


def _build_exception_kpi_chart_df(
    pending_dc_result: dict,
    open_delivery_result: dict,
    open_intransit_result: dict,
    open_sales_orders_result: dict,
    pending_invoices_result: dict,
    tank_reco_result: dict,
    open_short_sales_result: dict,
    open_short_sto_result: dict,
) -> pd.DataFrame:
    """Build chart input DataFrame from all main dashboard KPI tile values."""
    def _unique_billing_count(df: pd.DataFrame) -> float:
        """Return unique non-blank Billing Document count from shortage detail data."""
        if not isinstance(df, pd.DataFrame) or df.empty or "Billing Document" not in df.columns:
            try:
                exception_kpi_df = _build_exception_kpi_chart_df(
                    pending_dc_result_filtered,
                    open_delivery_result_filtered,
                    open_intransit_result_filtered,
                    open_sales_orders_result_filtered,
                    pending_invoices_result_filtered,
                    tank_reco_result_filtered,
                    open_short_sales_result_filtered,
                    open_short_sto_result_filtered,
                )
                if (
                    exception_kpi_df is not None
                    and not exception_kpi_df.empty
                    and float(exception_kpi_df["KPI Value"].sum()) > 0
                ):
                    st.markdown(
                        "<div class='sec-title'>&#128202; Exception Mix Across All KPI Tiles</div>",
                        unsafe_allow_html=True,
                    )
                    st.caption("All KPI values in these charts are shown as exception record counts.")
                    _render_exception_kpi_charts(exception_kpi_df)
            except Exception as exc:
                st.warning(f"Exception KPI charts could not be rendered: {exc}")
    short_sales_count = float(open_short_sales_result.get("total_count", 0) or 0)
    short_sto_count = float(open_short_sto_result.get("total_count", 0) or 0)
    kpi_rows = [
        {"Exception KPI": "Pending DC", "KPI Value": float(pending_dc_result.get("total_count", 0) or 0), "Unit": "Count"},
        {"Exception KPI": "Open Delivery", "KPI Value": float(open_delivery_result.get("total_count", 0) or 0), "Unit": "Count"},
        {"Exception KPI": "Open In-Transit", "KPI Value": float(open_intransit_result.get("total_count", 0) or 0), "Unit": "Count"},
        {"Exception KPI": "Open Sales Orders", "KPI Value": float(open_sales_orders_result.get("total_count", 0) or 0), "Unit": "Count"},
        {"Exception KPI": "Pending Invoices", "KPI Value": float(pending_invoices_result.get("total_count", 0) or 0), "Unit": "Count"},
        {"Exception KPI": "Tank Reco", "KPI Value": float(tank_reco_result.get("total_count", 0) or 0), "Unit": "Count"},
        {"Exception KPI": "SHORTAGES - Ltrs (Sales)", "KPI Value": short_sales_count, "Unit": "Count"},
        {"Exception KPI": "SHORTAGES - Ltrs (STO)", "KPI Value": short_sto_count, "Unit": "Count"},
    ]

    chart_df = pd.DataFrame(kpi_rows)
    chart_df["KPI Value"] = pd.to_numeric(chart_df["KPI Value"], errors="coerce").fillna(0.0)
    chart_df = chart_df.sort_values("KPI Value", ascending=False).reset_index(drop=True)

    chart_df["Display Value"] = chart_df["KPI Value"].apply(
        lambda v: f"{int(round(v)):,}"
    )
    return chart_df


def _render_exception_kpi_charts(chart_df: pd.DataFrame) -> None:
    """Render colorful bar and donut charts for all exception KPI tiles."""
    palette = [
        "#0B3D91", "#1B66C9", "#2A9D8F", "#F4A261", "#E76F51",
        "#7B2CBF", "#3A86FF", "#43AA8B", "#F9C74F", "#577590",
        "#90BE6D", "#F94144",
    ]
    chart_df["Color"] = [palette[idx % len(palette)] for idx in range(len(chart_df))]

    bar_col, pie_col = st.columns([2.15, 1.15], gap="medium")

    with bar_col:
        y_max = float(pd.to_numeric(chart_df["KPI Value"], errors="coerce").fillna(0).max())
        y_upper = max(1.0, (y_max * 1.16) + 1.0)
        fig_bar = px.bar(
            chart_df,
            x="Exception KPI",
            y="KPI Value",
            text="Display Value",
            labels={"KPI Value": "KPI Value", "Exception KPI": "Exception KPI"},
        )
        fig_bar.update_traces(
            marker_color=chart_df["Color"],
            marker_line_color="#FFFFFF",
            marker_line_width=1.5,
            textposition="outside",
            cliponaxis=False,
            textfont_size=20,
            textfont_color="#163A63",
            hovertemplate="<b>%{x}</b><br>Count: %{y:,.0f}<extra></extra>",
        )
        fig_bar.update_layout(
            plot_bgcolor="#F8FAFD",
            paper_bgcolor="white",
            font=dict(family="Segoe UI", size=20, color="#163A63"),
            margin=dict(l=10, r=10, t=34, b=95),
            showlegend=False,
            height=430,
            xaxis=dict(
                tickangle=-28,
                tickfont=dict(size=18, color="#42566E"),
                title=None,
                showgrid=False,
                zeroline=False,
            ),
            yaxis=dict(
                title="Exception Count",
                title_font=dict(size=18, color="#163A63"),
                tickfont=dict(size=18, color="#42566E"),
                gridcolor="#DCE6F2",
                zeroline=False,
                range=[0, y_upper],
            ),
        )
        st.plotly_chart(fig_bar, width='stretch', config={"displayModeBar": False})

    with pie_col:
        pie_df = chart_df.copy()

        fig_pie = px.pie(
            pie_df,
            names="Exception KPI",
            values="KPI Value",
            hole=0.58,
            color="Exception KPI",
            color_discrete_sequence=pie_df["Color"].tolist(),
        )
        fig_pie.update_traces(
            textposition="inside",
            textinfo="percent",
            textfont_size=18,
            textfont_color="#FFFFFF",
            marker=dict(line=dict(color="white", width=2)),
            customdata=pie_df[["Display Value"]].values,
            hovertemplate="<b>%{label}</b><br>Count: %{customdata[0]}<br>Share: %{percent}<extra></extra>",
        )
        fig_pie.update_layout(
            paper_bgcolor="white",
            plot_bgcolor="white",
            font=dict(family="Segoe UI", size=18, color="#163A63"),
            margin=dict(l=6, r=6, t=10, b=6),
            height=380,
            showlegend=False,
            annotations=[
                dict(
                    text=f"<b>{len(chart_df)}</b><br>KPIs",
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                    font=dict(size=24, color="#0B3D91"),
                )
            ],
        )
        st.plotly_chart(fig_pie, width='stretch', config={"displayModeBar": False})

        legend_rows = "".join(
            f'<div style="display:flex;align-items:center;justify-content:space-between;gap:8px;padding:2px 0;">'
            f'<div style="display:flex;align-items:center;gap:8px;min-width:0;">'
            f'<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:{row["Color"]};flex:0 0 auto;"></span>'
            f'<span style="font-size:13px;color:#163A63;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{html.escape(str(row["Exception KPI"]))}</span>'
            f'</div>'
            f'<b style="font-size:13px;color:#163A63;">{html.escape(str(row["Display Value"]))}</b>'
            f'</div>'
            for _, row in pie_df.iterrows()
        )
        st.markdown(
            f'<div style="margin-top:6px;padding:8px 10px;border:1px solid #DCE6F2;border-radius:8px;background:#F8FAFD;max-height:170px;overflow:auto;">'
            f'{legend_rows}'
            f'</div>',
            unsafe_allow_html=True,
        )


def _extract_zone_plant_metric(summary_df: pd.DataFrame, source_col: str, output_col: str) -> pd.DataFrame:
    """Return Zone+Plant metric DataFrame with a standard output column name."""
    cols = ["Zone Name", "Plant Name", output_col]
    if summary_df is None or summary_df.empty:
        return pd.DataFrame(columns=cols)
    required = {"Zone Name", "Plant Name", source_col}
    if not required.issubset(summary_df.columns):
        return pd.DataFrame(columns=cols)

    out_df = summary_df[["Zone Name", "Plant Name", source_col]].copy()
    out_df[source_col] = pd.to_numeric(out_df[source_col], errors="coerce").fillna(0)
    out_df = (
        out_df.groupby(["Zone Name", "Plant Name"], dropna=False, as_index=False)[source_col]
        .sum()
        .rename(columns={source_col: output_col})
    )
    return out_df


def _extract_shortage_billing_counts(detail_df: pd.DataFrame, output_col: str) -> pd.DataFrame:
    """Return shortage counts by Zone+Plant using unique non-blank Billing Document."""
    cols = ["Zone Name", "Plant Name", output_col]
    if detail_df is None or detail_df.empty:
        return pd.DataFrame(columns=cols)
    required = {"Zone Name", "Plant Name", "Billing Document"}
    if not required.issubset(detail_df.columns):
        return pd.DataFrame(columns=cols)

    work = detail_df[["Zone Name", "Plant Name", "Billing Document"]].copy()
    work["Billing Document"] = work["Billing Document"].astype(str).str.strip()
    work = work[(work["Billing Document"] != "") & (work["Billing Document"].str.lower() != "nan")]
    if work.empty:
        return pd.DataFrame(columns=cols)

    out_df = (
        work.groupby(["Zone Name", "Plant Name"], dropna=False, as_index=False)["Billing Document"]
        .nunique()
        .rename(columns={"Billing Document": output_col})
    )
    return out_df


def _build_all_exception_plant_summary(
    pending_dc_result: dict,
    open_delivery_result: dict,
    open_intransit_result: dict,
    open_sales_orders_result: dict,
    pending_invoices_result: dict,
    tank_reco_result: dict,
    open_short_sales_result: dict,
    open_short_sto_result: dict,
) -> pd.DataFrame:
    """Build combined Zone+Plant exception summary across all KPI modules."""
    metric_cols = [
        "Pending DC",
        "Open Delivery",
        "Open In-Transit",
        "Open Sales Order",
        "Pending Invoice",
        "Shortage Sales (Billing Docs)",
        "Shortage STO (Billing Docs)",
    ]

    frames = [
        _extract_zone_plant_metric(pending_dc_result.get("summary_df", pd.DataFrame()), "Pending DC Count", "Pending DC"),
        _extract_zone_plant_metric(open_delivery_result.get("summary_df", pd.DataFrame()), "Open Delivery Count", "Open Delivery"),
        _extract_zone_plant_metric(open_intransit_result.get("summary_df", pd.DataFrame()), "Open In-Transit STO Count", "Open In-Transit"),
        _extract_zone_plant_metric(open_sales_orders_result.get("summary_df", pd.DataFrame()), "Open Sales Order Count", "Open Sales Order"),
        _extract_zone_plant_metric(pending_invoices_result.get("summary_df", pd.DataFrame()), "Pending Invoice Count", "Pending Invoice"),
        _extract_shortage_billing_counts(open_short_sales_result.get("detail_df", pd.DataFrame()), "Shortage Sales (Billing Docs)"),
        _extract_shortage_billing_counts(open_short_sto_result.get("detail_df", pd.DataFrame()), "Shortage STO (Billing Docs)"),
    ]

    merged_df = pd.DataFrame(columns=["Zone Name", "Plant Name"])
    for frame in frames:
        if frame is None or frame.empty:
            continue
        if merged_df.empty:
            merged_df = frame.copy()
        else:
            merged_df = merged_df.merge(frame, on=["Zone Name", "Plant Name"], how="outer")

    if merged_df.empty:
        return pd.DataFrame(columns=["Zone Name", "Plant Name", *metric_cols, "Total Exceptions"])

    for col in metric_cols:
        if col not in merged_df.columns:
            merged_df[col] = 0
        merged_df[col] = pd.to_numeric(merged_df[col], errors="coerce").fillna(0).round().astype(int)

    merged_df["Total Exceptions"] = merged_df[metric_cols].sum(axis=1).astype(int)
    merged_df = merged_df.sort_values(
        ["Total Exceptions", "Zone Name", "Plant Name"],
        ascending=[False, True, True],
    ).reset_index(drop=True)

    return merged_df[["Zone Name", "Plant Name", *metric_cols, "Total Exceptions"]]


def _build_zone_exception_summary(all_exception_plant_df: pd.DataFrame) -> pd.DataFrame:
    """Build zone totals from the combined Zone+Plant all-exception summary."""
    if all_exception_plant_df is None or all_exception_plant_df.empty:
        return pd.DataFrame(columns=["Zone Name", "Locations", "Total Exceptions"])

    metric_cols = [
        c for c in all_exception_plant_df.columns
        if c not in {"Zone Name", "Plant Name", "Total Exceptions"}
    ]

    zone_totals = (
        all_exception_plant_df.groupby("Zone Name", dropna=False, as_index=False)[metric_cols + ["Total Exceptions"]]
        .sum()
    )
    zone_locations = (
        all_exception_plant_df.groupby("Zone Name", dropna=False, as_index=False)["Plant Name"]
        .nunique()
        .rename(columns={"Plant Name": "Locations"})
    )

    zone_summary = zone_totals.merge(zone_locations, on="Zone Name", how="left")
    zone_summary["Locations"] = pd.to_numeric(zone_summary["Locations"], errors="coerce").fillna(0).astype(int)
    zone_summary = zone_summary.sort_values("Total Exceptions", ascending=False).reset_index(drop=True)
    return zone_summary[["Zone Name", "Locations", "Total Exceptions", *metric_cols]]


def _build_combined_shortage_location_summary(
    open_short_sales_result: dict,
    open_short_sto_result: dict,
) -> pd.DataFrame:
    """Merge Sales and STO shortage summaries into one Zone+Location quantity table."""
    output_cols = [
        "Zone Name",
        "Plant Name",
        "Sales Shortage Quantity (in Ltrs)",
        "STO Shortage Quantity (in Ltrs)",
        "Total Pending Shortage Quantity (in Ltrs)",
    ]

    merged_df = pd.DataFrame(columns=["Zone Name", "Plant Name"])

    sales_df = open_short_sales_result.get("summary_df", pd.DataFrame())
    if sales_df is not None and not sales_df.empty and "Total Shortage Quantity (in Ltrs)" in sales_df.columns:
        sales_df = sales_df[["Zone Name", "Plant Name", "Total Shortage Quantity (in Ltrs)"]].rename(
            columns={"Total Shortage Quantity (in Ltrs)": "Sales Shortage Quantity (in Ltrs)"}
        )
        merged_df = sales_df.copy() if merged_df.empty else merged_df.merge(sales_df, on=["Zone Name", "Plant Name"], how="outer")

    sto_df = open_short_sto_result.get("summary_df", pd.DataFrame())
    if sto_df is not None and not sto_df.empty and "Total STO Shortage Quantity (in Ltrs)" in sto_df.columns:
        sto_df = sto_df[["Zone Name", "Plant Name", "Total STO Shortage Quantity (in Ltrs)"]].rename(
            columns={"Total STO Shortage Quantity (in Ltrs)": "STO Shortage Quantity (in Ltrs)"}
        )
        merged_df = sto_df.copy() if merged_df.empty else merged_df.merge(sto_df, on=["Zone Name", "Plant Name"], how="outer")

    if merged_df.empty:
        return pd.DataFrame(columns=output_cols)

    for col in ["Sales Shortage Quantity (in Ltrs)", "STO Shortage Quantity (in Ltrs)"]:
        if col not in merged_df.columns:
            merged_df[col] = 0.0
        merged_df[col] = pd.to_numeric(merged_df[col], errors="coerce").fillna(0.0)

    merged_df["Total Pending Shortage Quantity (in Ltrs)"] = (
        merged_df["Sales Shortage Quantity (in Ltrs)"]
        + merged_df["STO Shortage Quantity (in Ltrs)"]
    )
    merged_df = merged_df.sort_values(
        ["Total Pending Shortage Quantity (in Ltrs)", "Zone Name", "Plant Name"],
        ascending=[False, True, True],
    ).reset_index(drop=True)
    return merged_df[output_cols]


def _build_combined_shortage_zone_summary(shortage_location_df: pd.DataFrame) -> pd.DataFrame:
    """Build zone-level shortage totals from the combined shortage location summary."""
    output_cols = [
        "Zone Name",
        "Locations",
        "Sales Shortage Quantity (in Ltrs)",
        "STO Shortage Quantity (in Ltrs)",
        "Total Pending Shortage Quantity (in Ltrs)",
    ]
    if shortage_location_df is None or shortage_location_df.empty:
        return pd.DataFrame(columns=output_cols)

    zone_totals = (
        shortage_location_df.groupby("Zone Name", dropna=False, as_index=False)[
            [
                "Sales Shortage Quantity (in Ltrs)",
                "STO Shortage Quantity (in Ltrs)",
                "Total Pending Shortage Quantity (in Ltrs)",
            ]
        ]
        .sum()
    )
    zone_locations = (
        shortage_location_df.groupby("Zone Name", dropna=False, as_index=False)["Plant Name"]
        .nunique()
        .rename(columns={"Plant Name": "Locations"})
    )

    zone_df = zone_totals.merge(zone_locations, on="Zone Name", how="left")
    zone_df["Locations"] = pd.to_numeric(zone_df["Locations"], errors="coerce").fillna(0).astype(int)
    zone_df = zone_df.sort_values("Total Pending Shortage Quantity (in Ltrs)", ascending=False).reset_index(drop=True)
    return zone_df[output_cols]


def _build_combined_shortage_detail_df(
    open_short_sales_result: dict,
    open_short_sto_result: dict,
) -> pd.DataFrame:
    """Standardize Sales and STO shortage detail rows into one drilldown table."""
    output_cols = [
        "Shortage Type",
        "Zone Name",
        "Plant Name",
        "Billing Document",
        "Shipment Number",
        "Vehicle / TT Number",
        "Delivery",
        "Material",
        "Billed Quantity",
        "Shortage Quantity (in Ltrs)",
        "Shortage Age (Days)",
        "Created Date",
    ]
    frames = []

    sales_detail_df = open_short_sales_result.get("detail_df", pd.DataFrame())
    if sales_detail_df is not None and not sales_detail_df.empty:
        sales_map = {
            "Zone Name": "Zone Name",
            "Plant Name": "Plant Name",
            "Billing Document": "Billing Document",
            "Shipment Number": "Shipment Number",
            "TT Number": "Vehicle / TT Number",
            "Delivery": "Delivery",
            "Material": "Material",
            "Billed Quantity": "Billed Quantity",
            "Shortage Quantity (in Ltrs)": "Shortage Quantity (in Ltrs)",
            "Shortage Age (Days)": "Shortage Age (Days)",
            "Created on": "Created Date",
        }
        sales_out = pd.DataFrame()
        for src, dst in sales_map.items():
            if src in sales_detail_df.columns:
                sales_out[dst] = sales_detail_df[src]
        sales_out["Shortage Type"] = "Sales"
        frames.append(sales_out)

    sto_detail_df = open_short_sto_result.get("detail_df", pd.DataFrame())
    if sto_detail_df is not None and not sto_detail_df.empty:
        sto_map = {
            "Zone Name": "Zone Name",
            "Plant Name": "Plant Name",
            "Billing Document": "Billing Document",
            "Shipment Number": "Shipment Number",
            "Vehicle": "Vehicle / TT Number",
            "Delivery": "Delivery",
            "Material": "Material",
            "Billed Quantity": "Billed Quantity",
            "Shortage Quantity (in Ltrs)": "Shortage Quantity (in Ltrs)",
            "Shortage Age (Days)": "Shortage Age (Days)",
            "Created On": "Created Date",
        }
        sto_out = pd.DataFrame()
        for src, dst in sto_map.items():
            if src in sto_detail_df.columns:
                sto_out[dst] = sto_detail_df[src]
        sto_out["Shortage Type"] = "STO"
        frames.append(sto_out)

    if not frames:
        return pd.DataFrame(columns=output_cols)

    combined_df = pd.concat(frames, ignore_index=True, sort=False)
    for col in output_cols:
        if col not in combined_df.columns:
            combined_df[col] = ""

    combined_df["Shortage Quantity (in Ltrs)"] = pd.to_numeric(
        combined_df["Shortage Quantity (in Ltrs)"], errors="coerce"
    ).fillna(0.0)
    if "Billed Quantity" in combined_df.columns:
        combined_df["Billed Quantity"] = pd.to_numeric(combined_df["Billed Quantity"], errors="coerce")
    if "Shortage Age (Days)" in combined_df.columns:
        combined_df["Shortage Age (Days)"] = pd.to_numeric(combined_df["Shortage Age (Days)"], errors="coerce")
    combined_df = combined_df.sort_values("Shortage Quantity (in Ltrs)", ascending=False).reset_index(drop=True)
    return combined_df[output_cols]


def _build_vehicle_shortage_summary(detail_df: pd.DataFrame, id_col: str, output_label: str) -> pd.DataFrame:
    """Aggregate shortage quantity by TT number / vehicle for ranking pages."""
    output_cols = [output_label, "Zone Name", "Plant Name", "Records", "Zones", "Locations", "Total Shortage Quantity (in Ltrs)"]
    if detail_df is None or detail_df.empty or id_col not in detail_df.columns:
        return pd.DataFrame(columns=output_cols)

    work_df = detail_df.copy()
    work_df[id_col] = work_df[id_col].astype(str).str.strip()
    work_df = work_df[
        work_df[id_col].ne("")
        & work_df[id_col].str.lower().ne("nan")
        & work_df[id_col].str.lower().ne("none")
    ]
    if work_df.empty or "Shortage Quantity (in Ltrs)" not in work_df.columns:
        return pd.DataFrame(columns=output_cols)

    work_df["Shortage Quantity (in Ltrs)"] = pd.to_numeric(work_df["Shortage Quantity (in Ltrs)"], errors="coerce").fillna(0.0)

    summary_df = (
        work_df.groupby([id_col, "Zone Name", "Plant Name"], dropna=False)
        .agg(
            Records=(id_col, "size"),
            Zones=("Zone Name", "nunique"),
            Locations=("Plant Name", "nunique"),
            total_shortage=("Shortage Quantity (in Ltrs)", "sum"),
        )
        .reset_index()
        .rename(columns={id_col: output_label, "total_shortage": "Total Shortage Quantity (in Ltrs)"})
        .sort_values(["Total Shortage Quantity (in Ltrs)", output_label], ascending=[False, True])
        .reset_index(drop=True)
    )
    return summary_df[output_cols]


def _render_back_to_dashboard(button_key: str) -> None:
    """Render a standard back button for drilldown pages."""
    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key=button_key):
            st.session_state["page"] = "dashboard"
            st.rerun()


def _render_active_filter_badges(zone_filter: list, plant_filter: list) -> None:
    """Show selected filters consistently across drilldown pages."""
    if not zone_filter and not plant_filter:
        return
    badges = "".join(
        [f'<span class="fbadge">&#128205; {html.escape(str(z))}</span>' for z in zone_filter]
        + [f'<span class="fbadge">&#127981; {html.escape(str(p))}</span>' for p in plant_filter]
    )
    st.markdown(
        f'<div style="margin-bottom:12px;font-size:15px;">Active Filters:&nbsp;{badges}</div>',
        unsafe_allow_html=True,
    )


def _render_ranked_bar_chart(
    chart_df: pd.DataFrame,
    label_col: str,
    value_col: str,
    title: str,
    x_label: str,
    y_label: str,
    color: str = None,
    value_format: str = ",.2f",
) -> None:
    """Render a consistent horizontal ranking chart for critical-view pages."""
    if chart_df is None or chart_df.empty or label_col not in chart_df.columns or value_col not in chart_df.columns:
        return

    plot_df = chart_df[[label_col, value_col]].copy()
    plot_df[label_col] = plot_df[label_col].astype(str)
    plot_df[value_col] = pd.to_numeric(plot_df[value_col], errors="coerce")
    plot_df = plot_df.dropna(subset=[value_col]).sort_values(value_col, ascending=True)
    if plot_df.empty:
        return

    fig = px.bar(
        plot_df,
        x=value_col,
        y=label_col,
        orientation="h",
        text=value_col,
        labels={label_col: y_label, value_col: x_label},
    )
    fig.update_traces(
        marker_color=color or C["primary"],
        texttemplate=f"%{{x:{value_format}}}",
        textposition="outside",
        cliponaxis=False,
        hovertemplate=f"%{{y}}<br>{x_label}: %{{x:{value_format}}}<extra></extra>",
    )
    fig.update_layout(
        title=title,
        height=max(320, 54 * len(plot_df)),
        margin=dict(l=10, r=40, t=48, b=10),
        showlegend=False,
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(showgrid=True, gridcolor="#E6ECF5", zeroline=False),
        yaxis=dict(showgrid=False),
        title_font=dict(size=18, color="#163A63"),
    )
    st.plotly_chart(fig, width='stretch', config={"displayModeBar": False})


def _render_zone_exception_overview(zone_exception_summary_df: pd.DataFrame) -> None:
    """Render zone-wise total exceptions graph and highlight the highest zone."""
    if zone_exception_summary_df is None or zone_exception_summary_df.empty:
        return

    chart_df = zone_exception_summary_df.copy().sort_values("Total Exceptions", ascending=False)
    max_zone = str(chart_df.iloc[0]["Zone Name"])
    max_total = int(chart_df.iloc[0]["Total Exceptions"])

    st.markdown(
        "<div class='sec-title'>&#128205; Zone-wise Total Exceptions (All KPI Modules)</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='margin:-6px 0 10px 0;font-size:16px;color:#163A63;'>"
        f"Highest exception zone (out of {len(chart_df)} zones): "
        f"<b>{html.escape(max_zone)}</b> with <b>{max_total:,}</b> exceptions.</div>",
        unsafe_allow_html=True,
    )

    fig = px.bar(
        chart_df,
        x="Zone Name",
        y="Total Exceptions",
        text="Total Exceptions",
        labels={"Zone Name": "Zone", "Total Exceptions": "Total Exceptions"},
    )
    y_upper = max(1.0, (float(max_total) * 1.16) + 1.0)
    fig.update_traces(
        marker_color=["#C82333" if str(z) == max_zone else "#1B66C9" for z in chart_df["Zone Name"]],
        marker_line_color="#FFFFFF",
        marker_line_width=1.2,
        textposition="outside",
        cliponaxis=False,
        textfont_size=18,
        hovertemplate="<b>%{x}</b><br>Total Exceptions: %{y:,}<extra></extra>",
    )
    fig.update_layout(
        plot_bgcolor="#F8FAFD",
        paper_bgcolor="white",
        font=dict(family="Segoe UI", size=18, color="#163A63"),
        margin=dict(l=10, r=10, t=34, b=95),
        showlegend=False,
        height=430,
        xaxis=dict(tickangle=-28, title=None, showgrid=False, zeroline=False),
        yaxis=dict(title="Total Exceptions", gridcolor="#DCE6F2", zeroline=False, range=[0, y_upper]),
    )
    st.plotly_chart(fig, width='stretch', config={"displayModeBar": False})


def render_zone_exception_drilldown(
    zone_exception_summary_df: pd.DataFrame,
    all_exception_plant_df: pd.DataFrame,
    zone_filter: list,
    plant_filter: list,
) -> None:
    """Dedicated page: Top/Bottom 5 zones and locations by total exceptions."""
    render_header(subtitle="&#128205; Zone Exception Drill Down")

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key="btn_back_zone_drilldown"):
            st.session_state["page"] = "dashboard"
            st.rerun()

    if zone_filter or plant_filter:
        badges = "".join(
            [f'<span class="fbadge">&#128205; {z}</span>' for z in zone_filter]
            + [f'<span class="fbadge">&#127981; {p}</span>' for p in plant_filter]
        )
        st.markdown(
            f'<div style="margin-bottom:12px;font-size:15px;">Active Filters:&nbsp;{badges}</div>',
            unsafe_allow_html=True,
        )

    if zone_exception_summary_df is None or zone_exception_summary_df.empty or all_exception_plant_df is None or all_exception_plant_df.empty:
        st.info("&#8505; No all-exception summary data available for current filters.")
        return

    zone_sorted_desc = zone_exception_summary_df.sort_values("Total Exceptions", ascending=False).reset_index(drop=True)
    zone_sorted_asc = zone_exception_summary_df.sort_values("Total Exceptions", ascending=True).reset_index(drop=True)
    loc_sorted_desc = all_exception_plant_df.sort_values("Total Exceptions", ascending=False).reset_index(drop=True)
    loc_sorted_asc = all_exception_plant_df.sort_values("Total Exceptions", ascending=True).reset_index(drop=True)

    top_zones = zone_sorted_desc.head(5).copy()
    bottom_zones = zone_sorted_asc.head(5).copy()
    top_locations = loc_sorted_desc.head(5).copy()
    bottom_locations = loc_sorted_asc.head(5).copy()

    max_zone_name = str(top_zones.iloc[0]["Zone Name"]) if not top_zones.empty else "N/A"
    max_zone_count = int(top_zones.iloc[0]["Total Exceptions"]) if not top_zones.empty else 0
    max_loc_name = str(top_locations.iloc[0]["Plant Name"]) if not top_locations.empty else "N/A"
    max_loc_count = int(top_locations.iloc[0]["Total Exceptions"]) if not top_locations.empty else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Exceptions", f"{int(zone_sorted_desc['Total Exceptions'].sum()):,}")
    m2.metric("Zones Covered", f"{zone_sorted_desc['Zone Name'].nunique()}")
    m3.metric("Max Zone", f"{max_zone_name}", f"{max_zone_count:,}")
    m4.metric("Max Location", f"{max_loc_name}", f"{max_loc_count:,}")

    zc1, zc2 = st.columns(2, gap="medium")
    with zc1:
        st.markdown("<div class='sec-title'>&#11014; Top 5 Zones by Exceptions</div>", unsafe_allow_html=True)
        fig_top_zones = px.bar(
            top_zones.sort_values("Total Exceptions", ascending=True),
            x="Total Exceptions",
            y="Zone Name",
            orientation="h",
            text="Total Exceptions",
            labels={"Total Exceptions": "Exceptions", "Zone Name": "Zone"},
        )
        fig_top_zones.update_traces(marker_color="#003087", textposition="outside")
        fig_top_zones.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
        st.plotly_chart(fig_top_zones, width='stretch', config={"displayModeBar": False})

    with zc2:
        st.markdown("<div class='sec-title'>&#11015; Bottom 5 Zones by Exceptions</div>", unsafe_allow_html=True)
        fig_bottom_zones = px.bar(
            bottom_zones.sort_values("Total Exceptions", ascending=True),
            x="Total Exceptions",
            y="Zone Name",
            orientation="h",
            text="Total Exceptions",
            labels={"Total Exceptions": "Exceptions", "Zone Name": "Zone"},
        )
        fig_bottom_zones.update_traces(marker_color="#FF6600", textposition="outside")
        fig_bottom_zones.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
        st.plotly_chart(fig_bottom_zones, width='stretch', config={"displayModeBar": False})

    lc1, lc2 = st.columns(2, gap="medium")
    with lc1:
        st.markdown("<div class='sec-title'>&#127981; Top 5 Locations by Exceptions</div>", unsafe_allow_html=True)
        _render_html_table(
            top_locations[["Zone Name", "Plant Name", "Total Exceptions"]],
            col_labels={"Zone Name": "Zone", "Plant Name": "Location", "Total Exceptions": "Exceptions"},
            max_height=280,
        )

    with lc2:
        st.markdown("<div class='sec-title'>&#127981; Bottom 5 Locations by Exceptions</div>", unsafe_allow_html=True)
        _render_html_table(
            bottom_locations[["Zone Name", "Plant Name", "Total Exceptions"]],
            col_labels={"Zone Name": "Zone", "Plant Name": "Location", "Total Exceptions": "Exceptions"},
            max_height=280,
        )

    st.markdown("<div class='sec-title'>&#128196; Zone-level Full Summary</div>", unsafe_allow_html=True)
    _render_html_table(
        zone_sorted_desc,
        col_labels={
            "Zone Name": "Zone",
            "Total Exceptions": "Total",
            "Shortage Sales (Billing Docs)": "Short Sales",
            "Shortage STO (Billing Docs)": "Short STO",
        },
        max_height=420,
    )


def render_top_exception_zones_page(
    zone_exception_summary_df: pd.DataFrame,
    all_exception_plant_df: pd.DataFrame,
    zone_filter: list,
    plant_filter: list,
) -> None:
    """Sidebar page: top 3 zones with highest total exceptions."""
    render_header(subtitle="&#128293; Top 3 Zones with Highest Exceptions")
    _render_back_to_dashboard("btn_back_top_exception_zones")
    _render_active_filter_badges(zone_filter, plant_filter)

    if zone_exception_summary_df is None or zone_exception_summary_df.empty:
        st.info("&#8505; No zone exception data available for the current filters.")
        return

    top_zones = zone_exception_summary_df.sort_values("Total Exceptions", ascending=False).head(3).copy()
    top_zone = str(top_zones.iloc[0]["Zone Name"]) if not top_zones.empty else "N/A"
    top_zone_total = int(top_zones.iloc[0]["Total Exceptions"]) if not top_zones.empty else 0
    top_zone_locations = pd.DataFrame()
    if all_exception_plant_df is not None and not all_exception_plant_df.empty and not top_zones.empty:
        top_zone_locations = all_exception_plant_df[
            all_exception_plant_df["Zone Name"].isin(top_zones["Zone Name"])
        ].sort_values(["Total Exceptions", "Zone Name", "Plant Name"], ascending=[False, True, True]).reset_index(drop=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Top 3 Zones Total", f"{int(top_zones['Total Exceptions'].sum()):,}")
    m2.metric("Highest Zone", top_zone)
    m3.metric("Highest Zone Exceptions", f"{top_zone_total:,}")
    m4.metric("Locations in Top 3", f"{int(top_zones['Locations'].sum()):,}")

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        _download_excel_button(
            label="&#11015;  Download Zone Exception Report  (.xlsx)",
            file_prefix="TopExceptionZones_Report",
            sheets={
                "Top Zones": top_zones,
                "Zone Locations": top_zone_locations,
            },
            key="dl_top_exception_zones",
        )

    chart_col1, chart_col2 = st.columns(2, gap="medium")
    with chart_col1:
        _render_ranked_bar_chart(
            top_zones,
            label_col="Zone Name",
            value_col="Total Exceptions",
            title="Top Exception Zones",
            x_label="Exceptions",
            y_label="Zone",
            color=C["primary"],
            value_format=",.0f",
        )
    with chart_col2:
        if not top_zone_locations.empty:
            top_location_chart_df = top_zone_locations.head(10).copy()
            top_location_chart_df["Location Label"] = (
                top_location_chart_df["Plant Name"].astype(str)
                + " ("
                + top_location_chart_df["Zone Name"].astype(str)
                + ")"
            )
            _render_ranked_bar_chart(
                top_location_chart_df,
                label_col="Location Label",
                value_col="Total Exceptions",
                title="Top 10 Locations within Leading Zones",
                x_label="Exceptions",
                y_label="Location",
                color=C["accent"],
                value_format=",.0f",
            )

    st.markdown("<div class='sec-title'>&#128205; Ranked Zone Summary</div>", unsafe_allow_html=True)
    zone_cols = [
        c for c in [
            "Zone Name", "Locations", "Total Exceptions", "Pending DC", "Open Delivery",
            "Open In-Transit", "Open Sales Order", "Pending Invoice", "Tank Reco",
            "Shortage Sales (Billing Docs)", "Shortage STO (Billing Docs)"
        ] if c in top_zones.columns
    ]
    _render_html_table(
        top_zones[zone_cols],
        col_labels={
            "Zone Name": "Zone",
            "Total Exceptions": "Total",
            "Shortage Sales (Billing Docs)": "Short Sales",
            "Shortage STO (Billing Docs)": "Short STO",
        },
        max_height=260,
    )

    if all_exception_plant_df is not None and not all_exception_plant_df.empty:
        for zone_name in top_zones["Zone Name"].tolist():
            zone_locations = (
                all_exception_plant_df[all_exception_plant_df["Zone Name"] == zone_name]
                .sort_values("Total Exceptions", ascending=False)
                .head(10)
                .reset_index(drop=True)
            )
            st.markdown(
                f"<div class='sec-title'>&#127981; Top Locations in {html.escape(str(zone_name))}</div>",
                unsafe_allow_html=True,
            )
            _render_html_table(
                zone_locations,
                col_labels={
                    "Zone Name": "Zone",
                    "Plant Name": "Location",
                    "Total Exceptions": "Total",
                    "Shortage Sales (Billing Docs)": "Short Sales",
                    "Shortage STO (Billing Docs)": "Short STO",
                },
                max_height=280,
            )


def render_top_exception_locations_page(
    all_exception_plant_df: pd.DataFrame,
    zone_filter: list,
    plant_filter: list,
) -> None:
    """Sidebar page: top 10 locations with highest total exceptions."""
    render_header(subtitle="&#127981; Top 10 Locations with Highest Exceptions")
    _render_back_to_dashboard("btn_back_top_exception_locations")
    _render_active_filter_badges(zone_filter, plant_filter)

    if all_exception_plant_df is None or all_exception_plant_df.empty:
        st.info("&#8505; No location exception data available for the current filters.")
        return

    top_locations = all_exception_plant_df.sort_values("Total Exceptions", ascending=False).head(10).copy()
    top_location = str(top_locations.iloc[0]["Plant Name"]) if not top_locations.empty else "N/A"
    top_total = int(top_locations.iloc[0]["Total Exceptions"]) if not top_locations.empty else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Top 10 Total Exceptions", f"{int(top_locations['Total Exceptions'].sum()):,}")
    m2.metric("Zones Covered", f"{top_locations['Zone Name'].nunique()}")
    m3.metric("Highest Location", top_location)
    m4.metric("Highest Location Total", f"{top_total:,}")

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        _download_excel_button(
            label="&#11015;  Download Location Exception Report  (.xlsx)",
            file_prefix="TopExceptionLocations_Report",
            sheets={
                "Top Locations": top_locations,
            },
            key="dl_top_exception_locations",
        )

    top_location_chart_df = top_locations.copy()
    top_location_chart_df["Location Label"] = (
        top_location_chart_df["Plant Name"].astype(str)
        + " ("
        + top_location_chart_df["Zone Name"].astype(str)
        + ")"
    )
    _render_ranked_bar_chart(
        top_location_chart_df,
        label_col="Location Label",
        value_col="Total Exceptions",
        title="Top 10 Locations by Total Exceptions",
        x_label="Exceptions",
        y_label="Location",
        color=C["primary"],
        value_format=",.0f",
    )

    st.markdown("<div class='sec-title'>&#128196; Top 10 Location Exception Summary</div>", unsafe_allow_html=True)
    _render_html_table(
        top_locations,
        col_labels={
            "Zone Name": "Zone",
            "Plant Name": "Location",
            "Total Exceptions": "Total",
            "Shortage Sales (Billing Docs)": "Short Sales",
            "Shortage STO (Billing Docs)": "Short STO",
        },
        max_height=420,
    )


def render_top_shortage_zones_page(
    shortage_zone_summary_df: pd.DataFrame,
    shortage_location_summary_df: pd.DataFrame,
    combined_shortage_detail_df: pd.DataFrame,
    zone_filter: list,
    plant_filter: list,
) -> None:
    """Sidebar page: top 3 zones with maximum pending shortage quantity."""
    render_header(subtitle="&#128205; Top 3 Zones by Pending Shortage Quantity")
    _render_back_to_dashboard("btn_back_top_shortage_zones")
    _render_active_filter_badges(zone_filter, plant_filter)

    if shortage_zone_summary_df is None or shortage_zone_summary_df.empty:
        st.info("&#8505; No shortage quantity summary is available for the current filters.")
        return

    top_zones = shortage_zone_summary_df.sort_values("Total Pending Shortage Quantity (in Ltrs)", ascending=False).head(3).copy()
    top_zone = str(top_zones.iloc[0]["Zone Name"]) if not top_zones.empty else "N/A"
    top_qty = float(top_zones.iloc[0]["Total Pending Shortage Quantity (in Ltrs)"]) if not top_zones.empty else 0.0
    top_zone_locations = pd.DataFrame()
    top_zone_detail_df = pd.DataFrame()
    if shortage_location_summary_df is not None and not shortage_location_summary_df.empty and not top_zones.empty:
        top_zone_locations = shortage_location_summary_df[
            shortage_location_summary_df["Zone Name"].isin(top_zones["Zone Name"])
        ].sort_values(
            ["Total Pending Shortage Quantity (in Ltrs)", "Zone Name", "Plant Name"],
            ascending=[False, True, True],
        ).reset_index(drop=True)
    if combined_shortage_detail_df is not None and not combined_shortage_detail_df.empty and not top_zones.empty:
        top_zone_detail_df = combined_shortage_detail_df[
            combined_shortage_detail_df["Zone Name"].isin(top_zones["Zone Name"])
        ].sort_values(
            ["Shortage Quantity (in Ltrs)", "Zone Name", "Plant Name"],
            ascending=[False, True, True],
        ).reset_index(drop=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Top 3 Shortage Qty", f"{top_zones['Total Pending Shortage Quantity (in Ltrs)'].sum():,.2f}")
    m2.metric("Highest Zone", top_zone)
    m3.metric("Highest Zone Qty", f"{top_qty:,.2f}")
    m4.metric("Locations in Top 3", f"{int(top_zones['Locations'].sum()):,}")

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        _download_excel_button(
            label="&#11015;  Download Zone Shortage Report  (.xlsx)",
            file_prefix="TopShortageZones_Report",
            sheets={
                "Top Zones": top_zones,
                "Zone Locations": top_zone_locations,
                "Underlying Records": top_zone_detail_df,
            },
            key="dl_top_shortage_zones",
        )

    chart_col1, chart_col2 = st.columns(2, gap="medium")
    with chart_col1:
        _render_ranked_bar_chart(
            top_zones,
            label_col="Zone Name",
            value_col="Total Pending Shortage Quantity (in Ltrs)",
            title="Top Zones by Pending Shortage Quantity",
            x_label="Pending Shortage Qty (Ltrs)",
            y_label="Zone",
            color=C["primary"],
        )
    with chart_col2:
        if not top_zone_locations.empty:
            top_location_chart_df = top_zone_locations.head(10).copy()
            top_location_chart_df["Location Label"] = (
                top_location_chart_df["Plant Name"].astype(str)
                + " ("
                + top_location_chart_df["Zone Name"].astype(str)
                + ")"
            )
            _render_ranked_bar_chart(
                top_location_chart_df,
                label_col="Location Label",
                value_col="Total Pending Shortage Quantity (in Ltrs)",
                title="Top 10 Locations within Leading Shortage Zones",
                x_label="Pending Shortage Qty (Ltrs)",
                y_label="Location",
                color=C["accent"],
            )

    st.markdown("<div class='sec-title'>&#128202; Ranked Zone Shortage Summary</div>", unsafe_allow_html=True)
    _render_html_table(
        top_zones,
        col_labels={
            "Zone Name": "Zone",
            "Sales Shortage Quantity (in Ltrs)": "Sales Qty (Ltrs)",
            "STO Shortage Quantity (in Ltrs)": "STO Qty (Ltrs)",
            "Total Pending Shortage Quantity (in Ltrs)": "Total Qty (Ltrs)",
        },
        max_height=260,
    )

    for zone_name in top_zones["Zone Name"].tolist():
        zone_locations = (
            shortage_location_summary_df[shortage_location_summary_df["Zone Name"] == zone_name]
            .sort_values("Total Pending Shortage Quantity (in Ltrs)", ascending=False)
            .head(10)
            .reset_index(drop=True)
        )
        st.markdown(
            f"<div class='sec-title'>&#127981; Top Locations in {html.escape(str(zone_name))}</div>",
            unsafe_allow_html=True,
        )
        _render_html_table(
            zone_locations,
            col_labels={
                "Plant Name": "Location",
                "Sales Shortage Quantity (in Ltrs)": "Sales Qty (Ltrs)",
                "STO Shortage Quantity (in Ltrs)": "STO Qty (Ltrs)",
                "Total Pending Shortage Quantity (in Ltrs)": "Total Qty (Ltrs)",
            },
            max_height=280,
        )

        if combined_shortage_detail_df is not None and not combined_shortage_detail_df.empty:
            with st.expander(f"{zone_name}  |  Underlying shortage records", expanded=False):
                zone_detail_df = combined_shortage_detail_df[
                    combined_shortage_detail_df["Zone Name"] == zone_name
                ].copy()
                _render_html_table(zone_detail_df.head(50), max_height=360)


def render_top_shortage_locations_page(
    shortage_location_summary_df: pd.DataFrame,
    combined_shortage_detail_df: pd.DataFrame,
    zone_filter: list,
    plant_filter: list,
) -> None:
    """Sidebar page: top 10 locations with maximum pending shortage quantity."""
    render_header(subtitle="&#127981; Top 10 Locations by Pending Shortage Quantity")
    _render_back_to_dashboard("btn_back_top_shortage_locations")
    _render_active_filter_badges(zone_filter, plant_filter)

    if shortage_location_summary_df is None or shortage_location_summary_df.empty:
        st.info("&#8505; No shortage location data is available for the current filters.")
        return

    top_locations = shortage_location_summary_df.sort_values("Total Pending Shortage Quantity (in Ltrs)", ascending=False).head(10).copy()
    top_location = str(top_locations.iloc[0]["Plant Name"]) if not top_locations.empty else "N/A"
    top_qty = float(top_locations.iloc[0]["Total Pending Shortage Quantity (in Ltrs)"]) if not top_locations.empty else 0.0
    top_location_detail_df = pd.DataFrame()
    if combined_shortage_detail_df is not None and not combined_shortage_detail_df.empty and not top_locations.empty:
        top_pairs = set(zip(top_locations["Zone Name"].astype(str), top_locations["Plant Name"].astype(str)))
        top_location_detail_df = combined_shortage_detail_df[
            combined_shortage_detail_df.apply(
                lambda row: (str(row.get("Zone Name", "")), str(row.get("Plant Name", ""))) in top_pairs,
                axis=1,
            )
        ].sort_values(
            ["Shortage Quantity (in Ltrs)", "Zone Name", "Plant Name"],
            ascending=[False, True, True],
        ).reset_index(drop=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Top 10 Shortage Qty", f"{top_locations['Total Pending Shortage Quantity (in Ltrs)'].sum():,.2f}")
    m2.metric("Zones Covered", f"{top_locations['Zone Name'].nunique()}")
    m3.metric("Highest Location", top_location)
    m4.metric("Highest Location Qty", f"{top_qty:,.2f}")

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        _download_excel_button(
            label="&#11015;  Download Location Shortage Report  (.xlsx)",
            file_prefix="TopShortageLocations_Report",
            sheets={
                "Top Locations": top_locations,
                "Underlying Records": top_location_detail_df,
            },
            key="dl_top_shortage_locations",
        )

    top_location_chart_df = top_locations.copy()
    top_location_chart_df["Location Label"] = (
        top_location_chart_df["Plant Name"].astype(str)
        + " ("
        + top_location_chart_df["Zone Name"].astype(str)
        + ")"
    )
    _render_ranked_bar_chart(
        top_location_chart_df,
        label_col="Location Label",
        value_col="Total Pending Shortage Quantity (in Ltrs)",
        title="Top 10 Locations by Pending Shortage Quantity",
        x_label="Pending Shortage Qty (Ltrs)",
        y_label="Location",
        color=C["primary"],
    )

    st.markdown("<div class='sec-title'>&#128196; Top 10 Location Shortage Summary</div>", unsafe_allow_html=True)
    _render_html_table(
        top_locations,
        col_labels={
            "Zone Name": "Zone",
            "Plant Name": "Location",
            "Sales Shortage Quantity (in Ltrs)": "Sales Qty (Ltrs)",
            "STO Shortage Quantity (in Ltrs)": "STO Qty (Ltrs)",
            "Total Pending Shortage Quantity (in Ltrs)": "Total Qty (Ltrs)",
        },
        max_height=420,
    )

    if combined_shortage_detail_df is not None and not combined_shortage_detail_df.empty:
        for _, row in top_locations.iterrows():
            zone_name = row["Zone Name"]
            plant_name = row["Plant Name"]
            location_detail_df = combined_shortage_detail_df[
                (combined_shortage_detail_df["Zone Name"] == zone_name)
                & (combined_shortage_detail_df["Plant Name"] == plant_name)
            ].copy()
            with st.expander(f"{plant_name} ({zone_name})  |  Underlying shortage records", expanded=False):
                _render_html_table(location_detail_df.head(40), max_height=320)


def render_top_short_sales_vehicles_page(
    short_sales_vehicle_summary_df: pd.DataFrame,
    open_short_sales_result: dict,
    zone_filter: list,
    plant_filter: list,
) -> None:
    """Sidebar page: top 10 TT numbers / vehicles for Sales shortage bookings."""
    render_header(subtitle="&#128666; Top 10 TT Numbers by Pending Sales Shortage Quantity")
    _render_back_to_dashboard("btn_back_top_short_sales_vehicles")
    _render_active_filter_badges(zone_filter, plant_filter)

    detail_df = open_short_sales_result.get("detail_df", pd.DataFrame())
    if short_sales_vehicle_summary_df is None or short_sales_vehicle_summary_df.empty:
        st.info("&#8505; No TT Number based pending Sales shortage data is available for the current filters.")
        return

    top_items = short_sales_vehicle_summary_df.head(10).copy()
    top_item = str(top_items.iloc[0]["TT Number"]) if not top_items.empty else "N/A"
    top_qty = float(top_items.iloc[0]["Total Shortage Quantity (in Ltrs)"]) if not top_items.empty else 0.0
    top_item_detail_df = pd.DataFrame()
    if detail_df is not None and not detail_df.empty and "TT Number" in detail_df.columns and not top_items.empty:
        top_item_detail_df = detail_df[
            detail_df["TT Number"].astype(str).str.strip().isin(top_items["TT Number"].astype(str).str.strip())
        ].sort_values("Shortage Quantity (in Ltrs)", ascending=False).reset_index(drop=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Top 10 Sales TT Qty", f"{top_items['Total Shortage Quantity (in Ltrs)'].sum():,.2f}")
    m2.metric("TT Numbers", f"{len(top_items)}")
    m3.metric("Highest TT Number", top_item)
    m4.metric("Highest TT Qty", f"{top_qty:,.2f}")

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        _download_excel_button(
            label="&#11015;  Download Pending Sales Shortage TT Report  (.xlsx)",
            file_prefix="TopSalesTT_Report",
            sheets={
                "Top Sales TT": top_items,
                "Underlying Records": top_item_detail_df,
            },
            key="dl_top_short_sales_vehicles",
        )

    chart_col1, chart_col2 = st.columns(2, gap="medium")
    with chart_col1:
        _render_ranked_bar_chart(
            top_items,
            label_col="TT Number",
            value_col="Total Shortage Quantity (in Ltrs)",
            title="Top 10 TT Numbers by Pending Sales Shortage Quantity",
            x_label="Pending Sales Shortage Qty (Ltrs)",
            y_label="TT Number",
            color=C["primary"],
        )
    with chart_col2:
        if not top_item_detail_df.empty and "Plant Name" in top_item_detail_df.columns:
            location_chart_df = (
                top_item_detail_df.groupby(["Zone Name", "Plant Name"], dropna=False, as_index=False)["Shortage Quantity (in Ltrs)"]
                .sum()
                .sort_values("Shortage Quantity (in Ltrs)", ascending=False)
                .head(10)
            )
            location_chart_df["Location Label"] = (
                location_chart_df["Plant Name"].astype(str)
                + " ("
                + location_chart_df["Zone Name"].astype(str)
                + ")"
            )
            _render_ranked_bar_chart(
                location_chart_df,
                label_col="Location Label",
                value_col="Shortage Quantity (in Ltrs)",
                title="Top 10 Locations behind Sales TT Shortages",
                x_label="Pending Sales Shortage Qty (Ltrs)",
                y_label="Location",
                color=C["accent"],
            )

    st.markdown("<div class='sec-title'>&#128202; Top 10 TT Numbers by Pending Sales Shortage Quantity</div>", unsafe_allow_html=True)
    _render_html_table(top_items, max_height=360)

    if detail_df is not None and not detail_df.empty and "TT Number" in detail_df.columns:
        detail_cols = [
            c for c in [
                "Zone Name", "Plant Name", "Billing Document", "Shipment Number", "TT Number",
                "Delivery", "Material", "Billed Quantity", "Shortage Quantity (in Ltrs)",
                "Shortage Age (Days)", "Created on"
            ] if c in detail_df.columns
        ]
        for tt_number in top_items["TT Number"].tolist():
            tt_detail_df = detail_df[detail_df["TT Number"].astype(str).str.strip() == str(tt_number)].copy()
            tt_detail_df = tt_detail_df.sort_values("Shortage Quantity (in Ltrs)", ascending=False).reset_index(drop=True)
            with st.expander(f"TT Number {tt_number}  |  {len(tt_detail_df)} record(s)", expanded=False):
                _render_html_table(tt_detail_df[detail_cols].head(40), max_height=320)


def render_top_short_sto_vehicles_page(
    short_sto_vehicle_summary_df: pd.DataFrame,
    open_short_sto_result: dict,
    zone_filter: list,
    plant_filter: list,
) -> None:
    """Sidebar page: top 10 STO vehicles by pending shortage quantity."""
    render_header(subtitle="&#128666; Top 10 Vehicles by Pending STO Shortage Quantity")
    _render_back_to_dashboard("btn_back_top_short_sto_vehicles")
    _render_active_filter_badges(zone_filter, plant_filter)

    detail_df = open_short_sto_result.get("detail_df", pd.DataFrame())
    if short_sto_vehicle_summary_df is None or short_sto_vehicle_summary_df.empty:
        st.info("&#8505; No Vehicle based pending STO shortage data is available for the current filters.")
        return

    top_items = short_sto_vehicle_summary_df.head(10).copy()
    top_item = str(top_items.iloc[0]["Vehicle"]) if not top_items.empty else "N/A"
    top_qty = float(top_items.iloc[0]["Total Shortage Quantity (in Ltrs)"]) if not top_items.empty else 0.0
    top_item_detail_df = pd.DataFrame()
    if detail_df is not None and not detail_df.empty and "Vehicle" in detail_df.columns and not top_items.empty:
        top_item_detail_df = detail_df[
            detail_df["Vehicle"].astype(str).str.strip().isin(top_items["Vehicle"].astype(str).str.strip())
        ].sort_values("Shortage Quantity (in Ltrs)", ascending=False).reset_index(drop=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Top 10 STO Vehicle Qty", f"{top_items['Total Shortage Quantity (in Ltrs)'].sum():,.2f}")
    m2.metric("Vehicles", f"{len(top_items)}")
    m3.metric("Highest Vehicle", top_item)
    m4.metric("Highest Vehicle Qty", f"{top_qty:,.2f}")

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        _download_excel_button(
            label="&#11015;  Download Pending STO Shortage Vehicle Report  (.xlsx)",
            file_prefix="TopSTOVehicles_Report",
            sheets={
                "Top STO Vehicles": top_items,
                "Underlying Records": top_item_detail_df,
            },
            key="dl_top_short_sto_vehicles",
        )

    chart_col1, chart_col2 = st.columns(2, gap="medium")
    with chart_col1:
        _render_ranked_bar_chart(
            top_items,
            label_col="Vehicle",
            value_col="Total Shortage Quantity (in Ltrs)",
            title="Top 10 Vehicles by Pending STO Shortage Quantity",
            x_label="Pending STO Shortage Qty (Ltrs)",
            y_label="Vehicle",
            color=C["primary"],
        )
    with chart_col2:
        if not top_item_detail_df.empty and "Plant Name" in top_item_detail_df.columns:
            location_chart_df = (
                top_item_detail_df.groupby(["Zone Name", "Plant Name"], dropna=False, as_index=False)["Shortage Quantity (in Ltrs)"]
                .sum()
                .sort_values("Shortage Quantity (in Ltrs)", ascending=False)
                .head(10)
            )
            location_chart_df["Location Label"] = (
                location_chart_df["Plant Name"].astype(str)
                + " ("
                + location_chart_df["Zone Name"].astype(str)
                + ")"
            )
            _render_ranked_bar_chart(
                location_chart_df,
                label_col="Location Label",
                value_col="Shortage Quantity (in Ltrs)",
                title="Top 10 Locations behind STO Vehicle Shortages",
                x_label="Pending STO Shortage Qty (Ltrs)",
                y_label="Location",
                color=C["accent"],
            )

    st.markdown("<div class='sec-title'>&#128202; Top 10 Vehicles by Pending STO Shortage Quantity</div>", unsafe_allow_html=True)
    _render_html_table(top_items, max_height=360)

    if detail_df is not None and not detail_df.empty and "Vehicle" in detail_df.columns:
        detail_cols = [
            c for c in [
                "Zone Name", "Plant Name", "Supplying Plant", "Billing Document", "Shipment Number",
                "Vehicle", "Delivery", "Material", "Billed Quantity", "Sales Unit",
                "Shortage Quantity (in Ltrs)", "Shortage Age (Days)", "Created On"
            ] if c in detail_df.columns
        ]
        for vehicle in top_items["Vehicle"].tolist():
            vehicle_detail_df = detail_df[detail_df["Vehicle"].astype(str).str.strip() == str(vehicle)].copy()
            vehicle_detail_df = vehicle_detail_df.sort_values("Shortage Quantity (in Ltrs)", ascending=False).reset_index(drop=True)
            with st.expander(f"Vehicle {vehicle}  |  {len(vehicle_detail_df)} record(s)", expanded=False):
                _render_html_table(vehicle_detail_df[detail_cols].head(40), max_height=320)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: PENDING DC DETAILS (DRILL-DOWN)
# ─────────────────────────────────────────────────────────────────────────────

def render_pending_dc_details(
    pending_dc_result : dict,
    zone_filter       : list,
    plant_filter      : list,
) -> None:
    """Drill-down detail page: zone pivot, plant tabs, raw data, download."""
    render_header(subtitle="&#128666; Pending DC's &#8212; Drill Down")

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key="btn_back"):
            st.session_state["page"] = "dashboard"
            st.rerun()

    st.markdown("""
    <div class="detail-hdr">
        <h3>&#128666; Pending DC's &#8212; Detailed Exception View</h3>
        <p>Zone-wise and Plant-wise breakdown of all Pending Delivery Challans</p>
    </div>
    """, unsafe_allow_html=True)

    summary_df   = pending_dc_result.get("summary_df",   pd.DataFrame())
    zone_summary = pending_dc_result.get("zone_summary",  pd.DataFrame())
    detail_df    = pending_dc_result.get("detail_df",    pd.DataFrame())
    col1, col2 = st.columns([2, 1])
    with col1:
        total_dc   = pending_dc_result.get("total_count", 0)
        s_df       = pending_dc_result.get("summary_df", pd.DataFrame())
        z_df       = pending_dc_result.get("zone_summary", pd.DataFrame())
        detail_str = f"{len(z_df)} zones  |  {len(s_df)} plants affected"
        color_cls  = "c-danger" if total_dc > 50 else ("c-warning" if total_dc > 20 else "")
        clicked_dc = kpi_card(
            label       = "Pending DC's",
            value       = total_dc,
            detail      = detail_str,
            icon        = "&#128666;",
            color_class = color_cls,
            key         = "tile_pending_dc",
        )
        if clicked_dc:
            st.session_state["page"] = "pending_dc_details"
            st.rerun()
    with col2:
        try:
            dq = detail_df["QUANTITY"].sum()
            st.metric("Total Qty (L)", f"{dq:,.0f}")
        except Exception:
            st.metric("Total Qty (L)", "N/A")

    st.markdown("---")

    # ── Zone-level table ──────────────────────────────────────────────────────
    st.markdown(
        "<div class='sec-title'>&#128205; Zone-wise Summary</div>",
        unsafe_allow_html=True,
    )
    if not zone_summary.empty:
        _render_html_table(
            zone_summary,
            col_labels={"Plants": "Plants Affected", "Pending DC Count": "Pending DC's"},
            max_height=380,
        )

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#127981; Plant-wise Drill Down (Zone &#8594; Plant &#8594; Count)</div>",
        unsafe_allow_html=True,
    )

    # ── Zone tabs ─────────────────────────────────────────────────────────────
    all_zones_in_data = sorted(summary_df["Zone Name"].dropna().unique().tolist())

    if len(all_zones_in_data) <= 10:
        tabs = st.tabs(all_zones_in_data)
        for tab, zone in zip(tabs, all_zones_in_data):
            with tab:
                z_df       = summary_df[summary_df["Zone Name"] == zone][
                    ["Plant Name", "Pending DC Count"]
                ].reset_index(drop=True)
                zone_total = int(z_df["Pending DC Count"].sum())
                st.markdown(
                    f"<p style='font-size:18px;font-weight:700;color:#1B3552;margin:4px 0 10px;'>"
                    f"&#127981; {zone} &nbsp;—&nbsp; {len(z_df)} plant(s) &nbsp;|&nbsp; "
                    f"<span style='color:#C0392B;'>{zone_total} Pending DC&#39;s</span></p>",
                    unsafe_allow_html=True,
                )
                _render_html_table(
                    z_df,
                    col_labels={"Plant Name": "Plant", "Pending DC Count": "Pending DC's"},
                    max_height=420,
                )
    else:
        sel_zone = st.selectbox(
            "Select Zone to Expand",
            ["— All Zones —"] + all_zones_in_data,
            key="sel_zone_detail",
        )
        disp_df = (
            summary_df.reset_index(drop=True)
            if sel_zone == "— All Zones —"
            else summary_df[summary_df["Zone Name"] == sel_zone][
                ["Plant Name", "Pending DC Count"]
            ].reset_index(drop=True)
        )
        _render_html_table(
            disp_df,
            col_labels={"Zone Name": "Zone", "Plant Name": "Plant", "Pending DC Count": "Pending DC's"},
            max_height=500,
        )

    # ── Raw shipment detail ───────────────────────────────────────────────────
    if not detail_df.empty:
        with st.expander("&#128269;  Raw Shipment-level Records", expanded=False):
            show_cols = [
                "Zone Name", "Plant Name", "SENDING PLANT",
                "SHIPMENT", "MATERIAL", "DELIVERY", "DELIVERY STATUS",
                "SHIPMENT STATUS", "BILLING DATE",
                "ORDER NO", "VEHICLE NUMBER", "QUANTITY", "QTY UOM",
            ]
            show_cols = [c for c in show_cols if c in detail_df.columns]
            _render_html_table(
                detail_df[show_cols]
                .sort_values([c for c in ["Zone Name", "Plant Name", "SHIPMENT"] if c in detail_df.columns])
                .reset_index(drop=True),
                max_height=560,
            )

    # ── Download button ───────────────────────────────────────────────────────
    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        xlsx_bytes = export_to_excel({
            "Zone Summary" : zone_summary,
            "Plant Summary": summary_df,
            "Raw Data"     : detail_df if not detail_df.empty else pd.DataFrame(),
        })
        st.download_button(
            label     = "&#11015;  Download Report  (.xlsx)",
            data      = xlsx_bytes,
            file_name = f"PendingDC_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key       = "dl_pending_dc",
        )


def render_open_delivery_details(
    open_delivery_result: dict,
    zone_filter         : list,
    plant_filter        : list,
) -> None:
    """Drill-down detail page for Open Deliveries."""
    render_header(subtitle="&#128230; Open Deliveries &#8212; Drill Down")

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key="btn_back_open_delivery"):
            st.session_state["page"] = "dashboard"
            st.rerun()

    st.markdown("""
    <div class="detail-hdr">
        <h3>&#128230; Open Deliveries &#8212; Detailed Exception View</h3>
        <p>Zone-wise and Plant-wise breakdown of unique open Delivery numbers,
        mapped from Shipping Point/Receiving Pt to Plant Master.</p>
    </div>
    """, unsafe_allow_html=True)

    summary_df   = open_delivery_result.get("summary_df", pd.DataFrame())
    zone_summary = open_delivery_result.get("zone_summary", pd.DataFrame())
    detail_df    = open_delivery_result.get("detail_df", pd.DataFrame())
    total_count  = int(open_delivery_result.get("total_count", 0) or 0)

    if summary_df.empty:
        st.info("&#8505; No Open Delivery data available for the current filter selection.")
        return

    total_zones  = int(summary_df["Zone Name"].nunique())
    total_plants = int(summary_df["Plant Name"].nunique())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Open Deliveries (Total)", f"{total_count:,}")
    m2.metric("Zones Affected", f"{total_zones}")
    m3.metric("Plants Affected", f"{total_plants}")
    if not detail_df.empty and "Delivery Age (Days)" in detail_df.columns:
        try:
            avg_age = pd.to_numeric(detail_df["Delivery Age (Days)"], errors="coerce").mean()
            m4.metric("Avg Delivery Age (Days)", f"{avg_age:.1f}" if pd.notna(avg_age) else "N/A")
        except Exception:
            m4.metric("Avg Delivery Age (Days)", "N/A")

    st.markdown("---")

    st.markdown(
        "<div class='sec-title'>&#128205; Zone-wise Open Deliveries Summary</div>",
        unsafe_allow_html=True,
    )
    if not zone_summary.empty:
        _render_html_table(
            zone_summary,
            col_labels={"Plants": "Plants Affected"},
            max_height=360,
        )

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#127981; Pivot View (Zone &#8594; Plant &#8594; Delivery)</div>",
        unsafe_allow_html=True,
    )

    detail_cols = [
        "Zone Name", "Plant Name", "Delivery", "Volume",
        "Goods Issue Date", "Delivery Age (Days)",
    ]
    detail_cols = [c for c in detail_cols if c in detail_df.columns]

    all_zones = sorted(summary_df["Zone Name"].dropna().unique().tolist())
    tabs = st.tabs(all_zones) if all_zones else []
    for tab, zone in zip(tabs, all_zones):
        with tab:
            zone_rows = detail_df[detail_df["Zone Name"] == zone].copy()
            if zone_rows.empty:
                st.info("No records in this zone.")
                continue

            plant_summary = (
                zone_rows.groupby("Plant Name", dropna=False)
                .agg(open_deliveries=("Delivery", "nunique"))
                .reset_index()
                .rename(columns={"open_deliveries": "Open Delivery Count"})
                .sort_values("Open Delivery Count", ascending=False)
            )
            _render_html_table(
                plant_summary,
                col_labels={"Plant Name": "Plant"},
                max_height=260,
            )

            for _, prow in plant_summary.iterrows():
                plant = prow["Plant Name"]
                sort_cols = [c for c in ["Delivery Age (Days)", "Delivery"] if c in detail_cols]
                sort_asc  = [False if c == "Delivery Age (Days)" else True for c in sort_cols]
                p_df = zone_rows[zone_rows["Plant Name"] == plant][detail_cols]
                if sort_cols:
                    p_df = p_df.sort_values(sort_cols, ascending=sort_asc)
                p_df = p_df.reset_index(drop=True)
                with st.expander(f"{plant}  |  {len(p_df)} delivery record(s)", expanded=False):
                    _render_html_table(p_df, max_height=340)

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#128270; Sortable Open Delivery Records</div>",
        unsafe_allow_html=True,
    )
    sortable_df = detail_df[detail_cols].copy() if detail_cols else detail_df.copy()
    if not sortable_df.empty:
        if "Delivery Age (Days)" in sortable_df.columns:
            sortable_df["Delivery Age (Days)"] = pd.to_numeric(
                sortable_df["Delivery Age (Days)"], errors="coerce"
            )
        _render_streamlit_dataframe(sortable_df, max_height=420, hide_index=True)

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        xlsx_bytes = export_to_excel({
            "Zone Summary"    : zone_summary,
            "Plant Summary"   : summary_df,
            "Detailed Records": detail_df if not detail_df.empty else pd.DataFrame(),
        })
        st.download_button(
            label     = "&#11015;  Download Open Delivery Report  (.xlsx)",
            data      = xlsx_bytes,
            file_name = f"OpenDelivery_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key       = "dl_open_delivery",
        )


def render_open_intransit_details(
    open_intransit_result: dict,
    zone_filter          : list,
    plant_filter         : list,
) -> None:
    """Drill-down detail page for Open In-Transit STOs."""
    render_header(subtitle="&#128699; Open In-Transit &#8212; Drill Down")

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key="btn_back_open_intransit"):
            st.session_state["page"] = "dashboard"
            st.rerun()

    st.markdown("""
    <div class="detail-hdr">
        <h3>&#128699; Open In-Transit &#8212; Detailed Exception View</h3>
        <p>Pivot-style Zone and Plant grouping for open STO in-transit transactions.</p>
    </div>
    """, unsafe_allow_html=True)

    summary_df   = open_intransit_result.get("summary_df", pd.DataFrame())
    zone_summary = open_intransit_result.get("zone_summary", pd.DataFrame())
    detail_df    = open_intransit_result.get("detail_df", pd.DataFrame())
    total_count  = int(open_intransit_result.get("total_count", 0) or 0)

    if summary_df.empty:
        st.info("&#8505; No Open In-Transit data available for the current filter selection.")
        return

    total_zones  = int(summary_df["Zone Name"].nunique())
    total_plants = int(summary_df["Plant Name"].nunique())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Open In-Transit STO (Total)", f"{total_count:,}")
    m2.metric("Zones Affected", f"{total_zones}")
    m3.metric("Plants Affected", f"{total_plants}")
    if not detail_df.empty and "In-Transit Age (Days)" in detail_df.columns:
        try:
            avg_age = pd.to_numeric(detail_df["In-Transit Age (Days)"], errors="coerce").mean()
            m4.metric("Avg In-Transit Age (Days)", f"{avg_age:.1f}" if pd.notna(avg_age) else "N/A")
        except Exception:
            m4.metric("Avg In-Transit Age (Days)", "N/A")

    st.markdown("---")

    st.markdown(
        "<div class='sec-title'>&#128205; Zone-wise Open In-Transit Summary</div>",
        unsafe_allow_html=True,
    )
    if not zone_summary.empty:
        _render_html_table(
            zone_summary,
            col_labels={"Plants": "Plants Affected"},
            max_height=360,
        )

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#127981; Pivot View (Zone &#8594; Plant &#8594; STO Order)</div>",
        unsafe_allow_html=True,
    )

    detail_cols = [
        "Zone Name", "Plant Name", "STO Order", "Receiving Plant", "Dispatch Date",
        "Inco Terms", "Delivery", "Shipment", "Invoice", "Net Value",
        "Material", "Material Description", "Load Quantity", "Open Quantity",
        "In-Transit Age (Days)",
    ]
    detail_cols = [c for c in detail_cols if c in detail_df.columns]

    all_zones = sorted(summary_df["Zone Name"].dropna().unique().tolist())
    tabs = st.tabs(all_zones) if all_zones else []
    for tab, zone in zip(tabs, all_zones):
        with tab:
            zone_rows = detail_df[detail_df["Zone Name"] == zone].copy()
            if zone_rows.empty:
                st.info("No records in this zone.")
                continue

            plant_summary = (
                zone_rows.groupby("Plant Name", dropna=False)
                .agg(open_intransit_sto=("STO Order", "nunique"))
                .reset_index()
                .rename(columns={"open_intransit_sto": "Open In-Transit STO Count"})
                .sort_values("Open In-Transit STO Count", ascending=False)
            )
            _render_html_table(
                plant_summary,
                col_labels={"Plant Name": "Plant"},
                max_height=260,
            )

            for _, prow in plant_summary.iterrows():
                plant = prow["Plant Name"]
                sort_cols = [c for c in ["In-Transit Age (Days)", "STO Order"] if c in detail_cols]
                sort_asc  = [False if c == "In-Transit Age (Days)" else True for c in sort_cols]
                p_df = zone_rows[zone_rows["Plant Name"] == plant][detail_cols]
                if sort_cols:
                    p_df = p_df.sort_values(sort_cols, ascending=sort_asc)
                p_df = p_df.reset_index(drop=True)
                with st.expander(f"{plant}  |  {len(p_df)} record(s)", expanded=False):
                    _render_html_table(p_df, max_height=360)

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#128270; Sortable Open In-Transit Records</div>",
        unsafe_allow_html=True,
    )
    sortable_df = detail_df[detail_cols].copy() if detail_cols else detail_df.copy()
    if not sortable_df.empty:
        if "In-Transit Age (Days)" in sortable_df.columns:
            sortable_df["In-Transit Age (Days)"] = pd.to_numeric(
                sortable_df["In-Transit Age (Days)"], errors="coerce"
            )
        _render_streamlit_dataframe(sortable_df, max_height=420, hide_index=True)

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        xlsx_bytes = export_to_excel({
            "Zone Summary"    : zone_summary,
            "Plant Summary"   : summary_df,
            "Detailed Records": detail_df if not detail_df.empty else pd.DataFrame(),
        })
        st.download_button(
            label     = "&#11015;  Download Open In-Transit Report  (.xlsx)",
            data      = xlsx_bytes,
            file_name = f"OpenInTransit_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key       = "dl_open_intransit",
        )


def render_open_sales_orders_details(
    open_sales_orders_result: dict,
    zone_filter             : list,
    plant_filter            : list,
) -> None:
    """Drill-down detail page for Open Sales Orders."""
    render_header(subtitle="&#128203; Open Sales Orders &#8212; Drill Down")

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key="btn_back_open_so"):
            st.session_state["page"] = "dashboard"
            st.rerun()

    st.markdown("""
    <div class="detail-hdr">
        <h3>&#128203; Open Sales Orders &#8212; Detailed Exception View</h3>
        <p>Pivot-style Zone and Plant grouping for open sales orders.</p>
    </div>
    """, unsafe_allow_html=True)

    summary_df   = open_sales_orders_result.get("summary_df", pd.DataFrame())
    zone_summary = open_sales_orders_result.get("zone_summary", pd.DataFrame())
    detail_df    = open_sales_orders_result.get("detail_df", pd.DataFrame())
    total_count  = int(open_sales_orders_result.get("total_count", 0) or 0)

    if summary_df.empty:
        st.info("&#8505; No Open Sales Order data available for the current filter selection.")
        return

    total_zones  = int(summary_df["Zone Name"].nunique())
    total_plants = int(summary_df["Plant Name"].nunique())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Open Sales Orders (Total)", f"{total_count:,}")
    m2.metric("Zones Affected", f"{total_zones}")
    m3.metric("Plants Affected", f"{total_plants}")
    if not detail_df.empty and "Sales Order Age (Days)" in detail_df.columns:
        try:
            avg_age = pd.to_numeric(detail_df["Sales Order Age (Days)"], errors="coerce").mean()
            m4.metric("Avg Sales Order Age (Days)", f"{avg_age:.1f}" if pd.notna(avg_age) else "N/A")
        except Exception:
            m4.metric("Avg Sales Order Age (Days)", "N/A")

    st.markdown("---")

    st.markdown(
        "<div class='sec-title'>&#128205; Zone-wise Open Sales Order Summary</div>",
        unsafe_allow_html=True,
    )
    if not zone_summary.empty:
        _render_html_table(
            zone_summary,
            col_labels={"Plants": "Plants Affected"},
            max_height=360,
        )

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#127981; Pivot View (Zone &#8594; Plant &#8594; Sales Document)</div>",
        unsafe_allow_html=True,
    )

    detail_cols = [
        "Zone Name", "Plant Name", "Sales Document", "Sales Document Type",
        "Sold-to Party", "Sold-to Party Name", "Material", "Material Description",
        "Order Quantity (Item)", "Sales Unit", "Document Date", "Net Value (Item)",
        "Shipping Point/Receiving Pt", "Confirmed Quantity (Item)", "Sales Order Age (Days)",
    ]
    detail_cols = [c for c in detail_cols if c in detail_df.columns]

    all_zones = sorted(summary_df["Zone Name"].dropna().unique().tolist())
    tabs = st.tabs(all_zones) if all_zones else []
    for tab, zone in zip(tabs, all_zones):
        with tab:
            zone_rows = detail_df[detail_df["Zone Name"] == zone].copy()
            if zone_rows.empty:
                st.info("No records in this zone.")
                continue

            plant_summary = (
                zone_rows.groupby("Plant Name", dropna=False)
                .agg(open_so=("Sales Document", "nunique"))
                .reset_index()
                .rename(columns={"open_so": "Open Sales Order Count"})
                .sort_values("Open Sales Order Count", ascending=False)
            )
            _render_html_table(
                plant_summary,
                col_labels={"Plant Name": "Plant"},
                max_height=260,
            )

            for _, prow in plant_summary.iterrows():
                plant = prow["Plant Name"]
                sort_cols = [c for c in ["Sales Order Age (Days)", "Sales Document"] if c in detail_cols]
                sort_asc  = [False if c == "Sales Order Age (Days)" else True for c in sort_cols]
                p_df = zone_rows[zone_rows["Plant Name"] == plant][detail_cols]
                if sort_cols:
                    p_df = p_df.sort_values(sort_cols, ascending=sort_asc)
                p_df = p_df.reset_index(drop=True)
                with st.expander(f"{plant}  |  {len(p_df)} record(s)", expanded=False):
                    _render_html_table(p_df, max_height=360)

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#128270; Sortable Open Sales Order Records</div>",
        unsafe_allow_html=True,
    )
    sortable_df = detail_df[detail_cols].copy() if detail_cols else detail_df.copy()
    if not sortable_df.empty:
        if "Sales Order Age (Days)" in sortable_df.columns:
            sortable_df["Sales Order Age (Days)"] = pd.to_numeric(
                sortable_df["Sales Order Age (Days)"], errors="coerce"
            )
        _render_streamlit_dataframe(sortable_df, max_height=420, hide_index=True)

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        xlsx_bytes = export_to_excel({
            "Zone Summary"    : zone_summary,
            "Plant Summary"   : summary_df,
            "Detailed Records": detail_df if not detail_df.empty else pd.DataFrame(),
        })
        st.download_button(
            label     = "&#11015;  Download Open Sales Order Report  (.xlsx)",
            data      = xlsx_bytes,
            file_name = f"OpenSalesOrder_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key       = "dl_open_so",
        )


def render_pending_invoices_details(
    pending_invoices_result: dict,
    zone_filter           : list,
    plant_filter          : list,
) -> None:
    """Drill-down detail page for Pending Invoices."""
    render_header(subtitle="&#129534; Pending Invoices &#8212; Drill Down")

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key="btn_back_pending_inv"):
            st.session_state["page"] = "dashboard"
            st.rerun()

    st.markdown("""
    <div class="detail-hdr">
        <h3>&#129534; Pending Invoices &#8212; Detailed Exception View</h3>
        <p>Pivot-style Zone and Plant grouping for pending invoice transactions.</p>
    </div>
    """, unsafe_allow_html=True)

    summary_df   = pending_invoices_result.get("summary_df", pd.DataFrame())
    zone_summary = pending_invoices_result.get("zone_summary", pd.DataFrame())
    detail_df    = pending_invoices_result.get("detail_df", pd.DataFrame())
    total_count  = int(pending_invoices_result.get("total_count", 0) or 0)

    if summary_df.empty:
        st.info("&#8505; No Pending Invoice data available for the current filter selection.")
        return

    total_zones  = int(summary_df["Zone Name"].nunique())
    total_plants = int(summary_df["Plant Name"].nunique())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Pending Invoices (Total)", f"{total_count:,}")
    m2.metric("Zones Affected", f"{total_zones}")
    m3.metric("Plants Affected", f"{total_plants}")
    if not detail_df.empty and "Invoice Age (Days)" in detail_df.columns:
        try:
            avg_age = pd.to_numeric(detail_df["Invoice Age (Days)"], errors="coerce").mean()
            m4.metric("Avg Invoice Age (Days)", f"{avg_age:.1f}" if pd.notna(avg_age) else "N/A")
        except Exception:
            m4.metric("Avg Invoice Age (Days)", "N/A")

    st.markdown("---")

    st.markdown(
        "<div class='sec-title'>&#128205; Zone-wise Pending Invoice Summary</div>",
        unsafe_allow_html=True,
    )
    if not zone_summary.empty:
        _render_html_table(
            zone_summary,
            col_labels={"Plants": "Plants Affected"},
            max_height=360,
        )

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#127981; Pivot View (Zone &#8594; Plant &#8594; Delivery)</div>",
        unsafe_allow_html=True,
    )

    detail_cols = [
        "Zone Name", "Plant Name", "Sending Location", "Receiving Location", "MOT",
        "Purchase Order", "TD Shipment", "Delivery", "Material Document", "Quantity",
        "Created By", "Description", "Created Date", "Invoice Age (Days)",
    ]
    detail_cols = [c for c in detail_cols if c in detail_df.columns]

    all_zones = sorted(summary_df["Zone Name"].dropna().unique().tolist())
    tabs = st.tabs(all_zones) if all_zones else []
    for tab, zone in zip(tabs, all_zones):
        with tab:
            zone_rows = detail_df[detail_df["Zone Name"] == zone].copy()
            if zone_rows.empty:
                st.info("No records in this zone.")
                continue

            plant_summary = (
                zone_rows.groupby("Plant Name", dropna=False)
                .agg(pending_invoices=("Delivery", "nunique"))
                .reset_index()
                .rename(columns={"pending_invoices": "Pending Invoice Count"})
                .sort_values("Pending Invoice Count", ascending=False)
            )
            _render_html_table(
                plant_summary,
                col_labels={"Plant Name": "Plant"},
                max_height=260,
            )

            for _, prow in plant_summary.iterrows():
                plant = prow["Plant Name"]
                sort_cols = [c for c in ["Invoice Age (Days)", "Delivery"] if c in detail_cols]
                sort_asc  = [False if c == "Invoice Age (Days)" else True for c in sort_cols]
                p_df = zone_rows[zone_rows["Plant Name"] == plant][detail_cols]
                if sort_cols:
                    p_df = p_df.sort_values(sort_cols, ascending=sort_asc)
                p_df = p_df.reset_index(drop=True)
                with st.expander(f"{plant}  |  {len(p_df)} record(s)", expanded=False):
                    _render_html_table(p_df, max_height=360)

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#128270; Sortable Pending Invoice Records</div>",
        unsafe_allow_html=True,
    )
    sortable_df = detail_df[detail_cols].copy() if detail_cols else detail_df.copy()
    if not sortable_df.empty:
        if "Invoice Age (Days)" in sortable_df.columns:
            sortable_df["Invoice Age (Days)"] = pd.to_numeric(
                sortable_df["Invoice Age (Days)"], errors="coerce"
            )
        _render_streamlit_dataframe(sortable_df, max_height=420, hide_index=True)

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        xlsx_bytes = export_to_excel({
            "Zone Summary"    : zone_summary,
            "Plant Summary"   : summary_df,
            "Detailed Records": detail_df if not detail_df.empty else pd.DataFrame(),
        })
        st.download_button(
            label     = "&#11015;  Download Pending Invoice Report  (.xlsx)",
            data      = xlsx_bytes,
            file_name = f"PendingInvoice_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key       = "dl_pending_inv",
        )


def render_dummy_tank_details(
    dummy_tank_df: pd.DataFrame,
    total_dummy_qty: float,
    error_message: str = "",
    missing_columns: list | None = None,
) -> None:
    """Drill-down detail page for Dummy Tank Quantity."""
    st.markdown("<div class='sec-title'>&#128736; Dummy Tank Qty &#8212; Drill Down</div>", unsafe_allow_html=True)

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#11013; Back to Dashboard", key="btn_back_dummy_tank"):
            st.session_state["pl_unblock_clicked"] = False
            st.session_state["dummy_tank_clicked"] = False
            st.session_state["tank_turns_page"] = "main"
            st.rerun()

    if error_message:
        st.warning(error_message)
        return

    if missing_columns:
        st.warning("Missing required column(s): " + ", ".join(missing_columns))
        return

    if dummy_tank_df is None or dummy_tank_df.empty:
        st.info("No Dummy Tank data available after applying filters.")
        return

    if "Zone" not in dummy_tank_df.columns and "Zone Name" in dummy_tank_df.columns:
        dummy_tank_df = dummy_tank_df.copy()
        dummy_tank_df["Zone"] = dummy_tank_df["Zone Name"]

    display_cols = ["Material", "Plant", "Zone", "Storage Location", "Base Unit of Measure", "Unrestricted"]
    display_cols = [c for c in display_cols if c in dummy_tank_df.columns]
    display_df = dummy_tank_df[display_cols].copy()
    if "Unrestricted" in display_df.columns:
        display_df["Unrestricted"] = pd.to_numeric(display_df["Unrestricted"], errors="coerce").fillna(0).map(lambda v: f"{v:,.0f}")

    st.metric("Total Dummy Tank Quantity", f"{total_dummy_qty:,.0f}")
    st.markdown("<div class='sec-title'>&#128203; Dummy Tank Details</div>", unsafe_allow_html=True)
    _render_html_table(display_df, max_height=420)

    _download_excel_button(
        label="&#11015;  Download Dummy Tank Raw Data  (.xlsx)",
        file_prefix="dummy_tank_stock",
        sheets={"Dummy_Tank_Details": dummy_tank_df[display_cols].copy()},
        key="dl_dummy_tank",
    )

    top5_plants = pd.DataFrame()
    if "Plant" in dummy_tank_df.columns and "Unrestricted" in dummy_tank_df.columns:
        top5_plants = (
            dummy_tank_df.groupby("Plant", dropna=False, as_index=False)["Unrestricted"]
            .sum()
            .sort_values("Unrestricted", ascending=False)
            .head(5)
        )

    top5_zones = pd.DataFrame()
    if "Zone" in dummy_tank_df.columns and "Unrestricted" in dummy_tank_df.columns:
        zone_series = dummy_tank_df["Zone"].fillna("").astype(str).str.strip()
        if zone_series.ne("").any():
            zone_df = dummy_tank_df.copy()
            zone_df["Zone"] = zone_series
            zone_df = zone_df[zone_df["Zone"] != ""]
            top5_zones = (
                zone_df.groupby("Zone", dropna=False, as_index=False)["Unrestricted"]
                .sum()
                .sort_values("Unrestricted", ascending=False)
                .head(5)
            )

    st.markdown("<div class='sec-title'>&#128200; Top 5 Analysis</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("Top 5 Plants")
        if not top5_plants.empty:
            top5_plants_display = top5_plants.copy()
            top5_plants_display["Unrestricted"] = top5_plants_display["Unrestricted"].map(lambda v: f"{v:,.3f}")
            _render_html_table(top5_plants_display, max_height=280)
        else:
            st.info("Plant-wise data not available.")

    with col2:
        st.markdown("Top 5 Zones")
        if not top5_zones.empty:
            top5_zones_display = top5_zones.copy()
            top5_zones_display["Unrestricted"] = top5_zones_display["Unrestricted"].map(lambda v: f"{v:,.3f}")
            _render_html_table(top5_zones_display, max_height=280)
        else:
            st.info("Zone column not available in source file.")

    charts_col1, charts_col2 = st.columns(2)
    with charts_col1:
        if not top5_plants.empty:
            donut = px.pie(
                top5_plants,
                names="Plant",
                values="Unrestricted",
                hole=0.55,
                title="Top 5 Plants by Unrestricted (Donut)",
            )
            st.plotly_chart(donut, use_container_width=True)

    with charts_col2:
        if not top5_plants.empty:
            bar_plants = px.bar(
                top5_plants,
                x="Plant",
                y="Unrestricted",
                title="Top 5 Plants by Unrestricted",
            )
            st.plotly_chart(bar_plants, use_container_width=True)

    if not top5_zones.empty:
        bar_zones = px.bar(
            top5_zones,
            x="Zone",
            y="Unrestricted",
            title="Top 5 Zones by Unrestricted",
        )
        st.plotly_chart(bar_zones, use_container_width=True)


def render_pl_unblock_details(
    pl_unblock_df: pd.DataFrame,
    total_pl_unblock_qty: float,
    error_message: str = "",
    missing_columns: list | None = None,
) -> None:
    """Drill-down detail page for PL Unblock Quantity."""
    st.markdown("<div class='sec-title'>&#128295; PL Unblock Qty &#8212; Drill Down</div>", unsafe_allow_html=True)

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#11013; Back to Dashboard", key="btn_back_pl_unblock"):
            st.session_state["dummy_tank_clicked"] = False
            st.session_state["pl_unblock_clicked"] = False
            st.session_state["tank_turns_page"] = "main"
            st.rerun()

    if error_message:
        st.warning(error_message)
        return

    if missing_columns:
        st.warning("Missing required column(s): " + ", ".join(missing_columns))
        return

    if pl_unblock_df is None or pl_unblock_df.empty:
        st.info("No PL Unblock data available after applying filters.")
        return

    if "Zone" not in pl_unblock_df.columns and "Zone Name" in pl_unblock_df.columns:
        pl_unblock_df = pl_unblock_df.copy()
        pl_unblock_df["Zone"] = pl_unblock_df["Zone Name"]

    display_cols = ["Material", "Plant", "Zone", "Storage location", "Base Unit of Measure", "Unrestricted", "Blocked"]
    display_cols = [c for c in display_cols if c in pl_unblock_df.columns]
    display_df = pl_unblock_df[display_cols].copy()

    st.metric("Total Pipeline Unblock Quantity", f"{total_pl_unblock_qty:,.0f}")
    st.markdown("<div class='sec-title'>&#128203; PL Unblock Details</div>", unsafe_allow_html=True)

    if not display_df.empty:
        display_df = display_df.copy()
        for num_col in ["Unrestricted", "Blocked"]:
            if num_col in display_df.columns:
                display_df[num_col] = pd.to_numeric(display_df[num_col], errors="coerce").fillna(0).map(lambda v: f"{v:,.0f}")
        _render_html_table(display_df, max_height=420)

    _download_excel_button(
        label="&#11015;  Download PL Unblock Raw Data  (.xlsx)",
        file_prefix="pl_unblock_stock",
        sheets={"PL_Unblock_Details": pl_unblock_df[display_cols].copy()},
        key="dl_pl_unblock",
    )

    top3_plants = pd.DataFrame()
    if "Plant" in pl_unblock_df.columns and "Unrestricted" in pl_unblock_df.columns:
        top3_plants = (
            pl_unblock_df.groupby("Plant", dropna=False, as_index=False)["Unrestricted"]
            .sum()
            .sort_values("Unrestricted", ascending=False)
            .head(3)
        )

    top3_zones = pd.DataFrame()
    if "Zone" in pl_unblock_df.columns and "Unrestricted" in pl_unblock_df.columns:
        zone_series = pl_unblock_df["Zone"].fillna("").astype(str).str.strip()
        if zone_series.ne("").any():
            zone_df = pl_unblock_df.copy()
            zone_df["Zone"] = zone_series
            zone_df = zone_df[zone_df["Zone"] != ""]
            top3_zones = (
                zone_df.groupby("Zone", dropna=False, as_index=False)["Unrestricted"]
                .sum()
                .sort_values("Unrestricted", ascending=False)
                .head(3)
            )

    st.markdown("<div class='sec-title'>&#128200; Top 3 Analysis</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("Top 3 Plants")
        if not top3_plants.empty:
            top3_plants_table = top3_plants.copy()
            top3_plants_table["Unrestricted"] = top3_plants_table["Unrestricted"].map(lambda v: f"{v:,.3f}")
            _render_html_table(top3_plants_table, max_height=280)
        else:
            st.info("Plant-wise data not available.")

    with col2:
        st.markdown("Top 3 Zones")
        if not top3_zones.empty:
            top3_zones_table = top3_zones.copy()
            top3_zones_table["Unrestricted"] = top3_zones_table["Unrestricted"].map(lambda v: f"{v:,.3f}")
            _render_html_table(top3_zones_table, max_height=280)
        else:
            st.info("Zone column not available in source file.")

    charts_col1, charts_col2 = st.columns(2)
    with charts_col1:
        if not top3_plants.empty:
            donut = px.pie(
                top3_plants,
                names="Plant",
                values="Unrestricted",
                hole=0.55,
                title="Top 3 Plants by Unrestricted (Donut)",
            )
            st.plotly_chart(donut, use_container_width=True)

    with charts_col2:
        if not top3_plants.empty:
            bar_plants = px.bar(
                top3_plants,
                x="Plant",
                y="Unrestricted",
                title="Top 3 Plants by Unrestricted",
            )
            st.plotly_chart(bar_plants, use_container_width=True)

    if not top3_zones.empty:
        bar_zones = px.bar(
            top3_zones,
            x="Zone",
            y="Unrestricted",
            title="Top 3 Zones by Unrestricted",
        )
        st.plotly_chart(bar_zones, use_container_width=True)


def render_tank_turns_details(
    tt_df          : "pd.DataFrame",
    tank_turns_val : float,
    error_message  : str  = "",
    missing_columns: list | None = None,
) -> None:
    """Drill-down detail page for Tank Turns KPI."""
    import plotly.express as px

    st.markdown("<div class='sec-title'>&#128167; Tank Turns &#8212; Drill Down</div>",
                unsafe_allow_html=True)

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#11013; Back to Dashboard", key="btn_back_tank_turns"):
            st.session_state["tank_turns_page"] = "main"
            st.session_state["dummy_tank_clicked"] = False
            st.session_state["pl_unblock_clicked"] = False
            st.rerun()

    if error_message:
        st.warning(error_message)
        return
    if missing_columns:
        st.warning("Missing required column(s): " + ", ".join(missing_columns))
        return
    if tt_df is None or tt_df.empty:
        st.info("No Tank Turns data available after applying filters.")
        return

    st.metric("Tank Turns (Dispatches / Capacity)", f"{tank_turns_val:.2f}")

    # ── Drill-down filters ────────────────────────────────────────────────────
    st.markdown("<div class='sec-title'>&#128269; Drill-Down Filters</div>",
                unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns(4)

    def _sorted_opts(series):
        return ["All"] + sorted(series.dropna().astype(str).unique().tolist())

    sel_zone_dd = f1.selectbox(
        "Zone",
        _sorted_opts(tt_df.get("Zone Name", tt_df.get("Zone", pd.Series(dtype=str)))),
        key="tt_filter_zone",
    )
    sel_mat_desc = f2.selectbox(
        "Material Description",
        _sorted_opts(tt_df.get("Material Description", pd.Series(dtype=str))),
        key="tt_filter_mat_desc",
    )
    sel_mat = f3.selectbox(
        "Material",
        _sorted_opts(tt_df.get("Material", pd.Series(dtype=str))),
        key="tt_filter_mat",
    )
    sel_plant_name = f4.selectbox(
        "Plant Name",
        _sorted_opts(tt_df.get("Plant Name_master",
                               tt_df.get("Plant Name", pd.Series(dtype=str)))),
        key="tt_filter_plant",
    )

    # Apply local filters
    df_view = tt_df.copy()
    zone_col = "Zone Name" if "Zone Name" in df_view.columns else "Zone"
    pm_col   = "Plant Name_master" if "Plant Name_master" in df_view.columns else "Plant Name"

    if sel_zone_dd != "All" and zone_col in df_view.columns:
        df_view = df_view[df_view[zone_col].astype(str) == sel_zone_dd]
    if sel_mat_desc != "All" and "Material Description" in df_view.columns:
        df_view = df_view[df_view["Material Description"].astype(str) == sel_mat_desc]
    if sel_mat != "All" and "Material" in df_view.columns:
        df_view = df_view[df_view["Material"].astype(str) == sel_mat]
    if sel_plant_name != "All" and pm_col in df_view.columns:
        df_view = df_view[df_view[pm_col].astype(str) == sel_plant_name]

    # Recompute KPI for filtered view
    flt_dispatch  = pd.to_numeric(df_view["Dispatches"],    errors="coerce").fillna(0).sum()
    flt_capacity  = pd.to_numeric(df_view["Tank Capacity"], errors="coerce").fillna(0).sum()
    flt_turns_val = (flt_dispatch / flt_capacity) if flt_capacity != 0 else 0.0
    st.metric("Filtered Tank Turns", f"{flt_turns_val:.2f}")

    # ── Pivot Table ───────────────────────────────────────────────────────────
    st.markdown("<div class='sec-title'>&#128203; Tank Turns Details (Pivot)</div>",
                unsafe_allow_html=True)

    pivot_index = [c for c in ["Plant", "Unique Ref Id", "Tank",
                               "Tank Capacity", "Tank Type", "Tank Status"]
                  if c in df_view.columns]
    pivot_values = {c: v for c, v in [
        ("Opening Stock", "sum"), ("Receipts", "sum"),
        ("Dispatches", "sum"), ("Closing Stock", "sum"), ("Turn", "mean"),
    ] if c in df_view.columns}

    if pivot_index and pivot_values:
        try:
            pivot = df_view.groupby(pivot_index, dropna=False).agg(pivot_values).reset_index()
            pivot.rename(columns={
                "Opening Stock": "Sum of Opening Stock",
                "Receipts"     : "Sum of Receipts",
                "Dispatches"   : "Sum of Dispatches",
                "Closing Stock": "Sum of Closing Stock",
                "Turn"         : "Average of Turn",
            }, inplace=True)

            # Format numeric display columns
            for nc in ["Sum of Opening Stock", "Sum of Receipts",
                       "Sum of Dispatches", "Sum of Closing Stock"]:
                if nc in pivot.columns:
                    pivot[nc] = pd.to_numeric(pivot[nc], errors="coerce").fillna(0).map(lambda v: f"{v:,.0f}")
            if "Average of Turn" in pivot.columns:
                pivot["Average of Turn"] = pd.to_numeric(
                    pivot["Average of Turn"], errors="coerce").fillna(0).map(lambda v: f"{v:.2f}")
            if "Tank Capacity" in pivot.columns:
                pivot["Tank Capacity"] = pd.to_numeric(
                    pivot["Tank Capacity"], errors="coerce").fillna(0).map(lambda v: f"{v:,.0f}")

            _render_html_table(pivot, max_height=480)
        except Exception as e:
            st.error(f"Pivot build error: {e}")
    else:
        st.info("Insufficient columns for pivot table.")

    _download_excel_button(
        label="&#11015;  Download Tank Turns Raw Data  (.xlsx)",
        file_prefix="tank_turns",
        sheets={"Tank_Turns_Raw": df_view.reset_index(drop=True)},
        key="dl_tank_turns",
    )

    # ── Top Analysis ──────────────────────────────────────────────────────────
    st.markdown("<div class='sec-title'>&#128200; Top Analysis</div>",
                unsafe_allow_html=True)

    df_num = df_view.copy()
    df_num["Dispatches"]    = pd.to_numeric(df_num["Dispatches"],    errors="coerce").fillna(0)
    df_num["Tank Capacity"] = pd.to_numeric(df_num["Tank Capacity"], errors="coerce").fillna(0)

    top3_plants = pd.DataFrame()
    if "Plant" in df_num.columns and "Turn" in df_num.columns:
        _grp_cols = ["Plant"] + (["Plant Name"] if "Plant Name" in df_num.columns else [])
        top3_plants = (
            df_num.groupby(_grp_cols, dropna=False, as_index=False)["Turn"]
            .mean().rename(columns={"Turn": "Average Turn"})
            .sort_values("Average Turn", ascending=False).head(3)
        )

    top3_zones = pd.DataFrame()
    zone_col_disp = "Zone Name" if "Zone Name" in df_num.columns else ("Zone" if "Zone" in df_num.columns else None)
    if zone_col_disp and "Turn" in df_num.columns:
        z_series = df_num[zone_col_disp].fillna("").astype(str).str.strip()
        if z_series.ne("").any():
            tmp = df_num.copy(); tmp["_zone"] = z_series
            tmp = tmp[tmp["_zone"] != ""]
            top3_zones = (
                tmp.groupby("_zone", dropna=False, as_index=False)["Turn"]
                .mean().rename(columns={"Turn": "Average Turn", "_zone": "Zone"})
                .sort_values("Average Turn", ascending=False).head(3)
                .rename(columns={"_zone": "Zone"})
            )

    ca, cb = st.columns(2)
    with ca:
        st.markdown("Top 3 Plants")
        if not top3_plants.empty:
            tp_disp = top3_plants.copy()
            tp_disp["Average Turn"] = tp_disp["Average Turn"].map(lambda v: f"{v:,.2f}")
            _cols_order = ["Plant"] + (["Plant Name"] if "Plant Name" in tp_disp.columns else []) + ["Average Turn"]
            _render_html_table(tp_disp[_cols_order], max_height=220)
        else:
            st.info("Plant data not available.")
    with cb:
        st.markdown("Top 3 Zones")
        if not top3_zones.empty:
            tz_disp = top3_zones.copy()
            tz_disp["Average Turn"] = tz_disp["Average Turn"].map(lambda v: f"{v:,.2f}")
            _render_html_table(tz_disp, max_height=220)
        else:
            st.info("Zone data not available.")

    # ── Charts ────────────────────────────────────────────────────────────────
    ch1, ch2 = st.columns(2)
    with ch1:
        if not top3_plants.empty:
            donut = px.pie(top3_plants, names="Plant", values="Average Turn",
                           hole=0.55, title="Top 3 Plants by Average Turn (Donut)")
            st.plotly_chart(donut, use_container_width=True)
    with ch2:
        if not top3_plants.empty:
            bar_p = px.bar(top3_plants, x="Plant", y="Average Turn",
                           title="Top 3 Plants by Average Turn")
            st.plotly_chart(bar_p, use_container_width=True)

    if not top3_zones.empty:
        bar_z = px.bar(top3_zones, x="Zone", y="Average Turn",
                       title="Top 3 Zones by Average Turn")
        st.plotly_chart(bar_z, use_container_width=True)


def render_tank_reco_details(
    tank_reco_result: dict,
    zone_filter     : list,
    plant_filter    : list,
) -> None:
    """Drill-down detail page for Tank Reco exceptions."""
    render_header(subtitle="&#128738; Tank Reco &#8212; Drill Down")

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key="btn_back_tank_reco"):
            st.session_state["page"] = "dashboard"
            st.rerun()

    st.markdown("""
    <div class="detail-hdr">
        <h3>&#128738; Tank Reco &#8212; Detailed Exception View</h3>
        <p>Unique exceptions counted as Plant + Tank + Material combinations.</p>
    </div>
    """, unsafe_allow_html=True)

    summary_df   = tank_reco_result.get("summary_df", pd.DataFrame())
    zone_summary = tank_reco_result.get("zone_summary", pd.DataFrame())
    detail_df    = tank_reco_result.get("detail_df", pd.DataFrame())
    total_count  = int(tank_reco_result.get("total_count", 0) or 0)

    if summary_df.empty:
        st.info("&#8505; No Tank Reco data available for the current filter selection.")
        return

    total_zones  = int(summary_df["Zone Name"].nunique())
    total_plants = int(summary_df["Plant Name"].nunique())

    approved_count = 0
    if not detail_df.empty and "Reco Status" in detail_df.columns:
        approved_count = int(
            detail_df["Reco Status"]
            .astype(str)
            .str.upper()
            .str.contains("APPROV", na=False)
            .sum()
        )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Tank Reco (Total)", f"{total_count:,}")
    m2.metric("Zones Affected", f"{total_zones}")
    m3.metric("Plants Affected", f"{total_plants}")
    m4.metric("Approved Reco", f"{approved_count:,}")

    st.markdown("---")

    st.markdown(
        "<div class='sec-title'>&#128205; Zone-wise Tank Reco Summary</div>",
        unsafe_allow_html=True,
    )
    if not zone_summary.empty:
        _render_html_table(
            zone_summary,
            col_labels={"Plants": "Plants Affected"},
            max_height=360,
        )

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#127981; Pivot View (Zone &#8594; Plant &#8594; Tank Reco)</div>",
        unsafe_allow_html=True,
    )

    detail_cols = [
        "Zone Name", "Plant Name", "Plant", "Tank No.", "Material Code", "Dip Type",
        "Reco Status", "Reco Initiator", "Physical Stock", "Book Stock @ Dip",
        "Book Stock @ Posting", "Gain/Loss Booked", "Type", "Posting Date",
        "Material Doc No", "Material Doc Year", "Reco Approver", "Approval Date",
        "Comments for Abnormal G/L", "Description of Reason", "Remarks for Manual Dip",
        "Dip Date", "Phy Inv Doc", "Tank Reco Key",
    ]
    detail_cols = [c for c in detail_cols if c in detail_df.columns]

    all_zones = sorted(summary_df["Zone Name"].dropna().unique().tolist())
    tabs = st.tabs(all_zones) if all_zones else []
    for tab, zone in zip(tabs, all_zones):
        with tab:
            zone_rows = detail_df[detail_df["Zone Name"] == zone].copy()
            if zone_rows.empty:
                st.info("No records in this zone.")
                continue

            plant_summary = (
                zone_rows.groupby("Plant Name", dropna=False)
                .agg(tank_reco_count=("Tank Reco Key", "nunique"))
                .reset_index()
                .rename(columns={"tank_reco_count": "Tank Reco Count"})
                .sort_values("Tank Reco Count", ascending=False)
            )
            _render_html_table(
                plant_summary,
                col_labels={"Plant Name": "Plant"},
                max_height=260,
            )

            for _, prow in plant_summary.iterrows():
                plant = prow["Plant Name"]
                sort_cols = [c for c in ["Posting Date", "Dip Date", "Tank No.", "Material Code"] if c in detail_cols]
                p_df = zone_rows[zone_rows["Plant Name"] == plant][detail_cols]
                if sort_cols:
                    p_df = p_df.sort_values(sort_cols, ascending=[False, False, True, True][:len(sort_cols)])
                p_df = p_df.reset_index(drop=True)
                with st.expander(f"{plant}  |  {len(p_df)} record(s)", expanded=False):
                    _render_html_table(p_df, max_height=360)

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#128270; Sortable Tank Reco Records</div>",
        unsafe_allow_html=True,
    )
    sortable_df = detail_df[detail_cols].copy() if detail_cols else detail_df.copy()
    if not sortable_df.empty:
        _render_html_table(sortable_df, max_height=420)

    unmatched = tank_reco_result.get("unmatched", [])
    if unmatched:
        with st.expander(
            f"&#9888; {len(unmatched)} Plant Code(s) not found in PlantMaster",
            expanded=False,
        ):
            st.warning(
                "The following Plant codes could not be mapped to PlantMaster: "
                + ", ".join(sorted(map(str, unmatched)))
            )

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        xlsx_bytes = export_to_excel({
            "Zone Summary"    : zone_summary,
            "Plant Summary"   : summary_df,
            "Detailed Records": detail_df if not detail_df.empty else pd.DataFrame(),
        })
        st.download_button(
            label     = "&#11015;  Download Tank Reco Report  (.xlsx)",
            data      = xlsx_bytes,
            file_name = f"TankReco_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key       = "dl_tank_reco",
        )


def render_open_shortages_sales_details(
    open_short_sales_result: dict,
    zone_filter           : list,
    plant_filter          : list,
) -> None:
    """Drill-down detail page for Open Shortages (Sales)."""
    render_header(subtitle="&#128202; Open Shortages (Sales) &#8212; Drill Down")

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key="btn_back_open_short_sales"):
            st.session_state["page"] = "dashboard"
            st.rerun()

    st.markdown("""
    <div class="detail-hdr">
        <h3>&#128202; Open Shortages (Sales) &#8212; Detailed Exception View</h3>
        <p>Pivot-style Zone and Plant grouping for shortage transactions.</p>
    </div>
    """, unsafe_allow_html=True)

    summary_df   = open_short_sales_result.get("summary_df", pd.DataFrame())
    zone_summary = open_short_sales_result.get("zone_summary", pd.DataFrame())
    detail_df    = open_short_sales_result.get("detail_df", pd.DataFrame())
    total_qty    = float(open_short_sales_result.get("total_count", 0) or 0)

    if summary_df.empty:
        st.info("&#8505; No Open Shortages (Sales) data available for the current filter selection.")
        return

    total_zones  = int(summary_df["Zone Name"].nunique())
    total_plants = int(summary_df["Plant Name"].nunique())
    avg_age = pd.NA
    if "Shortage Age (Days)" in detail_df.columns and not detail_df.empty:
        avg_age = pd.to_numeric(detail_df["Shortage Age (Days)"], errors="coerce").mean()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Shortage Quantity (Ltrs)", f"{total_qty:,.2f}")
    m2.metric("Zones Affected", f"{total_zones}")
    m3.metric("Plants Affected", f"{total_plants}")
    m4.metric("Avg Shortage Age (Days)", f"{avg_age:.1f}" if pd.notna(avg_age) else "N/A")

    st.markdown("---")

    st.markdown(
        "<div class='sec-title'>&#128205; Zone-wise Shortage Quantity Summary</div>",
        unsafe_allow_html=True,
    )
    if not zone_summary.empty:
        _render_html_table(
            zone_summary,
            col_labels={"Plants": "Plants Affected"},
            max_height=360,
        )

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#127981; Pivot View (Zone &#8594; Plant &#8594; Shortage)</div>",
        unsafe_allow_html=True,
    )

    detail_cols = [
        "Zone Name", "Plant Name", "Plant", "Billing Document", "Shipment Number",
        "Sold-to Party", "Service Agent", "Sales Organization", "Delivery", "Material",
        "Billed Quantity", "Shortage Quantity (in Ltrs)", "TT Number", "Created on", "Shortage Age (Days)",
    ]
    detail_cols = [c for c in detail_cols if c in detail_df.columns]

    all_zones = sorted(summary_df["Zone Name"].dropna().unique().tolist())
    tabs = st.tabs(all_zones) if all_zones else []
    for tab, zone in zip(tabs, all_zones):
        with tab:
            zone_rows = detail_df[detail_df["Zone Name"] == zone].copy()
            if zone_rows.empty:
                st.info("No records in this zone.")
                continue

            plant_summary = (
                zone_rows.groupby("Plant Name", dropna=False)
                .agg(total_shortage=("Shortage Quantity (in Ltrs)", "sum"))
                .reset_index()
                .rename(columns={"total_shortage": "Total Shortage Quantity (in Ltrs)"})
                .sort_values("Total Shortage Quantity (in Ltrs)", ascending=False)
            )
            _render_html_table(
                plant_summary,
                col_labels={"Plant Name": "Plant"},
                max_height=260,
            )

            for _, prow in plant_summary.iterrows():
                plant = prow["Plant Name"]
                sort_cols = [c for c in ["Shortage Age (Days)", "Shortage Quantity (in Ltrs)"] if c in detail_cols]
                sort_asc  = [False, False][:len(sort_cols)]
                p_df = zone_rows[zone_rows["Plant Name"] == plant][detail_cols]
                if sort_cols:
                    p_df = p_df.sort_values(sort_cols, ascending=sort_asc)
                p_df = p_df.reset_index(drop=True)
                with st.expander(f"{plant}  |  {len(p_df)} record(s)", expanded=False):
                    _render_html_table(p_df, max_height=360)

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#128270; Sortable Open Shortages (Sales) Records</div>",
        unsafe_allow_html=True,
    )
    sortable_df = detail_df[detail_cols].copy() if detail_cols else detail_df.copy()
    if not sortable_df.empty:
        _render_streamlit_dataframe(sortable_df, max_height=420, hide_index=True)

    unmatched = open_short_sales_result.get("unmatched", [])
    if unmatched:
        with st.expander(
            f"&#9888; {len(unmatched)} Plant Code(s) not found in PlantMaster",
            expanded=False,
        ):
            st.warning(
                "The following Plant codes could not be mapped to PlantMaster: "
                + ", ".join(sorted(map(str, unmatched)))
            )

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        xlsx_bytes = export_to_excel({
            "Zone Summary"    : zone_summary,
            "Plant Summary"   : summary_df,
            "Detailed Records": detail_df if not detail_df.empty else pd.DataFrame(),
        })
        st.download_button(
            label     = "&#11015;  Download Open Shortages (Sales) Report  (.xlsx)",
            data      = xlsx_bytes,
            file_name = f"OpenShortagesSales_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key       = "dl_open_short_sales",
        )


def render_open_shortages_sto_details(
    open_short_sto_result: dict,
    zone_filter         : list,
    plant_filter        : list,
) -> None:
    """Drill-down detail page for Open Shortages (STO)."""
    render_header(subtitle="&#128202; Open Shortages (STO) &#8212; Drill Down")

    back_col, _ = st.columns([1, 6])
    with back_col:
        if st.button("&#9664;  Back to Dashboard", key="btn_back_open_short_sto"):
            st.session_state["page"] = "dashboard"
            st.rerun()

    st.markdown("""
    <div class="detail-hdr">
        <h3>&#128202; Open Shortages (STO) &#8212; Detailed Exception View</h3>
        <p>Pivot-style Zone and Plant grouping for STO shortage transactions.</p>
    </div>
    """, unsafe_allow_html=True)

    summary_df   = open_short_sto_result.get("summary_df", pd.DataFrame())
    zone_summary = open_short_sto_result.get("zone_summary", pd.DataFrame())
    detail_df    = open_short_sto_result.get("detail_df", pd.DataFrame())
    total_qty    = float(open_short_sto_result.get("total_count", 0) or 0)

    if summary_df.empty:
        st.info("&#8505; No Open Shortages (STO) data available for the current filter selection.")
        return

    total_zones  = int(summary_df["Zone Name"].nunique())
    total_plants = int(summary_df["Plant Name"].nunique())
    avg_age = pd.NA
    if "Shortage Age (Days)" in detail_df.columns and not detail_df.empty:
        avg_age = pd.to_numeric(detail_df["Shortage Age (Days)"], errors="coerce").mean()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total STO Shortage Quantity (Ltrs)", f"{total_qty:,.2f}")
    m2.metric("Zones Affected", f"{total_zones}")
    m3.metric("Plants Affected", f"{total_plants}")
    m4.metric("Avg Shortage Age (Days)", f"{avg_age:.1f}" if pd.notna(avg_age) else "N/A")

    st.markdown("---")

    st.markdown(
        "<div class='sec-title'>&#128205; Zone-wise STO Shortage Quantity Summary</div>",
        unsafe_allow_html=True,
    )
    if not zone_summary.empty:
        _render_html_table(
            zone_summary,
            col_labels={"Plants": "Plants Affected"},
            max_height=360,
        )

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#127981; Pivot View (Zone &#8594; Plant &#8594; STO Shortage)</div>",
        unsafe_allow_html=True,
    )

    detail_cols = [
        "Zone Name", "Plant Name", "Supplying Plant", "Billing Document", "Shipment Number",
        "Plant", "Service Agent", "Sales Organization", "Delivery", "Vehicle", "Material",
        "Billed Quantity", "Sales Unit", "Shortage Quantity (in Ltrs)", "Created By", "Created On", "Shortage Age (Days)",
    ]
    detail_cols = [c for c in detail_cols if c in detail_df.columns]

    all_zones = sorted(summary_df["Zone Name"].dropna().unique().tolist())
    tabs = st.tabs(all_zones) if all_zones else []
    for tab, zone in zip(tabs, all_zones):
        with tab:
            zone_rows = detail_df[detail_df["Zone Name"] == zone].copy()
            if zone_rows.empty:
                st.info("No records in this zone.")
                continue

            plant_summary = (
                zone_rows.groupby("Plant Name", dropna=False)
                .agg(total_shortage=("Shortage Quantity (in Ltrs)", "sum"))
                .reset_index()
                .rename(columns={"total_shortage": "Total STO Shortage Quantity (in Ltrs)"})
                .sort_values("Total STO Shortage Quantity (in Ltrs)", ascending=False)
            )
            _render_html_table(
                plant_summary,
                col_labels={"Plant Name": "Plant"},
                max_height=260,
            )

            for _, prow in plant_summary.iterrows():
                plant = prow["Plant Name"]
                sort_cols = [c for c in ["Shortage Age (Days)", "Shortage Quantity (in Ltrs)"] if c in detail_cols]
                sort_asc  = [False, False][:len(sort_cols)]
                p_df = zone_rows[zone_rows["Plant Name"] == plant][detail_cols]
                if sort_cols:
                    p_df = p_df.sort_values(sort_cols, ascending=sort_asc)
                p_df = p_df.reset_index(drop=True)
                with st.expander(f"{plant}  |  {len(p_df)} record(s)", expanded=False):
                    _render_html_table(p_df, max_height=360)

    st.markdown(
        "<div class='sec-title' style='margin-top:20px;'>"
        "&#128270; Sortable Open Shortages (STO) Records</div>",
        unsafe_allow_html=True,
    )
    sortable_df = detail_df[detail_cols].copy() if detail_cols else detail_df.copy()
    if not sortable_df.empty:
        _render_streamlit_dataframe(sortable_df, max_height=420, hide_index=True)

    unmatched = open_short_sto_result.get("unmatched", [])
    if unmatched:
        with st.expander(
            f"&#9888; {len(unmatched)} Plant Code(s) not found in PlantMaster",
            expanded=False,
        ):
            st.warning(
                "The following Supplying Plant codes could not be mapped to PlantMaster: "
                + ", ".join(sorted(map(str, unmatched)))
            )

    st.markdown("---")
    dl_col, _ = st.columns([1, 4])
    with dl_col:
        xlsx_bytes = export_to_excel({
            "Zone Summary"    : zone_summary,
            "Plant Summary"   : summary_df,
            "Detailed Records": detail_df if not detail_df.empty else pd.DataFrame(),
        })
        st.download_button(
            label     = "&#11015;  Download Open Shortages (STO) Report  (.xlsx)",
            data      = xlsx_bytes,
            file_name = f"OpenShortagesSTO_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key       = "dl_open_short_sto",
        )


# ─────────────────────────────────────────────────────────────────────────────
# APPLICATION ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    """Bootstrap: session state → CSS → master data → sidebar → data → page."""

    if "page" not in st.session_state:
        st.session_state["page"] = "dashboard"
    if "location_visit_page" not in st.session_state:
        st.session_state["location_visit_page"] = "main"
    if "lv_sub_page" not in st.session_state:
        st.session_state["lv_sub_page"] = "summary"
    if "tank_turns_page" not in st.session_state:
        st.session_state["tank_turns_page"] = "main"
    if "open_mail_center" not in st.session_state:
        st.session_state["open_mail_center"] = False

    inject_css()

    # ── Force light theme via JS MutationObserver ──────────────────────
    # Streamlit's React engine re-applies dark background-color as an inline
    # style after every render, defeating CSS-only fixes. This iframe (same-
    # origin) uses setProperty(...,'important') + MutationObserver to
    # continuously re-enforce the light palette whenever Streamlit touches
    # the DOM.  setInterval(100ms) catches any React re-renders that slip past.
    import streamlit.components.v1 as _components
    _components.html("""
<script>
(function(){
  var BG  = '#F4F6FA';
  var BG2 = '#F0F2F6';
  var TX  = '#262730';

  function forceSidebar() {
    try {
      var doc = window.parent.document;
      var sb = doc.querySelector('[data-testid="stSidebar"]');
      if (sb) {
        sb.style.setProperty('display',     'block',   'important');
        sb.style.setProperty('visibility',  'visible', 'important');
        sb.style.setProperty('transform',   'none',    'important');
        sb.style.setProperty('min-width',   '256px',   'important');
        sb.style.setProperty('width',       '256px',   'important');
        sb.setAttribute('aria-expanded', 'true');
      }
    } catch(e) {}
  }

  function applyLight() {
    forceSidebar();
    try {
      var doc = window.parent.document;
      // Force body & root
      doc.documentElement.style.setProperty('background-color', BG,  'important');
      doc.documentElement.style.setProperty('color',            TX,  'important');
      doc.body.style.setProperty('background-color', BG, 'important');
      doc.body.style.setProperty('color',            TX, 'important');

      // Force every Streamlit container selector
      var sels = [
        '#root',
        '.stApp',
        '[data-testid="stApp"]',
        '[data-testid="stAppViewContainer"]',
        '[data-testid="stAppViewBlockContainer"]',
        '[data-testid="stMain"]',
        '[data-testid="stMainBlockContainer"]',
        '[data-testid="block-container"]',
        'section.main',
        '.main'
      ];
      sels.forEach(function(s){
        var el = doc.querySelector(s);
        if (el) {
          el.style.setProperty('background-color', BG, 'important');
          el.style.setProperty('color',            TX, 'important');
        }
      });

      // Clear stored dark-theme localStorage key (all known variants)
      var ls = window.parent.localStorage;
      ['stActiveTheme','theme','st_theme','streamlit_theme'].forEach(function(k){
        var v = ls.getItem(k);
        if (v && v.toLowerCase().includes('dark')) { ls.removeItem(k); }
      });
    } catch(e) {}
  }

  // Fire immediately and after short delay (wait for React to mount)
  applyLight();
  setTimeout(applyLight, 300);
  setTimeout(applyLight, 800);
  setTimeout(applyLight, 2000);

  // MutationObserver: re-apply whenever Streamlit mutates the DOM/styles
  var obs = new MutationObserver(applyLight);
  try {
    obs.observe(window.parent.document.body, {
      attributes: true, attributeFilter: ['style','class'],
      childList: true,  subtree: true
    });
  } catch(e) {}

  // Periodic safety net every 500 ms
  setInterval(applyLight, 500);
})();
</script>
""", height=0, scrolling=False)

    # Load master data
    try:
        df_plant = load_plant_master()
    except FileNotFoundError:
        st.error(
            f"PlantMaster.xlsx not found at: `{PLANT_MASTER_PATH}`\n\n"
            "Ensure the MAster/ folder sits next to app.py."
        )
        st.stop()
    except Exception as exc:
        st.error(f"Failed to load PlantMaster: {exc}")
        st.stop()

    zone_count = int(df_plant["Zone Name"].nunique()) if "Zone Name" in df_plant.columns else 0
    if zone_count != 16:
        st.sidebar.warning(
            f"PlantMaster currently has {zone_count} active zone(s). Expected 16 as per latest structure."
        )

    try:
        load_zone_master()
    except Exception as exc:
        st.sidebar.warning(f"Zone master not loaded: {exc}")

    # Sidebar
    selected_zones, selected_plants, uploaded_dc, sidebar_system_info_slot, as_of_date = render_sidebar(df_plant)

    # Resolve data source
    pending_dc_xls = os.path.join(REPORTS_DIR, "PENDING_DC_SOD.xls")
    if os.path.exists(pending_dc_xls):
        # Convert .xls to .xlsx permanently with robust fallback
        try:
            try:
                df_xls = pd.read_excel(pending_dc_xls, engine="xlrd")
            except Exception as exc_xlrd:
                st.sidebar.warning(f"xlrd failed: {exc_xlrd}. Trying openpyxl...")
                try:
                    df_xls = pd.read_excel(pending_dc_xls, engine="openpyxl")
                except Exception as exc_openpyxl:
                    st.sidebar.error(f"❌ Both xlrd and openpyxl failed: {exc_openpyxl}")
                    df_xls = None
            if df_xls is not None:
                pending_dc_xlsx = os.path.join(REPORTS_DIR, "PENDING_DC_SOD.xlsx")
                df_xls.to_excel(pending_dc_xlsx, index=False)
                st.sidebar.success("Pending DC .xls converted to .xlsx.")
            else:
                st.sidebar.error("❌ Could not convert Pending DC .xls to .xlsx.")
        except Exception as exc:
            st.sidebar.error(f"❌ Error converting Pending DC .xls to .xlsx: {exc}")
    if uploaded_dc is not None:
        dc_source = uploaded_dc
    elif os.path.exists(PENDING_DC_PATH):
        dc_source = PENDING_DC_PATH
    else:
        dc_source = None

    # Load & process
    if dc_source is not None:
        with st.spinner("Loading Pending DC data …"):
            df_dc = load_pending_dc(dc_source)
        if df_dc.empty:
            st.sidebar.warning("⚠️ Pending DC file is empty.")
        pending_dc_result = process_pending_dc(
            df_dc,
            df_plant,
            zone_filter  = selected_zones  or None,
            plant_filter = selected_plants or None,
            as_of_date   = as_of_date,
        )
    else:
        pending_dc_result = {
            "total_count" : 0,
            "summary_df"  : pd.DataFrame(),
            "zone_summary": pd.DataFrame(),
            "detail_df"   : pd.DataFrame(),
            "unmatched"   : [],
        }
        st.sidebar.warning("No Pending DC file found. Upload via the sidebar.")

    # Load & process Open Deliveries (default file path based)
    if os.path.exists(OPEN_DELIVERY_PATH):
        with st.spinner("Loading Open Delivery data …"):
            df_open_delivery = load_open_delivery(OPEN_DELIVERY_PATH)
        open_delivery_result = process_open_deliveries(
            df_open_delivery,
            df_plant,
            zone_filter  = selected_zones  or None,
            plant_filter = selected_plants or None,
            as_of_date   = as_of_date,
        )
    else:
        open_delivery_result = {
            "total_count" : 0,
            "summary_df"  : pd.DataFrame(),
            "zone_summary": pd.DataFrame(),
            "detail_df"   : pd.DataFrame(),
            "unmatched"   : [],
        }

    # Load & process Open In-Transit (default file path based)
    if os.path.exists(OPEN_INTRANSIT_PATH):
        with st.spinner("Loading Open In-Transit data …"):
            df_open_intransit = load_open_intransit(OPEN_INTRANSIT_PATH)
        open_intransit_result = process_open_intransit(
            df_open_intransit,
            df_plant,
            zone_filter  = selected_zones  or None,
            plant_filter = selected_plants or None,
            as_of_date   = as_of_date,
        )
    else:
        open_intransit_result = {
            "total_count" : 0,
            "summary_df"  : pd.DataFrame(),
            "zone_summary": pd.DataFrame(),
            "detail_df"   : pd.DataFrame(),
            "unmatched"   : [],
        }

    # Load & process Open Sales Orders (default file path based)
    if os.path.exists(OPEN_SO_PATH):
        with st.spinner("Loading Open Sales Orders data …"):
            df_open_so = load_open_sales_orders(OPEN_SO_PATH)
        open_sales_orders_result = process_open_sales_orders(
            df_open_so,
            df_plant,
            zone_filter  = selected_zones  or None,
            plant_filter = selected_plants or None,
            as_of_date   = as_of_date,
        )
    else:
        open_sales_orders_result = {
            "total_count" : 0,
            "summary_df"  : pd.DataFrame(),
            "zone_summary": pd.DataFrame(),
            "detail_df"   : pd.DataFrame(),
            "unmatched"   : [],
        }

    # Load & process Pending Invoices (default file path based)
    if os.path.exists(PEND_INV_PATH):
        with st.spinner("Loading Pending Invoices data …"):
            df_pending_inv = load_pending_invoices(PEND_INV_PATH)
        pending_invoices_result = process_pending_invoices(
            df_pending_inv,
            df_plant,
            zone_filter  = selected_zones  or None,
            plant_filter = selected_plants or None,
            as_of_date   = as_of_date,
        )
    else:
        pending_invoices_result = {
            "total_count" : 0,
            "summary_df"  : pd.DataFrame(),
            "zone_summary": pd.DataFrame(),
            "detail_df"   : pd.DataFrame(),
            "unmatched"   : [],
        }

    # Load & process Tank Reco (default file path based)
    if os.path.exists(TANK_RECO_PATH):
        with st.spinner("Loading Tank Reco data …"):
            df_tank_reco = load_tank_reco(TANK_RECO_PATH)
        tank_reco_result = process_tank_reco(
            df_tank_reco,
            df_plant,
            zone_filter  = selected_zones  or None,
            plant_filter = selected_plants or None,
        )
    else:
        tank_reco_result = {
            "total_count" : 0,
            "summary_df"  : pd.DataFrame(),
            "zone_summary": pd.DataFrame(),
            "detail_df"   : pd.DataFrame(),
            "unmatched"   : [],
        }

    # Load & process Open Shortages (Sales) (default file path based)
    if os.path.exists(SHORT_SALES_PATH):
        with st.spinner("Loading Open Shortages (Sales) data …"):
            df_short_sales = load_open_shortages_sales(SHORT_SALES_PATH)
        open_short_sales_result = process_open_shortages_sales(
            df_short_sales,
            df_plant,
            zone_filter  = selected_zones  or None,
            plant_filter = selected_plants or None,
            as_of_date   = as_of_date,
        )
    else:
        open_short_sales_result = {
            "total_count" : 0.0,
            "summary_df"  : pd.DataFrame(),
            "zone_summary": pd.DataFrame(),
            "detail_df"   : pd.DataFrame(),
            "unmatched"   : [],
        }

    # Load & process Open Shortages (STO) (default file path based)
    if os.path.exists(SHORT_STO_PATH):
        with st.spinner("Loading Open Shortages (STO) data …"):
            df_short_sto = load_open_shortages_sto(SHORT_STO_PATH)
        open_short_sto_result = process_open_shortages_sto(
            df_short_sto,
            df_plant,
            zone_filter  = selected_zones  or None,
            plant_filter = selected_plants or None,
            as_of_date   = as_of_date,
        )
    else:
        open_short_sto_result = {
            "total_count" : 0.0,
            "summary_df"  : pd.DataFrame(),
            "zone_summary": pd.DataFrame(),
            "detail_df"   : pd.DataFrame(),
            "unmatched"   : [],
        }

    # Build all-exception summary tables for dashboard and zone drill-down
    all_exception_plant_df = _build_all_exception_plant_summary(
        pending_dc_result,
        open_delivery_result,
        open_intransit_result,
        open_sales_orders_result,
        pending_invoices_result,
        tank_reco_result,
        open_short_sales_result,
        open_short_sto_result,
    )
    zone_exception_summary_df = _build_zone_exception_summary(all_exception_plant_df)
    shortage_location_summary_df = _build_combined_shortage_location_summary(
        open_short_sales_result,
        open_short_sto_result,
    )
    shortage_zone_summary_df = _build_combined_shortage_zone_summary(shortage_location_summary_df)
    combined_shortage_detail_df = _build_combined_shortage_detail_df(
        open_short_sales_result,
        open_short_sto_result,
    )
    short_sales_vehicle_summary_df = _build_vehicle_shortage_summary(
        open_short_sales_result.get("detail_df", pd.DataFrame()),
        "TT Number",
        "TT Number",
    )
    short_sto_vehicle_summary_df = _build_vehicle_shortage_summary(
        open_short_sto_result.get("detail_df", pd.DataFrame()),
        "Vehicle",
        "Vehicle",
    )

    sidebar_kpi_df = _build_exception_kpi_chart_df(
        pending_dc_result,
        open_delivery_result,
        open_intransit_result,
        open_sales_orders_result,
        pending_invoices_result,
        tank_reco_result,
        open_short_sales_result,
        open_short_sto_result,
    )
    _render_sidebar_system_info(
        sidebar_system_info_slot,
        df_plant,
        all_exception_plant_df=all_exception_plant_df,
        exception_kpi_df=sidebar_kpi_df,
    )

    # Page router
    page = st.session_state.get("page", "dashboard")

    if page == "dashboard":
        render_dashboard(
            df_plant,
            pending_dc_result,
            open_delivery_result,
            open_intransit_result,
            open_sales_orders_result,
            pending_invoices_result,
            tank_reco_result,
            open_short_sales_result,
            open_short_sto_result,
            all_exception_plant_df,
            zone_exception_summary_df,
            selected_zones,
            selected_plants,
            as_of_date=as_of_date,
        )
    elif page == "pending_dc_details":
        render_pending_dc_details(pending_dc_result, selected_zones, selected_plants)
    elif page == "open_delivery_details":
        render_open_delivery_details(open_delivery_result, selected_zones, selected_plants)
    elif page == "open_intransit_details":
        render_open_intransit_details(open_intransit_result, selected_zones, selected_plants)
    elif page == "open_sales_orders_details":
        render_open_sales_orders_details(open_sales_orders_result, selected_zones, selected_plants)
    elif page == "pending_invoices_details":
        render_pending_invoices_details(pending_invoices_result, selected_zones, selected_plants)
    elif page == "tank_reco_details":
        render_tank_reco_details(tank_reco_result, selected_zones, selected_plants)
    elif page == "open_shortages_sales_details":
        render_open_shortages_sales_details(open_short_sales_result, selected_zones, selected_plants)
    elif page == "open_shortages_sto_details":
        render_open_shortages_sto_details(open_short_sto_result, selected_zones, selected_plants)
    elif page == "zone_exception_drilldown":
        render_zone_exception_drilldown(
            zone_exception_summary_df,
            all_exception_plant_df,
            selected_zones,
            selected_plants,
        )
    elif page == "top_exception_zones":
        render_top_exception_zones_page(
            zone_exception_summary_df,
            all_exception_plant_df,
            selected_zones,
            selected_plants,
        )
    elif page == "top_exception_locations":
        render_top_exception_locations_page(
            all_exception_plant_df,
            selected_zones,
            selected_plants,
        )
    elif page == "top_shortage_zones":
        render_top_shortage_zones_page(
            shortage_zone_summary_df,
            shortage_location_summary_df,
            combined_shortage_detail_df,
            selected_zones,
            selected_plants,
        )
    elif page == "top_shortage_locations":
        render_top_shortage_locations_page(
            shortage_location_summary_df,
            combined_shortage_detail_df,
            selected_zones,
            selected_plants,
        )
    elif page == "top_short_sales_vehicles":
        render_top_short_sales_vehicles_page(
            short_sales_vehicle_summary_df,
            open_short_sales_result,
            selected_zones,
            selected_plants,
        )
    elif page == "top_short_sto_vehicles":
        render_top_short_sto_vehicles_page(
            short_sto_vehicle_summary_df,
            open_short_sto_result,
            selected_zones,
            selected_plants,
        )
    else:
        st.session_state["page"] = "dashboard"
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    main()
