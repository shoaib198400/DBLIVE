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

# --- LOCATION VISIT DATA LOADING & MAPPING ---
try:
    df_loc = pd.read_excel("Reports/LOCATION_VISIT.xls")
except Exception:
    df_loc = pd.DataFrame()

zone_mapping = {
    "NCZ": "North Central Zone (NCZ)",
    "NWF": "North West Frontier Zone (NWFZ)",
    "NZ": "North Zone (NZ)",
    "SZ": "South Zone (SZ)",
    "SWZ": "South West Zone (SWZ)",
    "WZ": "West Zone (WZ)",
    "ECZ": "East Central Zone (ECZ)",
    "SCZ": "South Central Zone (SCZ)",
    "EZ": "East Zone (EZ)",
    "NWZ": "North West Zone (NWZ)",
    "CNT": "Central Zone (CZ)",
    "NFZ": "North Frontier Zone (NFZ)"
}
reverse_zone_mapping = {v: k for k, v in zone_mapping.items()}

def render_location_visit_details(df):
    st.markdown("<div class='sec-title'>&#128205; Location Visit - %Compliance &#8212; Drill Down</div>", unsafe_allow_html=True)

    back_col, _ = st.columns([1, 7])
    with back_col:
        if st.button("&#11013; Back to Dashboard", key="btn_back_location_visit"):
            st.session_state["location_visit_page"] = "main"
            st.session_state["selected_tile"] = None
            st.session_state["dummy_tank_clicked"] = False
            st.session_state["pl_unblock_clicked"] = False
            st.session_state["tank_turns_page"] = "main"
            st.rerun()

    if df is None or df.empty:
        st.warning("No data available after applying filters.")
        return

    def _count_present(series: pd.Series) -> int:
        return int(
            series.astype(str)
            .str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaT": pd.NA})
            .dropna()
            .shape[0]
        )

    def _prepare_presence_flags(frame: pd.DataFrame) -> pd.DataFrame:
        flagged = frame.copy()
        flagged["_total_flag"] = (
            flagged["TotalRecomms"].astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaT": pd.NA}).notna().astype(int)
        )
        flagged["_closed_flag"] = (
            flagged["ClosedRecomms"].astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaT": pd.NA}).notna().astype(int)
        )
        flagged["_open_flag"] = (
            flagged["OpenRecomms"].astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaT": pd.NA}).notna().astype(int)
        )
        return flagged

    def _group_compliance(frame: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
        grouped = (
            _prepare_presence_flags(frame)
            .groupby(group_cols, dropna=False, as_index=False)
            .agg(
                Total_Recomms=("_total_flag", "sum"),
                Closed_Recomms=("_closed_flag", "sum"),
                Open_Recomms=("_open_flag", "sum"),
            )
        )
        grouped["Compliance Value"] = (
            grouped["Closed_Recomms"] / grouped["Total_Recomms"].replace(0, pd.NA)
        ).fillna(0) * 100
        return grouped

    required_cols = [
        "Zone", "Planning Plant", "Plant Desc.", "Audit Number",
        "Audit Start Date", "Audit End Date", "TotalRecomms", "ClosedRecomms", "OpenRecomms",
    ]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        st.warning("Missing required column(s): " + ", ".join(missing_cols))
        return

    df_source = df.copy()
    df_view = df.copy()

    st.markdown("<div class='sec-title'>&#128269; Drill-Down Filters</div>", unsafe_allow_html=True)
    f1, f2 = st.columns(2)

    zone_opts = ["All"] + sorted(df_view["Zone"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
    plant_desc_opts = ["All"] + sorted(df_view["Plant Desc."].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())

    sel_zone = f1.selectbox("Zone", zone_opts, key="lv_filter_zone")
    sel_plant_desc = f2.selectbox("Plant Desc.", plant_desc_opts, key="lv_filter_plant_desc")

    if sel_zone != "All":
        df_view = df_view[df_view["Zone"].astype(str) == sel_zone]
    if sel_plant_desc != "All":
        df_view = df_view[df_view["Plant Desc."].astype(str) == sel_plant_desc]

    if df_view.empty:
        st.info("No data available after applying drill-down filters.")
        return

    total_recomms = _count_present(df_view["TotalRecomms"])
    total_closed = _count_present(df_view["ClosedRecomms"])
    compliance = (total_closed / total_recomms) if total_recomms else 0.0

    m1, m2 = st.columns(2)
    m1.metric("Location Visit Count", f"{df_view['Planning Plant'].nunique():,}")
    m2.metric("Compliance (%)", f"{compliance * 100:.1f}%")

    st.markdown("<div class='sec-title'>&#128203; Location Visit Details (Pivot)</div>", unsafe_allow_html=True)

    pivot_df = (
        _prepare_presence_flags(df_view).groupby(
            ["Planning Plant", "Audit Number", "Audit Start Date", "Audit End Date"],
            as_index=False,
            dropna=False,
        )
        .agg({
            "_total_flag": "sum",
            "_closed_flag": "sum",
            "_open_flag": "sum",
        })
        .reset_index(drop=True)
    )

    pivot_df["Compliance (%)"] = (
        pivot_df["_closed_flag"] / pivot_df["_total_flag"].replace(0, pd.NA)
    ).fillna(0) * 100

    pivot_df = pivot_df.rename(columns={
        "_total_flag": "Count of Total Recomms",
        "_closed_flag": "Count of Closed Recomms",
        "_open_flag": "Count of Open Recomms",
    })

    pivot_display = pivot_df.copy()
    pivot_display["Compliance (%)"] = pivot_display["Compliance (%)"].map(lambda x: f"{x:.2f}%")
    _render_html_table(pivot_display, max_height=500)

    _download_excel_button(
        label="&#11015;  Download Location Visit Raw Data  (.xlsx)",
        file_prefix="location_visit",
        sheets={"Location_Visit_Raw": df_view.reset_index(drop=True)},
        key="dl_location_visit",
    )

    location_compliance = _group_compliance(df_view, ["Planning Plant", "Plant Desc."])
    location_compliance["Location Label"] = (
        location_compliance["Planning Plant"].astype(str).str.strip()
        + " - "
        + location_compliance["Plant Desc."].astype(str).str.strip()
    )
    zone_compliance = _group_compliance(df_view, ["Zone"])

    top5_locations = location_compliance.sort_values(
        ["Compliance Value", "Closed_Recomms", "Total_Recomms"],
        ascending=[False, False, False],
    ).head(5).reset_index(drop=True)
    bottom5_locations = location_compliance.sort_values(
        ["Compliance Value", "Total_Recomms", "Planning Plant"],
        ascending=[True, False, True],
    ).head(5).reset_index(drop=True)
    top5_zones = zone_compliance.sort_values(
        ["Compliance Value", "Closed_Recomms", "Total_Recomms"],
        ascending=[False, False, False],
    ).head(5).reset_index(drop=True)
    bottom5_zones = zone_compliance.sort_values(
        ["Compliance Value", "Total_Recomms", "Zone"],
        ascending=[True, False, True],
    ).head(5).reset_index(drop=True)

    st.markdown("<div class='sec-title'>&#128202; Top / Bottom 5 Compliance Analysis</div>", unsafe_allow_html=True)

    def _format_compliance_table(frame: pd.DataFrame, label_cols: list[str]) -> pd.DataFrame:
        display = frame[label_cols + ["Compliance Value", "Total_Recomms", "Closed_Recomms"]].copy()
        display = display.rename(columns={
            "Compliance Value": "Compliance (%)",
            "Total_Recomms": "Total Recomms",
            "Closed_Recomms": "Closed Recomms",
        })
        display["Compliance (%)"] = display["Compliance (%)"].map(lambda x: f"{x:.2f}%")
        display["Total Recomms"] = display["Total Recomms"].map(lambda x: f"{x:,}")
        display["Closed Recomms"] = display["Closed Recomms"].map(lambda x: f"{x:,}")
        return display

    def _render_compliance_charts(frame: pd.DataFrame, label_col: str, title_prefix: str, colors: list[str]) -> None:
        if frame.empty:
            st.info("No analysis available.")
            return
        chart_col1, chart_col2 = st.columns(2)
        with chart_col1:
            pie_fig = px.pie(
                frame,
                names=label_col,
                values="Compliance Value",
                hole=0.62,
                color_discrete_sequence=colors,
                title=f"{title_prefix} Compliance Share",
            )
            pie_fig.update_traces(textposition="inside", textinfo="percent+label")
            pie_fig.update_layout(margin=dict(l=20, r=20, t=60, b=20), height=420)
            st.plotly_chart(pie_fig, use_container_width=True)
        with chart_col2:
            bar_fig = px.bar(
                frame,
                x="Compliance Value",
                y=label_col,
                orientation="h",
                color="Compliance Value",
                color_continuous_scale=colors,
                title=f"{title_prefix} Compliance (%)",
            )
            bar_fig.update_layout(margin=dict(l=20, r=20, t=60, b=20), height=420, coloraxis_showscale=False)
            bar_fig.update_xaxes(title_text="Compliance (%)", ticksuffix="%")
            bar_fig.update_yaxes(title_text="")
            st.plotly_chart(bar_fig, use_container_width=True)

    st.markdown("<div class='sec-title'>&#11014; Top 5 Locations by Compliance</div>", unsafe_allow_html=True)
    _render_html_table(_format_compliance_table(top5_locations, ["Planning Plant", "Plant Desc."]), max_height=260)
    _render_compliance_charts(top5_locations, "Location Label", "Top 5 Locations", px.colors.sequential.Blues)

    st.markdown("<div class='sec-title'>&#11015; Bottom 5 Locations by Compliance</div>", unsafe_allow_html=True)
    _render_html_table(_format_compliance_table(bottom5_locations, ["Planning Plant", "Plant Desc."]), max_height=260)
    _render_compliance_charts(bottom5_locations, "Location Label", "Bottom 5 Locations", px.colors.sequential.Reds)

    st.markdown("<div class='sec-title'>&#11014; Top 5 Zones by Compliance</div>", unsafe_allow_html=True)
    _render_html_table(_format_compliance_table(top5_zones, ["Zone"]), max_height=260)
    _render_compliance_charts(top5_zones, "Zone", "Top 5 Zones", px.colors.sequential.Greens)

    st.markdown("<div class='sec-title'>&#11015; Bottom 5 Zones by Compliance</div>", unsafe_allow_html=True)
    _render_html_table(_format_compliance_table(bottom5_zones, ["Zone"]), max_height=260)
    _render_compliance_charts(bottom5_zones, "Zone", "Bottom 5 Zones", px.colors.sequential.OrRd)

    plant_master_scope = load_plant_master()[["Plant Code", "Plant Name", "Zone Name"]].copy()
    plant_master_scope["Plant Code"] = (
        plant_master_scope["Plant Code"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    )

    missing_scope = df_source.copy()
    if sel_zone != "All":
        missing_scope = missing_scope[missing_scope["Zone"].astype(str) == sel_zone]
        plant_master_scope = plant_master_scope[plant_master_scope["Zone Name"].astype(str) == sel_zone]

    source_codes = set(
        missing_scope["Planning Plant"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True).unique().tolist()
    )
    missing_locations = (
        plant_master_scope[~plant_master_scope["Plant Code"].isin(source_codes)]
        .rename(columns={
            "Zone Name": "Zone",
            "Plant Code": "Planning Plant",
            "Plant Name": "Location",
        })
        [["Zone", "Planning Plant", "Location"]]
        .sort_values(["Zone", "Planning Plant"])
        .reset_index(drop=True)
    )

    st.markdown("<div class='sec-title'>&#9888; Missing Zones/Locations</div>", unsafe_allow_html=True)
    if missing_locations.empty:
        st.success("All PlantMaster zones/locations in the current scope are present in the Location Visit source file.")
    else:
        _render_html_table(missing_locations, max_height=320)

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
LOCATION_VISIT_PATH = os.path.join(REPORTS_DIR, "LOCATION_VISIT.xls")
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
    [data-testid="stHeader"] > div {{
        display: none !important;
    }}
    [data-testid="stDecoration"] {{
        display: none !important;
    }}
    [data-testid="stToolbar"] {{
        display: none !important;
    }}
    [data-testid="stToolbar"] button,
    [data-testid="stToolbar"] [data-testid="baseButton-header"] {{
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
        padding-bottom: 1.5rem;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
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
        height: 192px;
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
        padding: 0 10px 4px 10px;
        filter: drop-shadow(0 0 1px rgba(0, 48, 135, 0.9))
                drop-shadow(0 0 2px rgba(0, 48, 135, 0.65));
    }}

    /* ── Info Strip below banner ──────────────────────── */
    .dash-header {{
        background: linear-gradient(135deg, {C['primary']} 0%, {C['secondary']} 100%);
        color: white;
        padding: 14px 28px;
        margin: 0 -1rem 20px -1rem;
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        box-shadow: 0 4px 16px rgba(0,48,135,0.35);
    }}
    .dash-header-main {{
        width: 100%;
        text-align: center;
    }}
    .dash-header-title {{
        font-size: 40px !important;
        font-weight: 800;
        letter-spacing: 0.02em;
        margin: 0;
        line-height: 1.15;
    }}
    .dash-header-sub {{
        font-size: 18px;
        opacity: 0.88;
        margin: 6px 0 0 0;
    }}
    .dash-header-meta {{
        position: absolute;
        right: 28px;
        top: 50%;
        transform: translateY(-50%);
        text-align: right;
        font-size: 18px;
        opacity: 0.90;
        line-height: 1.65;
    }}
    @media (max-width: 1100px) {{
        .dash-header {{
            align-items: flex-start;
            justify-content: flex-start;
            gap: 12px;
            flex-wrap: wrap;
        }}
        .dash-header-main {{
            text-align: center;
        }}
        .dash-header-title {{
            font-size: 32px !important;
        }}
        .dash-header-meta {{
            position: static;
            transform: none;
            width: 100%;
            text-align: center;
        }}
    }}
    @media (max-width: 700px) {{
        .dash-header {{
            padding: 12px 18px;
        }}
        .dash-header-title {{
            font-size: 26px !important;
        }}
        .dash-header-sub {{
            font-size: 16px;
        }}
        .dash-header-meta {{
            font-size: 15px;
        }}
    }}

    /* ── KPI Cards ─────────────────────────────────────── */
    .kpi-wrap {{
        background: {C['white']};
        border-radius: 12px;
        padding: 18px 20px 14px 20px;
        border-left: 5px solid {C['primary']};
        box-shadow: 0 2px 12px {C['shadow']};
        position: relative;
        overflow: hidden;
        transition: transform 0.18s ease, box-shadow 0.18s ease;
        min-height: 130px;
    }}
    .kpi-wrap:hover {{
        transform: translateY(-3px);
        box-shadow: 0 8px 22px rgba(0,48,135,0.22);
    }}
    .kpi-wrap::after {{
        content: '';
        position: absolute;
        top: 0; right: 0;
        width: 60px; height: 60px;
        background: {C['light_blue']};
        border-radius: 0 12px 0 60px;
    }}
    .kpi-icon {{
        position: absolute;
        top: 14px; right: 16px;
        font-size: 1.7rem;
        opacity: 0.55;
        z-index: 1;
    }}
    .kpi-label {{
        font-size: 20px;
        font-weight: 800;
        color: #46515F;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 8px;
        line-height: 1.25;
        text-shadow: 0 0 0 rgba(0,0,0,0.01);
    }}
    .kpi-value {{
        font-size: 2.6rem;
        font-weight: 800;
        color: {C['primary']};
        line-height: 1;
        margin-bottom: 6px;
    }}
    .kpi-detail {{
        font-size: 19px;
        font-weight: 700;
        color: #24405A;
        line-height: 1.35;
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

    /* ── Section Titles (≥ 20 px) ────────────────────────── */
    .sec-title {{
        font-size: 22px;
        font-weight: 700;
        color: {C['primary']};
        padding: 8px 0 6px 0;
        border-bottom: 2px solid {C['light_blue']};
        margin: 10px 0 16px 0;
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
        color: white;
        font-size: 22px;
        font-weight: 800;
        text-align: center;
        padding: 16px 18px;
        border-bottom: 2px solid #D5E2F3;
    }}
    .pro-table tbody td {{
        font-size: 21px;
        font-weight: 600;
        color: #1B3552;
        text-align: center;
        padding: 16px 18px;
        border-bottom: 1px solid #E2EAF4;
        word-wrap: break-word;
    }}
    .pro-table tbody tr:nth-child(odd) {{
        background: #FFFFFF;
    }}
    .pro-table tbody tr:nth-child(even) {{
        background: #F7FAFE;
    }}
    .pro-table tbody tr:hover {{
        background: #EAF2FF;
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
        padding: 0 0 6px 0 !important;
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
        font-size: 15px !important;
    }}
    [data-testid="stSidebar"] .stMultiSelect label {{
        color: #FFFFFF !important;
        font-size: 14px !important;
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
        font-size: 13px !important;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: #7AABF0 !important;
        margin: 14px 0 6px 0;
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
        border-radius: 7px;
        padding: 8px 22px;
        font-weight: 600;
        font-size: 15px;
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
        border-radius: 7px;
        font-weight: 600;
        font-size: 15px;
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

    /* ── Larger fonts for Streamlit dataframe tables used in drill-down pages ── */
    [data-testid="stDataFrame"] th {{
        font-size: 18px !important;
        font-weight: 700 !important;
    }}
    [data-testid="stDataFrame"] td {{
        font-size: 17px !important;
        font-weight: 600 !important;
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
    usecols = ["Plant Code", "Plant Name", "Zone Name", "Active"]
    df = pd.read_excel(PLANT_MASTER_PATH, dtype={"Plant Code": str}, engine="openpyxl", usecols=usecols)
    df.columns = df.columns.str.strip()
    df["Plant Code"] = df["Plant Code"].astype(str).str.strip()
    df["Plant Name"] = df["Plant Name"].astype(str).str.strip()
    df["Zone Name"]  = df["Zone Name"].astype(str).str.strip()
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


@st.cache_data(show_spinner=False)
def _load_excel_from_path(path: str) -> pd.DataFrame:
    """Internal helper: load any Excel file from a disk path (cached)."""
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
            return _load_excel_from_path(source)
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
            return _load_excel_from_path(source)
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
            return _load_excel_from_path(source)
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
            return _load_excel_from_path(source)
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
            return _load_excel_from_path(source)
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
            return _load_excel_from_path(source)
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
            return _load_excel_from_path(source)
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
            return _load_excel_from_path(source)
        name   = getattr(source, "name", "file.xlsx")
        ext    = os.path.splitext(name)[1].lower()
        df = _read_excel_flexible(source, ext_hint=ext)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading OPEN SHORTAGES - Ltrs (STO) file: {exc}")
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_dummy_tank_stock(path: str) -> pd.DataFrame:
    """Load Dummy Tank Stock report from disk and normalize critical columns."""
    try:
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
def load_tank_turns(path: str) -> pd.DataFrame:
    """Load Tank Turn report from disk and normalize critical columns."""
    try:
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
def load_location_visit(path: str) -> pd.DataFrame:
    """Load Location Visit report from disk and normalize critical columns."""
    try:
        ext = os.path.splitext(path)[1].lower()
        df = _read_excel_flexible(path, ext_hint=ext)
        df.columns = df.columns.astype(str).str.strip()

        canonical_cols = {
            "zone": "Zone",
            "planning plant": "Planning Plant",
            "plant desc.": "Plant Desc.",
            "audit number": "Audit Number",
            "audit start date": "Audit Start Date",
            "audit end date": "Audit End Date",
            "totalrecomms": "TotalRecomms",
            "closedrecomms": "ClosedRecomms",
            "openrecomms": "OpenRecomms",
        }
        rename_map = {}
        for col in df.columns:
            key = " ".join(str(col).strip().lower().split())
            if key in canonical_cols:
                rename_map[col] = canonical_cols[key]
        if rename_map:
            df = df.rename(columns=rename_map)

        for col in ["Planning Plant", "Plant Desc.", "Audit Number", "Zone"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()
                if col == "Planning Plant":
                    df[col] = df[col].str.replace(r"\.0$", "", regex=True)

        for col in ["TotalRecomms", "ClosedRecomms", "OpenRecomms"]:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
                df[col] = df[col].replace(r"^\s*$", pd.NA, regex=True)

        for col in ["Audit Start Date", "Audit End Date"]:
            if col in df.columns:
                dt = pd.to_datetime(df[col], errors="coerce")
                df[col] = dt.dt.strftime("%d/%m/%Y").fillna("")

        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_pipeline_stock(path: str) -> pd.DataFrame:
    """Load Pipeline Stock report and normalize key columns."""
    try:
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
    now  = datetime.now().strftime("%d %b %Y  |  %I:%M %p")
    subtitle_html = f'<p class="dash-header-sub">{subtitle}</p>' if subtitle else ""

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
        </div>"""
    elif logo_uri:
        banner_html = f"""
        <div class="hpcl-banner-wrap">
            <img class="hpcl-banner-fg" src="{logo_uri}" alt="HPCL SOD Exception Dashboard" />
        </div>"""
    else:
        banner_html = f"""
        <div class="hpcl-banner-wrap" style="
                background:linear-gradient(135deg,#003087 0%,#0057A8 100%);
                height:192px;display:flex;align-items:center;
                padding:0 40px;justify-content:space-between;">
            <div style="font-size:56px;font-weight:900;color:#FFFFFF;
                        letter-spacing:.08em;">&#9981; HPCL</div>
            <div style="font-size:22px;color:#AACCFF;font-style:italic;">
                We bring Energy to Life</div>
        </div>"""

    # ── Part 2: Dark-blue info strip ────────────────────────────────────────
    header_html = (
        '<div class="dashboard-header-shell">'
        f'{banner_html}'
        '<div class="dash-header">'
        '<div class="dash-header-main">'
        '<p class="dash-header-title">SOD Exception Dashboard</p>'
        f'{subtitle_html}'
        '</div>'
        '<div class="dash-header-meta">'
        f'<div style="margin-top:4px;font-size:18px;">{now}</div>'
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
    """Render navigation sidebar. Returns (zones, plants, uploaded_file, system_info_slot)."""
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

        st.markdown('<p class="sb-nav-lbl">&#128194; Data Upload</p>',
                    unsafe_allow_html=True)
        unique_uploader_key = f"uploader_pending_dc_{uuid.uuid4()}"
        uploaded_dc = st.file_uploader(
            "Pending DC File  (.xls / .xlsx)",
            type = ["xls", "xlsx"],
            key  = unique_uploader_key,
        )

    return None, None, uploaded_dc, system_info_slot


# ─────────────────────────────────────────────────────────────────────────────
# SHARED HTML TABLE RENDERER
# ─────────────────────────────────────────────────────────────────────────────

def _render_html_table(df: pd.DataFrame, col_labels: dict = None, max_height: int = 500) -> None:
    """Render a DataFrame as a styled pro-table HTML element."""
    if df is None or df.empty:
        st.info("No data to display.")
        return
    display_df = df.copy()
    if col_labels:
        display_df = display_df.rename(columns=col_labels)
    caller_frame = inspect.stack()[1]
    key_seed = "|".join(
        [
            str(caller_frame.function),
            str(caller_frame.lineno),
            *(str(col) for col in display_df.columns),
        ]
    )
    table_key = f"table_view_{hashlib.md5(key_seed.encode('utf-8')).hexdigest()[:12]}"
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
    font_boost = 4 if is_drilldown_view else 0
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
    key_seed = "|".join(
        [
            str(caller_frame.function),
            str(caller_frame.lineno),
            *(str(col) for col in df.columns),
        ]
    )
    table_key = f"stdf_view_{hashlib.md5(key_seed.encode('utf-8')).hexdigest()[:12]}"
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
) -> None:
    """Main dashboard page with KPI tiles, zone chart, and plant table."""



    render_header()
    # ── Navigation Filters (Zone & Plant) ────────────────────────────────────
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

    # --- LOCATION VISIT KPI LOGIC (after navigation filters) ---
    df_loc_filtered = pd.DataFrame()
    loc_visit_missing_columns = []
    loc_visit_error = ""
    kpi_location_visit = 0
    kpi_location_compliance = 0.0

    if os.path.exists(LOCATION_VISIT_PATH):
        df_loc = load_location_visit(LOCATION_VISIT_PATH)
        required_loc_cols = [
            "Planning Plant", "Plant Desc.", "Audit Number", "Audit Start Date", "Audit End Date",
            "TotalRecomms", "ClosedRecomms", "OpenRecomms",
        ]
        loc_visit_missing_columns = [c for c in required_loc_cols if c not in df_loc.columns]

        if not loc_visit_missing_columns:
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

            if selected_zone != "All Zones":
                df_loc_work = df_loc_work[df_loc_work["Zone Name"] == selected_zone]
            if selected_plant != "All Plants":
                df_loc_work = df_loc_work[df_loc_work["Plant Name"] == selected_plant]

            df_loc_filtered = df_loc_work.copy()

            total_recomms = int(
                df_loc_filtered["TotalRecomms"]
                .astype(str)
                .str.strip()
                .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
                .dropna()
                .shape[0]
            )
            total_closed = int(
                df_loc_filtered["ClosedRecomms"]
                .astype(str)
                .str.strip()
                .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
                .dropna()
                .shape[0]
            )
            kpi_location_visit = int(df_loc_filtered["Planning Plant"].nunique())
            kpi_location_compliance = (total_closed / total_recomms) if total_recomms else 0.0
    else:
        loc_visit_error = "Location Visit file not found at Reports/LOCATION_VISIT.xls"

    # Dummy Tank KPI data prep
    dummy_tank_filtered = pd.DataFrame()
    dummy_tank_missing_columns = []
    dummy_tank_error = ""
    total_dummy_qty = 0.0

    if os.path.exists(DUMMY_TANK_PATH):
        df_dummy_tank = load_dummy_tank_stock(DUMMY_TANK_PATH)
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
        df_pipeline_stock = load_pipeline_stock(PIPELINE_STOCK_PATH)
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
        df_tank_turns_raw = load_tank_turns(TANK_TURNS_PATH)
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

    st.markdown('<div class="sec-title">&#128202; Exception Parameters &#8212; Live Summary</div>', unsafe_allow_html=True)

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
        st.write(f"[DEBUG] Pending DC Clicked: {clicked_dc}")
        if clicked_dc:
            st.session_state["page"] = "pending_dc_details"
            st.write(f"[DEBUG] Set page to: {st.session_state['page']}")
            st.rerun()

    with col2:
        s_df = open_delivery_result_filtered.get("summary_df", pd.DataFrame())
        total_deliveries = int(s_df["Open Delivery Count"].sum()) if not s_df.empty and "Open Delivery Count" in s_df.columns else 0
        clicked_open = render_open_delivery_tile({**open_delivery_result_filtered, "total_count": total_deliveries})
        st.write(f"[DEBUG] Open Delivery Clicked: {clicked_open}")
        if clicked_open:
            st.session_state["page"] = "open_delivery_details"
            st.write(f"[DEBUG] Set page to: {st.session_state['page']}")
            st.rerun()
    with col3:
        s_df = pending_invoices_result_filtered.get("summary_df", pd.DataFrame())
        total_invoices = int(s_df["Pending Invoice Count"].sum()) if not s_df.empty and "Pending Invoice Count" in s_df.columns else 0
        clicked_pending_inv = render_pending_invoices_tile({**pending_invoices_result_filtered, "total_count": total_invoices})
        st.write(f"[DEBUG] Pending Invoices Clicked: {clicked_pending_inv}")
        if clicked_pending_inv:
            st.session_state["page"] = "pending_invoices_details"
            st.write(f"[DEBUG] Set page to: {st.session_state['page']}")
            st.rerun()
    with col4:
        s_df = open_sales_orders_result_filtered.get("summary_df", pd.DataFrame())
        total_so = int(s_df["Open Sales Order Count"].sum()) if not s_df.empty and "Open Sales Order Count" in s_df.columns else 0
        clicked_open_so = render_open_sales_orders_tile({**open_sales_orders_result_filtered, "total_count": total_so})
        st.write(f"[DEBUG] Open Sales Orders Clicked: {clicked_open_so}")
        if clicked_open_so:
            st.session_state["page"] = "open_sales_orders_details"
            st.write(f"[DEBUG] Set page to: {st.session_state['page']}")
            st.rerun()

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

    # ── Row 2: KPI tiles ─────────────────────────────────────────────────────
    col5, col6, col7, col8 = st.columns(4, gap="small")
    with col5:
        s_df = open_intransit_result_filtered.get("summary_df", pd.DataFrame())
        total_intransit = int(s_df["Open In-Transit STO Count"].sum()) if not s_df.empty and "Open In-Transit STO Count" in s_df.columns else 0
        clicked_intransit = render_open_intransit_tile({**open_intransit_result_filtered, "total_count": total_intransit})
        st.write(f"[DEBUG] Open Intransit Clicked: {clicked_intransit}")
        if clicked_intransit:
            st.session_state["page"] = "open_intransit_details"
            st.write(f"[DEBUG] Set page to: {st.session_state['page']}")
            st.rerun()
    with col6:
        s_df = open_short_sales_result_filtered.get("summary_df", pd.DataFrame())
        total_short_sales = int(s_df["Total Shortage Quantity (in Ltrs)"].sum()) if not s_df.empty and "Total Shortage Quantity (in Ltrs)" in s_df.columns else 0
        clicked_short_sales = render_open_shortages_sales_tile({**open_short_sales_result_filtered, "total_count": total_short_sales})
        st.write(f"[DEBUG] Open Short Sales Clicked: {clicked_short_sales}")
        if clicked_short_sales:
            st.session_state["page"] = "open_shortages_sales_details"
            st.write(f"[DEBUG] Set page to: {st.session_state['page']}")
            st.rerun()
    with col7:
        s_df = open_short_sto_result_filtered.get("summary_df", pd.DataFrame())
        total_short_sto = int(s_df["Total STO Shortage Quantity (in Ltrs)"].sum()) if not s_df.empty and "Total STO Shortage Quantity (in Ltrs)" in s_df.columns else 0
        clicked_short_sto = render_open_shortages_sto_tile({**open_short_sto_result_filtered, "total_count": total_short_sto})
        st.write(f"[DEBUG] Open Short STO Clicked: {clicked_short_sto}")
        if clicked_short_sto:
            st.session_state["page"] = "open_shortages_sto_details"
            st.write(f"[DEBUG] Set page to: {st.session_state['page']}")
            st.rerun()
    with col8:
        s_df = tank_reco_result_filtered.get("summary_df", pd.DataFrame())
        total_tank_reco = int(s_df["Tank Reco Count"].sum()) if not s_df.empty and "Tank Reco Count" in s_df.columns else 0
        clicked_tank = render_tank_reco_tile({**tank_reco_result_filtered, "total_count": total_tank_reco})
        st.write(f"[DEBUG] Tank Reco Clicked: {clicked_tank}")
        if clicked_tank:
            st.session_state["page"] = "tank_reco_details"
            st.write(f"[DEBUG] Set page to: {st.session_state['page']}")
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
        st.write(f"[DEBUG] PL Unblock Clicked: {clicked_pl_unblock}")
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
        st.write(f"[DEBUG] Dummy Tank Clicked: {clicked_dummy_tank}")
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
        clicked_location_visit = kpi_card(
            label="Location Visit - %Compliance",
            value=f"{kpi_location_visit:,} | {kpi_location_compliance * 100:.1f}%",
            detail="Location Visits and Compliance",
            icon="&#128205;",
            color_class="c-success" if kpi_location_visit > 0 else "c-muted",
            key="location_visit_btn"
        )
        st.write(f"[DEBUG] Location Visit Clicked: {clicked_location_visit}")
        if clicked_location_visit:
            st.session_state["pl_unblock_clicked"] = False
            st.session_state["dummy_tank_clicked"] = False
            st.session_state["tank_turns_page"] = "main"
            st.session_state["selected_tile"] = None
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
        "Total Exceptions", "Pending DC", "Open Delivery", "Open In-Transit", "Open Sales Order", "Pending Invoice", "Tank Reco", "Shortage Sales (Billing Docs)", "Shortage STO (Billing Docs)"
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
        "Tank Reco",
        "Shortage Sales (Billing Docs)",
        "Shortage STO (Billing Docs)",
    ]

    frames = [
        _extract_zone_plant_metric(pending_dc_result.get("summary_df", pd.DataFrame()), "Pending DC Count", "Pending DC"),
        _extract_zone_plant_metric(open_delivery_result.get("summary_df", pd.DataFrame()), "Open Delivery Count", "Open Delivery"),
        _extract_zone_plant_metric(open_intransit_result.get("summary_df", pd.DataFrame()), "Open In-Transit STO Count", "Open In-Transit"),
        _extract_zone_plant_metric(open_sales_orders_result.get("summary_df", pd.DataFrame()), "Open Sales Order Count", "Open Sales Order"),
        _extract_zone_plant_metric(pending_invoices_result.get("summary_df", pd.DataFrame()), "Pending Invoice Count", "Pending Invoice"),
        _extract_zone_plant_metric(tank_reco_result.get("summary_df", pd.DataFrame()), "Tank Reco Count", "Tank Reco"),
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
        <p>Zone-wise and Plant-wise breakdown of all Pending Delivery Challans
        if not df.empty:
            _render_streamlit_dataframe(df, max_height=420, hide_index=True)
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
    if "tank_turns_page" not in st.session_state:
        st.session_state["tank_turns_page"] = "main"

    inject_css()

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

    try:
        load_zone_master()
    except Exception as exc:
        st.sidebar.warning(f"Zone master not loaded: {exc}")

    # Sidebar
    selected_zones, selected_plants, uploaded_dc, sidebar_system_info_slot = render_sidebar(df_plant)

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
        # Diagnostic summary for Pending DC file
        st.sidebar.markdown("---")
        st.sidebar.markdown("**Pending DC Diagnostics**")
        st.sidebar.markdown(f"Rows loaded: **{len(df_dc)}**")
        if not df_dc.empty:
            unique_pairs = df_dc.drop_duplicates(subset=["SENDING PLANT", "SHIPMENT"])
            st.sidebar.markdown(f"Unique (SENDING PLANT, SHIPMENT): **{len(unique_pairs)}**")
            if len(unique_pairs) < 1:
                st.sidebar.warning("⚠️ No unique Pending DC pairs found. Check file content.")
        else:
            st.sidebar.warning("⚠️ Pending DC file is empty.")
        pending_dc_result = process_pending_dc(
            df_dc,
            df_plant,
            zone_filter  = selected_zones  or None,
            plant_filter = selected_plants or None,
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
