"""
╔══════════════════════════════════════════════════════════════════════════╗
║               SOD Exception Dashboard  —  app.py                       ║
║  Hindustan Petroleum Corporation Limited (HPCL)                         ║
║  Supply & Operations Division — Exception Monitoring System             ║
╚══════════════════════════════════════════════════════════════════════════╝

Usage:
    streamlit run app.py

Tech Stack:
    Python | Streamlit | Pandas | OpenPyXL | Plotly
"""

import base64
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

# Default data file paths (fallback when no file is uploaded)
PENDING_DC_PATH     = os.path.join(BASE_DIR, "PENDING_DC_SOD.xlsx")
OPEN_DELIVERY_PATH  = os.path.join(BASE_DIR, "OPEN_DELIVERY.xls")
OPEN_INTRANSIT_PATH = os.path.join(BASE_DIR, "OPEN_INTRANSIT_SOD.xls")
OPEN_SO_PATH        = os.path.join(BASE_DIR, "OPEN_SALES_ORDER.xls")
PEND_INV_PATH       = os.path.join(BASE_DIR, "PENDING_INVOICES_SOD.xls")
SHORT_SALES_PATH    = os.path.join(BASE_DIR, "SOD_OPEN_SHORTAGES_SALES.xls")
SHORT_STO_PATH      = os.path.join(BASE_DIR, "SOD_OPEN_SHORTAGES_STO.xls")
TANK_RECO_PATH      = os.path.join(BASE_DIR, "TANK_RECO_REPORT.xls")

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
    .main .block-container {{
        padding-top: 0 !important;
        padding-bottom: 1.5rem;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100%;
    }}

    /* ── Full-Width Title Banner (≈ 2 inches / 192 px tall) ── */
    .hpcl-banner-wrap {{
        margin: 0 -1rem;
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
        padding: 4px 10px;
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
    """
    df = pd.read_excel(PLANT_MASTER_PATH, dtype={"Plant Code": str}, engine="openpyxl")
    df.columns = df.columns.str.strip()
    df["Plant Code"] = df["Plant Code"].astype(str).str.strip()
    df["Plant Name"] = df["Plant Name"].astype(str).str.strip()
    df["Zone Name"]  = df["Zone Name"].astype(str).str.strip()
    if "Active" in df.columns:
        df = df[df["Active"].astype(str).str.strip().str.lower() == "yes"]
    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_zone_master() -> pd.DataFrame:
    """Load Zonewise MaiID Master.xlsx from disk (cached)."""
    df = pd.read_excel(ZONE_MASTER_PATH, engine="openpyxl")
    df.columns = df.columns.str.strip()
    df["Zone Name"] = df["Zone Name"].astype(str).str.strip()
    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def _load_excel_from_path(path: str) -> pd.DataFrame:
    """Internal helper: load any Excel file from a disk path (cached)."""
    ext = os.path.splitext(path)[1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    df = pd.read_excel(path, engine=engine)
    df.columns = df.columns.str.strip().str.upper()
    return df


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
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        df = pd.read_excel(source, engine=engine)
        df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as exc:
        st.error(f"❌ Error loading Pending DC file: {exc}")
        return pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
# DATA PROCESSING
# ─────────────────────────────────────────────────────────────────────────────

def process_pending_dc(
    df_dc        : pd.DataFrame,
    df_plant     : pd.DataFrame,
    zone_filter  : list = None,
    plant_filter : list = None,
) -> dict:
    """
    Process raw Pending DC data into aggregated exception metrics.

    1. De-dup on (SENDING PLANT, SHIPMENT):  each unique shipment = 1 pending DC.
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

    # Step 1: de-duplicate FOR COUNTING — each unique (SENDING PLANT, SHIPMENT) = 1 Pending DC
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

    mask_unmatched  = merged["Plant Name"].isna()
    unmatched_codes = merged.loc[mask_unmatched, "SENDING PLANT"].unique().tolist()
    merged.loc[mask_unmatched, "Plant Name"] = (
        "Unknown (" + merged.loc[mask_unmatched, "SENDING PLANT"] + ")"
    )
    merged.loc[merged["Zone Name"].isna(), "Zone Name"] = "Unmapped"

    # Step 2b: merge detail (material-level) with PlantMaster for display
    detail_merged = df_detail.merge(
        plant_map,
        left_on  = "SENDING PLANT",
        right_on = "Plant Code",
        how      = "left",
    )
    detail_merged.loc[detail_merged["Plant Name"].isna(), "Plant Name"] = (
        "Unknown (" + detail_merged.loc[detail_merged["Plant Name"].isna(), "SENDING PLANT"] + ")"
    )
    detail_merged.loc[detail_merged["Zone Name"].isna(), "Zone Name"] = "Unmapped"

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
        "unmatched"   : unmatched_codes,
    }


# ─────────────────────────────────────────────────────────────────────────────
# UI HELPER COMPONENTS
# ─────────────────────────────────────────────────────────────────────────────

def render_header(subtitle: str = "") -> None:
    """
    Render the two-part HPCL page header:
      1. Full-width brand banner image  (≈ 2 inches / 192 px tall).
                 Uses Title.png if available, then Master Logo.jpg, then pure-CSS.
      2. Dark-blue info strip: app title, subtitle, date/time.
    """
    now  = datetime.now().strftime("%d %b %Y  |  %I:%M %p")
    subtitle_html = f'<p class="dash-header-sub">{subtitle}</p>' if subtitle else ""

    # ── Part 1: Full-width brand banner ─────────────────────────────────────
    title_uri = ""
    logo_uri  = ""
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

    st.markdown(banner_html, unsafe_allow_html=True)

    # ── Part 2: Dark-blue info strip ────────────────────────────────────────
    header_html = (
        '<div class="dash-header">'
        '<div class="dash-header-main">'
        '<p class="dash-header-title">SOD Exception Dashboard</p>'
        f'{subtitle_html}'
        '</div>'
        '<div class="dash-header-meta">'
        f'<div style="margin-top:4px;font-size:18px;">{now}</div>'
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


def export_to_excel(df_dict: dict) -> bytes:
    """Serialise {sheet_name: DataFrame} to Excel bytes for st.download_button."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet, df in df_dict.items():
            if df is not None and not df.empty:
                df.to_excel(writer, index=False, sheet_name=sheet[:31])
    buf.seek(0)
    return buf.getvalue()


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

def render_sidebar(df_plant: pd.DataFrame) -> tuple:
    """Render navigation sidebar. Returns (zones, plants, uploaded_file)."""
    with st.sidebar:
        sidebar_logo_html = '<div style="font-size:2.6rem;">&#9981;</div>'
        try:
            if os.path.exists(LOGO_IMG_PATH):
                logo_uri = _load_img_b64(LOGO_IMG_PATH)
                sidebar_logo_html = (
                    f'<img src="{logo_uri}" alt="HPCL Corporate Logo" '
                    'style="height:52px;width:auto;display:block;margin:0 auto 6px auto;'
                    'object-fit:contain;" />'
                )
        except Exception:
            pass

        st.markdown(f"""
        <div style="text-align:center;padding:14px 0 6px;">
            {sidebar_logo_html}
            <div style="font-size:1.2rem;font-weight:700;letter-spacing:.06em;
                        color:#FFFFFF;">HPCL</div>
            <div style="font-size:0.75rem;opacity:.7;color:#AACCFF;">
                Exception Monitoring</div>
        </div>
        <hr/>
        """, unsafe_allow_html=True)

        st.markdown('<p class="sb-nav-lbl">&#128205; Navigation Filters</p>',
                    unsafe_allow_html=True)

        all_zones = sorted(df_plant["Zone Name"].dropna().unique().tolist())
        selected_zones = st.multiselect(
            "Zone",
            options     = all_zones,
            default     = [],
            placeholder = "All Zones",
        )

        if selected_zones:
            avail_plants = (
                df_plant[df_plant["Zone Name"].isin(selected_zones)]
                ["Plant Name"].dropna().unique().tolist()
            )
        else:
            avail_plants = df_plant["Plant Name"].dropna().unique().tolist()

        selected_plants = st.multiselect(
            "Plant / Location",
            options     = sorted(avail_plants),
            default     = [],
            placeholder = "All Plants",
        )

        st.markdown("<hr/>", unsafe_allow_html=True)

        st.markdown('<p class="sb-nav-lbl">&#128194; Data Upload</p>',
                    unsafe_allow_html=True)
        uploaded_dc = st.file_uploader(
            "Pending DC File  (.xls / .xlsx)",
            type = ["xls", "xlsx"],
            key  = "uploader_pending_dc",
        )

        st.markdown("<hr/>", unsafe_allow_html=True)

        st.markdown('<p class="sb-nav-lbl">&#8505;&#65039; System Info</p>',
                    unsafe_allow_html=True)
        st.markdown(f"""
        <div style="font-size:14px;line-height:2.1;opacity:.90;">
            &#127981; &nbsp;Active Plants : <b>{len(df_plant)}</b><br/>
            &#128506; &nbsp;Zones         : <b>{df_plant["Zone Name"].nunique()}</b><br/>
            &#128197; &nbsp;Date          : <b>{datetime.now().strftime('%d %b %Y')}</b>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<hr/>", unsafe_allow_html=True)
        if st.button("&#128260; Refresh Data", width='stretch',
                     key="btn_refresh"):
            st.cache_data.clear()
            st.rerun()

    return selected_zones, selected_plants, uploaded_dc


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
    headers_html = "".join(
        f"<th>{html.escape(str(c))}</th>" for c in display_df.columns
    )
    rows_html = "".join(
        "<tr>"
        + "".join(
            f"<td>{html.escape(str(v) if pd.notna(v) else '')}</td>"
            for v in row
        )
        + "</tr>"
        for _, row in display_df.iterrows()
    )
    st.markdown(
        f'<div class="pro-table-wrap" style="max-height:{max_height}px;">'
        f'<table class="pro-table">'
        f'<thead><tr>{headers_html}</tr></thead>'
        f'<tbody>{rows_html}</tbody>'
        f'</table></div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: MAIN DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────

def render_dashboard(
    pending_dc_result : dict,
    zone_filter       : list,
    plant_filter      : list,
) -> None:
    """Main dashboard page with KPI tiles, zone chart, and plant table."""
    render_header()

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

    # ── Row 1: KPI tiles ─────────────────────────────────────────────────────
    st.markdown(
        '<div class="sec-title">&#128202; Exception Parameters &#8212; Live Summary</div>',
        unsafe_allow_html=True,
    )

    col1, col2, col3, col4 = st.columns(4, gap="small")

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
        kpi_card("Open Deliveries",    0, "Coming Soon", "&#128230;", "c-muted", "tile_open_del")
    with col3:
        kpi_card("Pending Invoices",   0, "Coming Soon", "&#129534;", "c-muted", "tile_pend_inv")
    with col4:
        kpi_card("Open Sales Orders",  0, "Coming Soon", "&#128203;", "c-muted", "tile_open_so")

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

    # ── Row 2: KPI tiles ─────────────────────────────────────────────────────
    col5, col6, col7, col8 = st.columns(4, gap="small")
    with col5:
        kpi_card("Open Intransit",         0, "Coming Soon", "&#128699;", "c-muted", "tile_intrans")
    with col6:
        kpi_card("Open Shortages (Sales)", 0, "Coming Soon", "&#9888;",   "c-muted", "tile_sh_sal")
    with col7:
        kpi_card("Open Shortages (STO)",   0, "Coming Soon", "&#9888;",   "c-muted", "tile_sh_sto")
    with col8:
        kpi_card("Tank Reco",              0, "Coming Soon", "&#128738;", "c-muted", "tile_tank")

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # ── Zone bar chart ───────────────────────────────────────────────────────
    zone_summary = pending_dc_result.get("zone_summary", pd.DataFrame())
    if not zone_summary.empty:
        st.markdown(
            "<div class='sec-title'>&#128205; Zone-wise Pending DC's</div>",
            unsafe_allow_html=True,
        )
        _render_zone_bar(zone_summary)

    # ── Plant-wise summary table ──────────────────────────────────────────────
    summary_df = pending_dc_result.get("summary_df", pd.DataFrame())
    if not summary_df.empty:
        st.markdown(
            "<div class='sec-title'>&#127981; Plant-wise Summary</div>",
            unsafe_allow_html=True,
        )
        with st.expander("View Plant-wise Table", expanded=True):
            render_professional_summary_table(summary_df)

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


def _render_zone_bar(zone_summary: pd.DataFrame) -> None:
    """Render polished zone visuals with a colorful bar chart and donut chart."""
    chart_df = zone_summary.copy().sort_values("Pending DC Count", ascending=False)
    palette = [
        "#0B3D91", "#1B66C9", "#2A9D8F", "#F4A261", "#E76F51",
        "#7B2CBF", "#3A86FF", "#43AA8B", "#F9C74F", "#577590",
        "#90BE6D", "#F94144",
    ]
    chart_df["Color"] = [palette[idx % len(palette)] for idx in range(len(chart_df))]

    bar_col, pie_col = st.columns([2.2, 1], gap="medium")

    with bar_col:
        fig_bar = px.bar(
            chart_df,
            x="Zone Name",
            y="Pending DC Count",
            text="Pending DC Count",
            labels={"Pending DC Count": "Pending DCs", "Zone Name": "Zone"},
        )
        fig_bar.update_traces(
            marker_color=chart_df["Color"],
            marker_line_color="#FFFFFF",
            marker_line_width=1.5,
            textposition="outside",
            textfont_size=20,
            textfont_color="#163A63",
            hovertemplate="<b>%{x}</b><br>Pending DCs: %{y}<extra></extra>",
        )
        fig_bar.update_layout(
            plot_bgcolor="#F8FAFD",
            paper_bgcolor="white",
            font=dict(family="Segoe UI", size=20, color="#163A63"),
            margin=dict(l=10, r=10, t=18, b=95),
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
                title="Pending DCs",
                title_font=dict(size=18, color="#163A63"),
                tickfont=dict(size=18, color="#42566E"),
                gridcolor="#DCE6F2",
                zeroline=False,
            ),
        )
        st.plotly_chart(fig_bar, width='stretch', config={"displayModeBar": False})

    with pie_col:
        pie_df = chart_df.copy()
        if len(pie_df) > 6:
            top_df = pie_df.head(5).copy()
            other_value = int(pie_df.iloc[5:]["Pending DC Count"].sum())
            if other_value > 0:
                top_df.loc[len(top_df)] = {
                    "Zone Name": "Other Zones",
                    "Plants": pie_df.iloc[5:]["Plants"].sum() if "Plants" in pie_df.columns else 0,
                    "Pending DC Count": other_value,
                    "Color": "#C7D3E3",
                }
            pie_df = top_df

        fig_pie = px.pie(
            pie_df,
            names="Zone Name",
            values="Pending DC Count",
            hole=0.58,
            color="Zone Name",
            color_discrete_sequence=pie_df["Color"].tolist(),
        )
        fig_pie.update_traces(
            textposition="inside",
            textinfo="percent",
            textfont_size=18,
            textfont_color="#FFFFFF",
            marker=dict(line=dict(color="white", width=2)),
            hovertemplate="<b>%{label}</b><br>Pending DCs: %{value}<br>Share: %{percent}<extra></extra>",
        )
        fig_pie.update_layout(
            paper_bgcolor="white",
            plot_bgcolor="white",
            font=dict(family="Segoe UI", size=18, color="#163A63"),
            margin=dict(l=10, r=10, t=18, b=16),
            height=430,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.22,
                xanchor="center",
                x=0.5,
                font=dict(size=18, color="#163A63"),
            ),
            annotations=[
                dict(
                    text=f"<b>{int(chart_df['Pending DC Count'].sum())}</b><br>Total DCs",
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                    font=dict(size=24, color="#0B3D91"),
                )
            ],
        )
        st.plotly_chart(fig_pie, width='stretch', config={"displayModeBar": False})


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
        (counted as unique Shipment numbers per Sending Plant)</p>
    </div>
    """, unsafe_allow_html=True)

    summary_df   = pending_dc_result.get("summary_df",   pd.DataFrame())
    zone_summary = pending_dc_result.get("zone_summary",  pd.DataFrame())
    detail_df    = pending_dc_result.get("detail_df",    pd.DataFrame())
    total_count  = pending_dc_result.get("total_count",  0)

    if summary_df.empty:
        st.info("&#8505; No data available for the current filter selection.")
        return

    # ── Metric row ────────────────────────────────────────────────────────────
    total_zones  = int(summary_df["Zone Name"].nunique())
    total_plants = int(summary_df["Plant Name"].nunique())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Pending DC's (Total)", f"{total_count:,}",
              help="Count of unique shipments")
    m2.metric("Zones Affected",  f"{total_zones}")
    m3.metric("Plants Affected", f"{total_plants}")

    if not detail_df.empty and "QUANTITY" in detail_df.columns:
        try:
            dq = detail_df["QUANTITY"].sum()
            m4.metric("Total Qty (L)", f"{dq:,.0f}")
        except Exception:
            m4.metric("Total Qty (L)", "N/A")

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


# ─────────────────────────────────────────────────────────────────────────────
# APPLICATION ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    """Bootstrap: session state → CSS → master data → sidebar → data → page."""

    if "page" not in st.session_state:
        st.session_state["page"] = "dashboard"

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
    selected_zones, selected_plants, uploaded_dc = render_sidebar(df_plant)

    # Resolve data source
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

    # Page router
    page = st.session_state.get("page", "dashboard")

    if page == "dashboard":
        render_dashboard(pending_dc_result, selected_zones, selected_plants)
    elif page == "pending_dc_details":
        render_pending_dc_details(pending_dc_result, selected_zones, selected_plants)
    else:
        st.session_state["page"] = "dashboard"
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    main()
