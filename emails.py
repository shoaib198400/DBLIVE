"""HPCL LIVEDB — Exception mail notifications via Microsoft Outlook COM.

Emails are sent through the Outlook desktop application running on the same
Windows machine as this Streamlit app.  No SMTP credentials needed —
Outlook uses the currently signed-in HPCL account.

Requirements (local only):
  • Microsoft Outlook must be installed and OPEN on this PC
  • shoaibrehman@hpcl.in must be the active Outlook account
  • pywin32 installed:  pip install pywin32

On Streamlit Cloud (Linux), Outlook COM is unavailable — the UI shows a
notice and disables the Send button automatically.
"""

from __future__ import annotations

import gc
import io
import os
import platform
import re
import tempfile
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd

# ── Constants ─────────────────────────────────────────────────────────────────

SENDER_EMAIL = "shoaibrehman@hpcl.in"

BCC_EMAILS: List[str] = [
    "SOD.OPNS.HQO@hpcl.in",
    "bhsgk@hpcl.in",
    "shubham.tayal@hpcl.in",
]

# ── Zone → recipient map ──────────────────────────────────────────────────────

ZONE_EMAIL_MAP: Dict[str, Dict[str, str]] = {
    "Bengaluru Zone":       {"to": "brijeshkumar@hpcl.in",              "cc": "BLR.OND.IC@hpcl.in"},
    "Bhopal Zone":          {"to": "agajare@hpcl.in;twinacore@hpcl.in", "cc": "CZ.OND.IC@hpcl.in"},
    "Bhubaneshwar Zone":    {"to": "smarak.lenka@hpcl.in",              "cc": "ECZ.OND.IC@hpcl.in"},
    "Chandigarh Zone":      {"to": "haroonhamid@hpcl.in",               "cc": "NFZ.OND.IC@hpcl.in"},
    "Cochin Zone":          {"to": "kathir@hpcl.in",                    "cc": "kbanothu@hpcl.in"},
    "East Zone":            {"to": "sray@hpcl.in",                      "cc": "EZ.OND.IC@hpcl.in"},
    "Guwahati Zone":        {"to": "lodyuo@hpcl.in",                    "cc": "gurubachansingha@hpcl.in"},
    "Jaipur Zone":          {"to": "rjprasad@hpcl.in",                  "cc": "NWF.OND.IC@hpcl.in"},
    "Noida (UP-West) Zone": {"to": "aradhnat@hpcl.in",                  "cc": "chraghu@hpcl.in"},
    "North Central Zone":   {"to": "rvpandey@hpcl.in",                  "cc": "adeshmukh@hpcl.in"},
    "North West Zone":      {"to": "sanjaykdewangan@hpcl.in",           "cc": "NWZ.OND.IC@hpcl.in"},
    "North Zone (NZ)":      {"to": "ajaygr@hpcl.in",                    "cc": "NZ.OND.IC@hpcl.in"},
    "Patna Zone":           {"to": "ajaisingh@hpcl.in",                 "cc": "dastidar@hpcl.in"},
    "South Central Zone":   {"to": "suryabv@hpcl.in",                   "cc": "SCRZ.OND.IC@hpcl.in"},
    "South Zone":           {"to": "venkates@hpcl.in",                  "cc": "SZ.OND.IC@hpcl.in"},
    "West Zone (WZ)":       {"to": "ntgajbiye@hpcl.in",                 "cc": "WZ.OND.IC@hpcl.in"},
}

# ── Exception labels (key = column name in all_exception_plant_df) ────────────

EXCEPTION_LABELS: Dict[str, str] = {
    "Pending DC":                    "Pending DCs",
    "Open Delivery":                 "Open Deliveries",
    "Open In-Transit":               "Open In-Transit STOs",
    "Open Sales Order":              "Open Sales Orders",
    "Pending Invoice":               "Pending Invoices",
    "Shortage Sales (Billing Docs)": "Shortages - Sales",
    "Shortage STO (Billing Docs)":   "Shortages - STO",
}

# ── Outlook availability check ────────────────────────────────────────────────

def outlook_available() -> tuple[bool, str]:
    """Return (True, '') if Outlook COM can be used; else (False, reason)."""
    if platform.system() != "Windows":
        return False, "Outlook COM is only available on Windows (not on Streamlit Cloud)."
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        return True, ""
    except ImportError:
        return False, "pywin32 not installed. Run: pip install pywin32"


# ── Safe filename helper ──────────────────────────────────────────────────────

def _safe_name(s: str) -> str:
    """Strip characters not allowed in Windows filenames."""
    return re.sub(r'[\\/:*?"<>|]', '', s).strip()


# ── Low-level Outlook sender ──────────────────────────────────────────────────

def _send_via_outlook(
    to: str,
    subject: str,
    html_body: str,
    cc: str = "",
    bcc: str = "",
    attachments: Optional[List[Tuple[str, bytes]]] = None,
) -> dict:
    """Send one email through Outlook COM.

    attachments: list of (display_filename, bytes) tuples
    Returns {"ok": True/False, "mode": "sent"/"draft", "msg": "..."}
    """
    import pythoncom
    import win32com.client as win32

    com_init = False
    outlook = mail_item = None
    tmp_paths: List[str] = []

    try:
        pythoncom.CoInitialize()
        com_init = True

        outlook = win32.Dispatch("Outlook.Application")
        mail_item = outlook.CreateItem(0)
        mail_item.To      = to
        mail_item.Subject = subject
        mail_item.HTMLBody = html_body
        if cc.strip():
            mail_item.CC = cc
        if bcc.strip():
            mail_item.BCC = bcc

        if attachments:
            for fname, fbytes in attachments:
                ext = os.path.splitext(fname)[1] or ".xlsx"
                tmp = tempfile.NamedTemporaryFile(
                    delete=False, suffix=ext, prefix="hpcl_"
                )
                tmp.write(fbytes)
                tmp.close()
                tmp_paths.append((tmp.name, fname))

            for tmp_path, display_name in tmp_paths:
                mail_item.Attachments.Add(tmp_path, 1, 1, display_name)

        try:
            mail_item.Send()
            return {"ok": True, "mode": "sent", "msg": "Mail sent successfully."}
        except Exception as send_exc:
            try:
                mail_item.Save()
                return {"ok": True, "mode": "draft",
                        "msg": f"Saved to Drafts (Send failed: {send_exc})"}
            except Exception as draft_exc:
                return {"ok": False,
                        "msg": f"Send and draft both failed: {draft_exc}"}

    except Exception as exc:
        return {"ok": False, "msg": f"Outlook error: {exc}"}

    finally:
        try:
            del mail_item
        except Exception:
            pass
        try:
            del outlook
        except Exception:
            pass
        gc.collect()
        if com_init:
            try:
                import pythoncom as _pc
                _pc.CoUninitialize()
            except Exception:
                pass
        for tmp_path, _ in tmp_paths:
            try:
                os.remove(tmp_path)
            except Exception:
                pass


# ── Excel builder — one file per exception type ───────────────────────────────

def build_zone_excel_attachments(
    zone_name: str,
    detail_dfs: Dict[str, pd.DataFrame],
    selected_exceptions: List[str],
) -> List[Tuple[str, bytes]]:
    """Build one Excel file per selected exception type for a zone.

    Returns list of (filename, bytes) tuples ready to attach.
    Filename format:  ZoneName_ExceptionLabel.xlsx
    Only exception types that have rows for this zone are included.
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    results: List[Tuple[str, bytes]] = []

    for exc_key in selected_exceptions:
        label = EXCEPTION_LABELS.get(exc_key, exc_key)
        detail_df = detail_dfs.get(exc_key)

        if detail_df is None or not isinstance(detail_df, pd.DataFrame) or detail_df.empty:
            continue

        # Filter to this zone
        if "Zone Name" in detail_df.columns:
            zone_df = detail_df[detail_df["Zone Name"] == zone_name].copy()
        else:
            zone_df = detail_df.copy()  # can't filter — include all

        if zone_df.empty:
            continue

        # Drop purely internal helper columns
        drop_cols = [c for c in zone_df.columns
                     if c.startswith("_") or c in {"index", "level_0"}]
        zone_df = zone_df.drop(columns=drop_cols, errors="ignore").reset_index(drop=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            zone_df.to_excel(writer, index=False, sheet_name=label[:31])
            ws = writer.sheets[label[:31]]

            hdr_fill = PatternFill("solid", fgColor="003087")
            hdr_font = Font(bold=True, color="FFFFFF", size=10)
            thin      = Side(style="thin", color="D5E2F3")
            border    = Border(left=thin, right=thin, top=thin, bottom=thin)
            center    = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for cell in ws[1]:
                cell.fill      = hdr_fill
                cell.font      = hdr_font
                cell.alignment = center
                cell.border    = border

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border    = border

            for col in ws.columns:
                max_len = max(
                    (len(str(c.value or "")) for c in col),
                    default=10
                )
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 35)

            ws.row_dimensions[1].height = 28

        safe_zone  = _safe_name(zone_name)
        safe_label = _safe_name(label)
        filename   = f"{safe_zone}_{safe_label}.xlsx"
        results.append((filename, buf.getvalue()))

    return results


# ── HTML email builder ────────────────────────────────────────────────────────

def _kpi_tile_html(label: str, value: str, sub: str, border_color: str, bg: str) -> str:
    """Render one KPI tile cell (email-safe, table-based, no flex/grid)."""
    return (
        f'<td width="25%" style="padding:4px;">'
        f'<table width="100%" cellpadding="0" cellspacing="0">'
        f'<tr><td style="background:{bg};border-top:4px solid {border_color};'
        f'border-radius:8px;padding:12px 10px 10px;text-align:center;">'
        f'<div style="font-size:9px;font-weight:700;color:#666;text-transform:uppercase;'
        f'letter-spacing:.07em;margin-bottom:5px;">{label}</div>'
        f'<div style="font-size:20px;font-weight:900;color:{border_color};'
        f'line-height:1.1;">{value}</div>'
        f'<div style="font-size:9px;color:#888;margin-top:3px;">{sub}</div>'
        f'</td></tr></table>'
        f'</td>'
    )


def build_exception_email_html(
    zone_name: str,
    as_of_date: str,
    exception_summary_df: pd.DataFrame,
    selected_exceptions: List[str],
    custom_intro: str = "",
    attachment_names: Optional[List[str]] = None,
    zone_kpi_dict: Optional[dict] = None,
) -> str:
    """Build HPCL-branded HTML email body for one zone's exceptions.

    zone_kpi_dict: pre-computed per-zone KPI values dict (from _build_zone_kpi_dict).
    """

    primary   = "#003087"
    secondary = "#0057A8"
    accent    = "#FFD700"
    as_of     = as_of_date or datetime.now().strftime("%d %b %Y")
    kpis      = zone_kpi_dict or {}

    # Only include columns that exist and are selected
    exc_cols = [c for c in selected_exceptions if c in exception_summary_df.columns]

    if exc_cols:
        mask = (
            exception_summary_df[exc_cols]
            .apply(pd.to_numeric, errors="coerce")
            .fillna(0)
            .sum(axis=1) > 0
        )
        zone_df = exception_summary_df[mask].copy()
    else:
        zone_df = exception_summary_df.copy()

    total_locations = len(zone_df)

    # ── KPI Tiles block ────────────────────────────────────────────────────────
    def _n(key, default=0):
        return kpis.get(key, default)

    def _tile_row(tiles):
        cells = "".join(
            _kpi_tile_html(t["label"], t["value"], t["sub"], t["color"], t["bg"])
            for t in tiles
        )
        return f'<tr>{cells}</tr>'

    def _fmt_count(v):
        try:
            return f"{int(v):,}"
        except Exception:
            return str(v)

    def _fmt_ltrs(v):
        try:
            n = float(v)
            if n >= 1_000_000:
                return f"{n/1_000_000:.2f}M L"
            if n >= 1_000:
                return f"{n/1_000:.1f}K L"
            return f"{n:,.0f} L"
        except Exception:
            return str(v)

    row1_tiles = [
        {"label": "Pending DCs",        "value": _fmt_count(_n("Pending DC")),
         "sub": "Unique shipments",     "color": "#CC2929", "bg": "#FFF0F0"},
        {"label": "Open Deliveries",    "value": _fmt_count(_n("Open Delivery")),
         "sub": "Unique deliveries",    "color": "#CC2929", "bg": "#FFF0F0"},
        {"label": "Open In-Transit",    "value": _fmt_count(_n("Open In-Transit")),
         "sub": "Unique STO orders",    "color": "#0369A1", "bg": "#EFF8FF"},
        {"label": "Open Sales Orders",  "value": _fmt_count(_n("Open Sales Order")),
         "sub": "Unique sales docs",    "color": "#D97706", "bg": "#FFFBEB"},
    ]
    row2_tiles = [
        {"label": "Pending Invoices",   "value": _fmt_count(_n("Pending Invoice")),
         "sub": "Unique deliveries",    "color": "#D97706", "bg": "#FFFBEB"},
        {"label": "Shortage Sales",     "value": _fmt_ltrs(_n("Shortage Sales (Ltrs)")),
         "sub": "Total Litres",         "color": "#CC2929", "bg": "#FFF0F0"},
        {"label": "Shortage STO",       "value": _fmt_ltrs(_n("Shortage STO (Ltrs)")),
         "sub": "Total Litres",         "color": "#CC2929", "bg": "#FFF0F0"},
        {"label": "Tank Reco",          "value": _fmt_count(_n("Tank Reco")),
         "sub": "Plant+Tank+Material",  "color": "#7C3AED", "bg": "#F5F3FF"},
    ]
    row3_tiles = [
        {"label": "PL Unblock Qty",
         "value": f"{_n('PL Unblock (KL)', 0.0):.3f} KL",
         "sub": "Pipeline stock",       "color": "#0369A1", "bg": "#EFF8FF"},
        {"label": "Dummy Tank Qty",
         "value": f"{_n('Dummy Tank (KL)', 0.0):.3f} KL",
         "sub": "Excl. DBIT/DLUB/DSLP","color": "#D97706", "bg": "#FFFBEB"},
        {"label": "Tank Turns",
         "value": f"{_n('Tank Turns', 0.0):.2f}",
         "sub": "Dispatches÷Capacity",  "color": "#15803D", "bg": "#F0FDF4"},
        {"label": "Location Visit",
         "value": f"{int(_n('Locations Visited', 0))} loc | {_n('Location Compliance (%)', 0.0):.1f}%",
         "sub": "Visited | Compliance", "color": "#15803D", "bg": "#F0FDF4"},
    ]

    kpi_tiles_html = (
        '<table width="100%" cellpadding="0" cellspacing="0">'
        + _tile_row(row1_tiles)
        + '<tr><td colspan="4" style="height:6px;"></td></tr>'
        + _tile_row(row2_tiles)
        + '<tr><td colspan="4" style="height:6px;"></td></tr>'
        + _tile_row(row3_tiles)
        + '</table>'
    ) if kpis else ""

    # ── Location-wise exception table ──────────────────────────────────────────
    col_heads = "".join(
        f'<th style="background:{secondary};color:#fff;padding:7px 8px;'
        f'font-size:10px;font-weight:700;text-align:center;'
        f'border:1px solid rgba(255,255,255,0.2);white-space:nowrap;">'
        f'{EXCEPTION_LABELS.get(c, c)}</th>'
        for c in exc_cols
    )
    table_header = (
        f'<tr>'
        f'<th style="background:{primary};color:{accent};padding:7px 10px;'
        f'font-size:10px;font-weight:700;text-align:left;'
        f'border:1px solid rgba(255,255,255,0.2);">Location</th>'
        f'{col_heads}'
        f'<th style="background:{primary};color:{accent};padding:7px 8px;'
        f'font-size:10px;font-weight:700;text-align:center;'
        f'border:1px solid rgba(255,255,255,0.2);">Total</th>'
        f'</tr>'
    )

    def _fmt_cell(v):
        try:
            n = int(float(v))
            color  = "#CC0000" if n > 0 else "#007700"
            weight = "700" if n > 0 else "400"
            return f'<span style="color:{color};font-weight:{weight};">{n:,}</span>'
        except Exception:
            return str(v)

    rows_html = ""
    for i, (_, row) in enumerate(zone_df.iterrows()):
        bg    = "#FFFFFF" if i % 2 == 0 else "#F4F8FF"
        cells = "".join(
            f'<td style="padding:6px 8px;font-size:10px;text-align:center;'
            f'border:1px solid #E2EAF4;background:{bg};">'
            f'{_fmt_cell(row.get(c, 0))}</td>'
            for c in exc_cols
        )
        total_val  = int(pd.to_numeric(row.get("Total Exceptions", 0), errors="coerce") or 0)
        plant_name = row.get("Plant Name", row.get("Location", ""))
        rows_html += (
            f'<tr>'
            f'<td style="padding:6px 10px;font-size:10px;font-weight:600;'
            f'color:{primary};border:1px solid #E2EAF4;background:{bg};">'
            f'{plant_name}</td>'
            f'{cells}'
            f'<td style="padding:6px 8px;font-size:10px;font-weight:700;'
            f'text-align:center;border:1px solid #E2EAF4;background:{bg};'
            f'color:#CC0000;">{total_val:,}</td>'
            f'</tr>'
        )

    if not rows_html:
        rows_html = (
            f'<tr><td colspan="{len(exc_cols)+2}" style="text-align:center;'
            f'padding:16px;color:#666;font-size:11px;">'
            f'No location-level exceptions found for this zone.</td></tr>'
        )

    # ── Attachments list ───────────────────────────────────────────────────────
    if attachment_names:
        attach_items = "".join(
            f'<li style="margin:3px 0;font-size:11px;">{n}</li>'
            for n in attachment_names
        )
        attach_block = (
            f'<tr><td style="padding:10px 28px 16px;">'
            f'<p style="margin:0 0 5px;font-size:12px;font-weight:700;color:{primary};">'
            f'&#128206; Attachments:</p>'
            f'<ul style="margin:0;padding-left:18px;color:#333;">{attach_items}</ul>'
            f'</td></tr>'
        )
    else:
        attach_block = ""

    intro_block = (
        f'<p style="margin:0 0 12px;font-size:12px;line-height:1.6;color:#333;">'
        f'{custom_intro}</p>'
        if custom_intro else ""
    )

    kpi_section = (
        f'<!-- KPI Summary Tiles -->'
        f'<tr><td style="padding:4px 28px 2px;">'
        f'<p style="margin:0 0 6px;font-size:12px;font-weight:700;color:{primary};">'
        f'&#128202; Zone KPI Summary — {as_of}</p>'
        f'{kpi_tiles_html}'
        f'</td></tr>'
    ) if kpi_tiles_html else ""

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#F0F2F6;
             font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0"
       style="background:#F0F2F6;padding:20px 0;">
  <tr><td align="center">
  <table width="700" cellpadding="0" cellspacing="0"
         style="background:#fff;border-radius:12px;overflow:hidden;
                box-shadow:0 4px 20px rgba(0,48,135,0.12);">

    <!-- Header strip -->
    <tr><td style="background:linear-gradient(135deg,{primary} 0%,{secondary} 100%);
                   padding:18px 28px;">
      <table width="100%" cellpadding="0" cellspacing="0"><tr>
        <td>
          <div style="font-size:20px;font-weight:900;color:#fff;letter-spacing:0.04em;">
            &#9981;&nbsp; HPCL — SOD Exception Alert
          </div>
          <div style="font-size:12px;color:{accent};margin-top:3px;font-weight:600;">
            Hindustan Petroleum Corporation Limited
          </div>
        </td>
        <td align="right">
          <span style="font-size:11px;color:rgba(255,255,255,0.80);">Data as of: {as_of}</span>
        </td>
      </tr></table>
    </td></tr>

    <!-- Zone pill -->
    <tr><td style="padding:14px 28px 6px;">
      <div style="display:inline-block;background:#E8F0FE;
                  border-left:4px solid {primary};border-radius:6px;
                  padding:8px 16px;">
        <span style="font-size:14px;font-weight:700;color:{primary};">
          &#128205;&nbsp;{zone_name}
        </span>
        &nbsp;
        <span style="font-size:11px;color:#555;">
          {total_locations} location(s) with open exceptions
        </span>
      </div>
    </td></tr>

    <!-- Intro para -->
    <tr><td style="padding:8px 28px 10px;">
      {intro_block}
      <p style="margin:0;font-size:12px;color:#444;line-height:1.6;">
        Please find below a summary of <b>open SOD exceptions</b> for your zone
        as on <b>{as_of}</b>. Detailed data is attached as Excel file(s).
        Kindly take necessary action to clear the pending items at the earliest.
      </p>
    </td></tr>

    {kpi_section}

    <!-- Location-wise exception table heading -->
    <tr><td style="padding:14px 28px 4px;">
      <p style="margin:0;font-size:12px;font-weight:700;color:{primary};">
        &#127981; Location-wise Exception Detail
      </p>
    </td></tr>

    <!-- Exception summary table -->
    <tr><td style="padding:0 28px 16px;">
      <table width="100%" cellpadding="0" cellspacing="0"
             style="border-collapse:collapse;border-radius:8px;overflow:hidden;
                    border:1px solid #D5E2F3;">
        <thead>{table_header}</thead>
        <tbody>{rows_html}</tbody>
      </table>
    </td></tr>

    {attach_block}

    <!-- Footer -->
    <tr><td style="background:#F4F8FF;padding:14px 28px;
                   border-top:1px solid #D5E2F3;">
      <p style="margin:0;font-size:10px;color:#888;line-height:1.6;">
        This is an auto-generated alert from the
        <b>HPCL SOD Exception Dashboard (LIVEDB)</b>.
        For queries contact
        <a href="mailto:{SENDER_EMAIL}"
           style="color:{primary};">{SENDER_EMAIL}</a>.
      </p>
    </td></tr>

  </table>
  </td></tr>
</table>
</body></html>"""


# ── Public API ────────────────────────────────────────────────────────────────

def send_exception_mail_for_zone(
    zone_name: str,
    all_exception_plant_df: pd.DataFrame,
    detail_dfs: Dict[str, pd.DataFrame],
    selected_exceptions: List[str],
    as_of_date: str = "",
    custom_intro: str = "",
    test_mode: bool = False,
    test_email: str = "",
    zone_kpi_dict: Optional[dict] = None,
) -> dict:
    """Send exception mail for one zone with per-exception Excel attachments.

    For each selected exception type that has data for the zone, one Excel file
    is built and attached: '<ZoneName>_<ExceptionLabel>.xlsx'.
    zone_kpi_dict: per-zone KPI values rendered as tiles in the email body.

    test_mode=True  → mail goes to test_email only (no zone recipients, no BCC).
    Returns {"ok": bool, "mode": str, "msg": str, "attachments": [filenames]}
    """
    avail, reason = outlook_available()
    if not avail:
        return {"ok": False, "msg": reason}

    contacts = ZONE_EMAIL_MAP.get(zone_name)
    if not contacts:
        return {"ok": False,
                "msg": f"No email contacts configured for zone: {zone_name}"}

    # Summary rows for this zone (for mail body)
    zone_summary = all_exception_plant_df[
        all_exception_plant_df["Zone Name"] == zone_name
    ].copy()

    # Build per-exception Excel attachments
    attachments = build_zone_excel_attachments(zone_name, detail_dfs, selected_exceptions)
    if not attachments:
        return {"ok": False,
                "msg": f"No exception data found for zone '{zone_name}' "
                       f"for the selected exception types."}

    attachment_names = [fname for fname, _ in attachments]

    html_body = build_exception_email_html(
        zone_name, as_of_date, zone_summary, selected_exceptions,
        custom_intro, attachment_names,
        zone_kpi_dict=zone_kpi_dict,
    )

    as_of   = as_of_date or datetime.now().strftime("%d %b %Y")
    subject = f"[HPCL SOD Exception Alert] {zone_name} — {as_of}"
    if test_mode:
        subject = f"[TEST] {subject}"

    if test_mode:
        to_str  = test_email or SENDER_EMAIL
        cc_str  = ""
        bcc_str = ""
    else:
        to_str  = contacts["to"]
        cc_str  = contacts.get("cc", "")
        bcc_str = "; ".join(BCC_EMAILS)

    result = _send_via_outlook(
        to=to_str,
        subject=subject,
        html_body=html_body,
        cc=cc_str,
        bcc=bcc_str,
        attachments=attachments,
    )
    result["attachments"] = attachment_names
    return result
