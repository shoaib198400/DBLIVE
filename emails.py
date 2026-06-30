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
import platform
from datetime import datetime
from typing import Dict, List, Optional

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
    "Bhubneshwar Zone":     {"to": "smarak.lenka@hpcl.in",              "cc": "ECZ.OND.IC@hpcl.in"},
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

# ── Exception column display labels ──────────────────────────────────────────

EXCEPTION_LABELS: Dict[str, str] = {
    "Pending DC":                    "Pending DCs",
    "Open Delivery":                 "Open Deliveries",
    "Open In-Transit":               "Open In-Transit STOs",
    "Open Sales Order":              "Open Sales Orders",
    "Pending Invoice":               "Pending Invoices",
    "Shortage Sales (Billing Docs)": "Shortages — Sales (Billing Docs)",
    "Shortage STO (Billing Docs)":   "Shortages — STO (Billing Docs)",
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


# ── Low-level Outlook sender ──────────────────────────────────────────────────

def _send_via_outlook(
    to: str,
    subject: str,
    html_body: str,
    cc: str = "",
    bcc: str = "",
    attachments: Optional[List[tuple[str, bytes]]] = None,
) -> dict:
    """Send one email through Outlook COM.

    attachments: list of (filename, bytes) tuples
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
        mail_item.To = to
        mail_item.Subject = subject
        mail_item.HTMLBody = html_body
        if cc.strip():
            mail_item.CC = cc
        if bcc.strip():
            mail_item.BCC = bcc

        # Attach files from (name, bytes) pairs via temp files
        if attachments:
            import os, tempfile
            for fname, fbytes in attachments:
                tmp = tempfile.NamedTemporaryFile(
                    delete=False, suffix=os.path.splitext(fname)[1], prefix="hpcl_"
                )
                tmp.write(fbytes)
                tmp.close()
                tmp_paths.append(tmp.name)
                mail_item.Attachments.Add(tmp.name, 1, 1, fname)

        try:
            mail_item.Send()
            return {"ok": True, "mode": "sent"}
        except Exception as send_exc:
            try:
                mail_item.Save()
                return {"ok": True, "mode": "draft", "msg": str(send_exc)}
            except Exception as draft_exc:
                return {"ok": False, "msg": f"Send and draft both failed: {draft_exc}"}

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
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
        # Clean up temp attachment files
        import os
        for p in tmp_paths:
            try:
                os.remove(p)
            except Exception:
                pass


# ── HTML email builder ────────────────────────────────────────────────────────

def build_exception_email_html(
    zone_name: str,
    as_of_date: str,
    exception_data: pd.DataFrame,          # filtered to this zone; columns = plant + exception cols
    selected_exceptions: List[str],        # which exception columns to show
    custom_intro: str = "",
) -> str:
    """Build HPCL-branded HTML email body for one zone's exceptions."""

    primary   = "#003087"
    secondary = "#0057A8"
    accent    = "#FFD700"

    # Filter to only locations that have at least one selected exception > 0
    exc_cols = [c for c in selected_exceptions if c in exception_data.columns]
    if exc_cols:
        mask = exception_data[exc_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) > 0
        zone_df = exception_data[mask].copy()
    else:
        zone_df = exception_data.copy()

    # Build exception rows HTML
    def _fmt(v):
        try:
            n = int(float(v))
            return f'<span style="color:{"#CC0000" if n > 0 else "#008800"};font-weight:{"700" if n > 0 else "500"};">{n:,}</span>'
        except Exception:
            return str(v)

    # Table header
    col_heads = "".join(
        f'<th style="background:{secondary};color:#fff;padding:9px 12px;'
        f'font-size:12px;font-weight:700;text-align:center;'
        f'border:1px solid rgba(255,255,255,0.2);white-space:nowrap;">'
        f'{EXCEPTION_LABELS.get(c, c)}</th>'
        for c in exc_cols
    )
    table_header = (
        f'<tr>'
        f'<th style="background:{primary};color:{accent};padding:9px 14px;'
        f'font-size:12px;font-weight:700;text-align:left;'
        f'border:1px solid rgba(255,255,255,0.2);">Location</th>'
        f'{col_heads}'
        f'</tr>'
    )

    # Table rows
    rows_html = ""
    for i, (_, row) in enumerate(zone_df.iterrows()):
        bg = "#FFFFFF" if i % 2 == 0 else "#F4F8FF"
        cells = "".join(
            f'<td style="padding:8px 12px;font-size:12px;text-align:center;'
            f'border:1px solid #E2EAF4;background:{bg};">{_fmt(row.get(c, 0))}</td>'
            for c in exc_cols
        )
        rows_html += (
            f'<tr>'
            f'<td style="padding:8px 14px;font-size:12px;font-weight:600;'
            f'color:#003087;border:1px solid #E2EAF4;background:{bg};">'
            f'{row.get("Plant Name", row.get("Location", ""))}</td>'
            f'{cells}'
            f'</tr>'
        )

    if not rows_html:
        rows_html = (
            f'<tr><td colspan="{len(exc_cols)+1}" style="text-align:center;'
            f'padding:20px;color:#666;">No exceptions found for this zone.</td></tr>'
        )

    intro_block = (
        f'<p style="margin:0 0 18px;font-size:14px;line-height:1.6;color:#333;">{custom_intro}</p>'
        if custom_intro else ""
    )

    total_locations = len(zone_df)
    as_of = as_of_date or datetime.now().strftime("%d %b %Y")

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#F0F2F6;font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#F0F2F6;padding:20px 0;">
  <tr><td align="center">
  <table width="680" cellpadding="0" cellspacing="0"
         style="background:#fff;border-radius:12px;overflow:hidden;
                box-shadow:0 4px 20px rgba(0,48,135,0.12);">

    <!-- Header -->
    <tr><td style="background:linear-gradient(135deg,{primary} 0%,{secondary} 100%);
                   padding:22px 30px;">
      <table width="100%" cellpadding="0" cellspacing="0"><tr>
        <td>
          <div style="font-size:22px;font-weight:900;color:#fff;letter-spacing:0.04em;">
            &#9981; HPCL — SOD Exception Alert
          </div>
          <div style="font-size:13px;color:{accent};margin-top:4px;font-weight:600;">
            Hindustan Petroleum Corporation Limited
          </div>
        </td>
        <td align="right" style="font-size:12px;color:rgba(255,255,255,0.75);">
          {as_of}
        </td>
      </tr></table>
    </td></tr>

    <!-- Zone pill -->
    <tr><td style="padding:18px 30px 10px;">
      <div style="display:inline-block;background:#E8F0FE;border-left:4px solid {primary};
                  border-radius:6px;padding:10px 18px;">
        <span style="font-size:15px;font-weight:700;color:{primary};">
          &#128205; {zone_name}
        </span>
        &nbsp;
        <span style="font-size:12px;color:#555;">{total_locations} location(s) with open exceptions</span>
      </div>
    </td></tr>

    <!-- Intro -->
    <tr><td style="padding:8px 30px 14px;">
      {intro_block}
      <p style="margin:0;font-size:13px;color:#444;line-height:1.6;">
        Please find below the summary of <b>open SOD exceptions</b> for your zone as on
        <b>{as_of}</b>. Kindly take necessary action at the earliest to clear the pending items.
      </p>
    </td></tr>

    <!-- Exception table -->
    <tr><td style="padding:0 30px 22px;">
      <table width="100%" cellpadding="0" cellspacing="0"
             style="border-collapse:collapse;border-radius:8px;overflow:hidden;
                    border:1px solid #D5E2F3;">
        <thead>{table_header}</thead>
        <tbody>{rows_html}</tbody>
      </table>
    </td></tr>

    <!-- Footer -->
    <tr><td style="background:#F4F8FF;padding:16px 30px;border-top:1px solid #D5E2F3;">
      <p style="margin:0;font-size:11px;color:#888;line-height:1.6;">
        This is an auto-generated alert from the <b>HPCL SOD Exception Dashboard (LIVEDB)</b>.
        For queries, contact
        <a href="mailto:{SENDER_EMAIL}" style="color:{primary};">{SENDER_EMAIL}</a>.
        <br>A detailed Excel report is attached for your reference.
      </p>
    </td></tr>

  </table>
  </td></tr>
</table>
</body></html>"""


# ── Excel attachment builder ──────────────────────────────────────────────────

def build_zone_excel_attachment(
    zone_name: str,
    all_exception_plant_df: pd.DataFrame,
    selected_exceptions: List[str],
    as_of_date: str = "",
) -> bytes:
    """Build an Excel workbook for one zone's exception detail. Returns bytes."""
    exc_cols = [c for c in selected_exceptions if c in all_exception_plant_df.columns]
    mask = all_exception_plant_df["Zone Name"] == zone_name
    zone_df = all_exception_plant_df[mask][["Zone Name", "Plant Name"] + exc_cols + ["Total Exceptions"]].copy()
    zone_df = zone_df[
        zone_df[exc_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) > 0
    ]
    zone_df = zone_df.sort_values("Total Exceptions", ascending=False)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        zone_df.to_excel(writer, index=False, sheet_name="Exceptions")
        ws = writer.sheets["Exceptions"]

        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        hdr_fill = PatternFill("solid", fgColor="003087")
        hdr_font = Font(bold=True, color="FFFFFF", size=11)
        thin = Side(style="thin", color="D5E2F3")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 30)

    return buf.getvalue()


# ── Public API ────────────────────────────────────────────────────────────────

def send_exception_mail_for_zone(
    zone_name: str,
    all_exception_plant_df: pd.DataFrame,
    selected_exceptions: List[str],
    as_of_date: str = "",
    custom_intro: str = "",
    test_mode: bool = False,
    test_email: str = "",
) -> dict:
    """Send (or draft) exception mail for one zone.

    test_mode=True → send to test_email instead of the zone's real recipients.
    Returns {"ok": bool, "mode": str, "msg": str}
    """
    avail, reason = outlook_available()
    if not avail:
        return {"ok": False, "msg": reason}

    contacts = ZONE_EMAIL_MAP.get(zone_name)
    if not contacts:
        return {"ok": False, "msg": f"No email contacts configured for zone: {zone_name}"}

    zone_df = all_exception_plant_df[all_exception_plant_df["Zone Name"] == zone_name].copy()
    if zone_df.empty:
        return {"ok": False, "msg": f"No exception data found for zone: {zone_name}"}

    html_body = build_exception_email_html(
        zone_name, as_of_date, zone_df, selected_exceptions, custom_intro
    )
    excel_bytes = build_zone_excel_attachment(
        zone_name, all_exception_plant_df, selected_exceptions, as_of_date
    )
    safe_zone = zone_name.replace(" ", "_").replace("(", "").replace(")", "")
    as_of_safe = (as_of_date or datetime.now().strftime("%d%b%Y")).replace(" ", "")
    attachment_name = f"SOD_Exceptions_{safe_zone}_{as_of_safe}.xlsx"

    as_of = as_of_date or datetime.now().strftime("%d %b %Y")
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

    return _send_via_outlook(
        to=to_str,
        subject=subject,
        html_body=html_body,
        cc=cc_str,
        bcc=bcc_str,
        attachments=[(attachment_name, excel_bytes)],
    )
