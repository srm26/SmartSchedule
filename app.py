# -*- coding: utf-8 -*-
"""
GES SmartSchedule - AI-Powered Electrical Labor Calendar
"""
import os
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import re
import io
import json
import hashlib
from pathlib import Path
from datetime import datetime as dt
from zoneinfo import ZoneInfo

_PT = ZoneInfo("America/Los_Angeles")
import msal
import requests
from html import escape
from openpyxl import load_workbook
from apscheduler.schedulers.background import BackgroundScheduler
from schedule_config import FACILITIES, OT_RULES, HOLIDAYS, DT_RULES, OUTLOOK_CATEGORIES
from schedule_logic import facility_route, has_facility, compute_overtime, ot_summary

# ─── PAGE CONFIG ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="GES SmartSchedule",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  [data-testid="stSidebar"] { background: #1a1f36; }
  [data-testid="stSidebar"] * { color: #e2e8f0 !important; }
  [data-testid="stSidebar"] input { color: #1e293b !important; }
  [data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {
    background: #2d3452; border: 1px solid #64748b; }
  [data-testid="stSidebar"] [data-testid="stBaseButton-secondary"]:hover {
    background: #3b4470; border-color: #e2e8f0; }
  [data-testid="stSidebar"] .stRadio label { font-size: 15px; padding: 6px 0; }
  /* Status boxes in the sidebar ("No file loaded", "Last check", file name): white box, dark text */
  [data-testid="stSidebar"] [data-testid="stAlertContainer"] {
    background: #fff !important; border: 1px solid #cbd5e1; border-left: 4px solid #f59e0b; }
  [data-testid="stSidebar"] [data-testid="stAlertContainer"] * { color: #1e293b !important; }
  .metric-card {
    background:#fff; border:1px solid #e2e8f0; border-radius:12px;
    padding:20px 24px; box-shadow:0 1px 4px rgba(0,0,0,.06);
  }
  .metric-label { font-size:12px; color:#64748b; font-weight:700;
    letter-spacing:.5px; text-transform:uppercase; margin-bottom:4px; }
  .metric-value { font-size:30px; font-weight:800; color:#1e293b; }
  .metric-delta { font-size:12px; color:#22c55e; margin-top:3px; }
  .section-title {
    font-size:19px; font-weight:700; color:#1e293b;
    border-left:4px solid #f59e0b; padding-left:10px; margin:20px 0 12px;
  }
  .fn-badge {
    display:inline-block; padding:2px 10px; border-radius:20px;
    font-size:11px; font-weight:700;
  }
  .fn-set     { background:#dbeafe; color:#1d4ed8; }
  .fn-strike  { background:#fee2e2; color:#b91c1c; }
  .fn-standby { background:#fef9c3; color:#854d0e; }
  .notif-card {
    background:#f8fafc; border-left:4px solid #f59e0b;
    border-radius:6px; padding:12px 16px; margin-bottom:8px;
  }
  .ai-banner {
    background:linear-gradient(135deg,#1e293b,#334155);
    border-radius:12px; padding:18px 22px; color:white; margin:14px 0;
  }
  div[data-testid="stTabs"] button { font-size:15px; font-weight:600; padding:10px 20px; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ───────────────────────────────────────────────────────
FUNCTION_COLORS = {"SET": "#3b82f6", "STRIKE": "#ef4444", "STANDBY": "#f59e0b"}
LV_SHEET        = "By Day"
LV_CREW_COLS    = [1, 2, 3, 4, 5, 6, 7]   # numeric header columns for crew slots
LISTS_SHEET     = "Lists"                  # holds the BRANCH / STATE / TIME ZONE lookup
DEFAULT_TZ      = "Pacific Standard Time"  # used when a row's branch has no lookup entry (e.g. legacy LV IDs)
LINK_COL        = "SHAREPOINT LINK"        # optional By Day column; shown on the Outlook event notes
FACILITY_COL    = "FACILITY"               # By Day column; facility code that picks the Outlook calendar

# ─── OUTLOOK BODY MARKERS ────────────────────────────────────────────
# EVENT_ID and field hash are stored in the first line of the event body.
# Outlook is the source of truth — no local file needed.

def _field_hash(row):
    key = "|".join(str(row.get(f, "")) for f in ["DATE", "SHOW", "FUNCTION", "TIME", "LOCATION", "FOREMAN", "NOTES"])
    if row.get("LINK"):  # only when present, so rows without a link keep their existing hash
        key += "|" + str(row["LINK"])
    if _event_categories(row):  # same idea: only colored functions get a new hash (one-time update)
        key += "|cat:" + ",".join(_event_categories(row))
    return hashlib.md5(key.encode()).hexdigest()[:8]

def _event_categories(row):
    """Outlook categories (event color) for the row's FUNCTION, from OUTLOOK_CATEGORIES."""
    cat = OUTLOOK_CATEGORIES.get(str(row.get("FUNCTION", "")).upper())
    return [cat["name"]] if cat else []

_PRESET_NAMES = {"preset0": "red", "preset1": "orange", "preset2": "brown", "preset3": "yellow", "preset4": "green",
                 "preset5": "teal", "preset6": "olive", "preset7": "blue", "preset8": "purple",
                 "preset9": "cranberry", "preset12": "gray"}

def ensure_outlook_categories(token, mailbox):
    """Create missing OUTLOOK_CATEGORIES (name + color) in the mailbox's category list.
    Returns '' or an error message (e.g. missing MailboxSettings.ReadWrite permission)."""
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/outlook/masterCategories"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, timeout=15)
    if resp.status_code != 200:
        return (f"could not read Outlook categories in {mailbox} ({resp.status_code}). Events are still tagged; "
                f"to show colors, grant the app MailboxSettings.ReadWrite or create the categories by hand: "
                + ", ".join(f"{c['name']} ({_PRESET_NAMES.get(c['color'], c['color'])})"
                            for c in OUTLOOK_CATEGORIES.values()))
    have = {c["displayName"].lower() for c in resp.json().get("value", [])}
    for c in OUTLOOK_CATEGORIES.values():
        if c["name"].lower() not in have:
            r = requests.post(url, headers=headers, json={"displayName": c["name"], "color": c["color"]}, timeout=15)
            if r.status_code not in (200, 201):
                return f"could not create Outlook category {c['name']} in {mailbox} ({r.status_code})"
    return ""

def _event_notes(row):
    """Outlook notes text: foreman, the NOTES column, and the SharePoint link on its own line."""
    notes = f"Foreman: {row['FOREMAN']}. {row['NOTES']}".strip()
    if row.get("LINK"):
        label = row.get("LINK_TEXT") or row["LINK"]
        notes += f"\nSharePoint: [{label}]({row['LINK']})"  # turned into a clickable link by _make_body
    return notes

def _make_body(notes, event_id, field_hash):
    """HTML event body: marker line, then the notes; '[label](url)' becomes a clickable link."""
    marker = f"GES_ID:{event_id}|GES_HASH:{field_hash}"
    html = escape(f"{marker}\n{notes}".strip() if notes else marker)
    html = re.sub(r"\[([^\]\n]+)\]\((\S+?)\)", r'<a href="\2">\1</a>', html)
    return html.replace("\n", "<br>")

def _parse_body(body_content):
    """Extract (event_id, field_hash) from event body, or ('', '') if not present."""
    for line in (body_content or "").splitlines():
        line = line.strip()
        if line.startswith("GES_ID:") and "|GES_HASH:" in line:
            parts = line.split("|GES_HASH:")
            return parts[0].replace("GES_ID:", "").strip(), parts[1].strip() if len(parts) > 1 else ""
    return "", ""

def build_calendar_map(token, mailbox, start_date, end_date):
    """Query Outlook for the date range and return {EVENT_ID: {graph_id, hash}}."""
    url    = f"https://graph.microsoft.com/v1.0/users/{mailbox}/calendarView"
    params = {
        "startDateTime": start_date.strftime("%Y-%m-%dT00:00:00Z"),
        "endDateTime":   (end_date + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00Z"),
        "$select": "id,body",
        "$top": "500",
    }
    headers = {"Authorization": f"Bearer {token}", "Prefer": "outlook.body-content-type='text'"}
    mapping, next_url, next_params = {}, url, params
    while next_url:
        resp = requests.get(next_url, headers=headers, params=next_params, timeout=15)
        if resp.status_code != 200:
            break
        data = resp.json()
        for e in data.get("value", []):
            body_text = e.get("body", {}).get("content", "")
            eid, fhash = _parse_body(body_text)
            if eid:
                mapping[eid] = {"graph_id": e["id"], "hash": fhash}
        next_url, next_params = data.get("@odata.nextLink"), None
    return mapping

# ─── GRAPH HELPERS ───────────────────────────────────────────────────
def _graph_cfg():
    """Read Graph credentials from env vars (App Service) or secrets.toml (local)."""
    env = {
        "app_id":         os.environ.get("GRAPH_APP_ID"),
        "tenant_id":      os.environ.get("GRAPH_TENANT_ID"),
        "client_secret":  os.environ.get("GRAPH_CLIENT_SECRET"),
        "shared_mailbox": os.environ.get("GRAPH_SHARED_MAILBOX"),
    }
    if all(env.values()):
        return env
    try:
        return dict(st.secrets["graph"])
    except Exception:
        return env

def _mailbox_setting(env_name, secrets_key):
    """{KEY: mailbox} from a 'KEY=a@x.com; KEY2=b@x.com' app setting, or the secrets.toml key locally."""
    raw = os.environ.get(env_name)
    if raw is None:
        raw = _graph_cfg().get(secrets_key, "")
    mapping = {}
    for pair in str(raw).split(";"):
        if "=" in pair:
            k, mbx = pair.split("=", 1)
            if mbx.strip():
                mapping[k.strip().upper()] = mbx.strip()
    return mapping

def _mailbox_for(branch):
    """Shared mailbox for a branch, from GRAPH_BRANCH_MAILBOXES ('LV=a@x.com; ORL=b@x.com',
    app setting or secrets.toml branch_mailboxes). Unlisted branches use the default shared mailbox."""
    mapping = _mailbox_setting("GRAPH_BRANCH_MAILBOXES", "branch_mailboxes")
    return mapping.get(str(branch or "").upper()) or _graph_cfg()["shared_mailbox"]

def _route_rows(df):
    """(mailbox, error) per row. Rows with a FACILITY go to that facility's calendar (schedule_config.py +
    GRAPH_CALENDAR_MAILBOXES); an unknown or unmapped facility gets an error and is not pushed.
    A blank FACILITY (cities without facilities) goes to the city's calendar from the ID prefix, only if
    that branch is listed in GRAPH_BRANCH_MAILBOXES; otherwise it's an error too.
    Files without a FACILITY column keep the older routing by ID prefix (_mailbox_for)."""
    cal_mbx    = _mailbox_setting("GRAPH_CALENDAR_MAILBOXES", "calendar_mailboxes")
    branch_mbx = _mailbox_setting("GRAPH_BRANCH_MAILBOXES", "branch_mailboxes")

    def route(r):
        if not has_facility(r["FACILITY"]):
            return _mailbox_for(r["BRANCH"]), ""
        if r["FACILITY"]:
            return facility_route(r["FACILITY"], cal_mbx, r["BRANCH"])
        branch = str(r["BRANCH"] or "").upper()
        if branch in branch_mbx:
            return branch_mbx[branch], ""
        return None, ("FACILITY is blank and " + (f"branch {branch} has no city calendar" if branch
                      else "the ID has no branch prefix") + " in GRAPH_BRANCH_MAILBOXES")

    routes = [route(r) for _, r in df.iterrows()]
    return (pd.Series([m for m, _ in routes], index=df.index, dtype=object),
            pd.Series([e for _, e in routes], index=df.index, dtype=object))

def _graph_secrets_ok():
    try:
        return bool(_graph_cfg()["client_secret"])
    except Exception:
        return False

def get_graph_token():
    g = _graph_cfg()
    app = msal.ConfidentialClientApplication(
        client_id=g["app_id"],
        client_credential=g["client_secret"],
        authority=f"https://login.microsoftonline.com/{g['tenant_id']}",
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    raise RuntimeError(result.get("error_description", "Graph auth failed"))

def _make_event_id(show, function, date):
    slug = re.sub(r"[^A-Za-z0-9]", "", str(show))[:20].upper()
    return f"{slug}-{function}-{date.strftime('%Y%m%d')}"

def patch_graph_event(token, mailbox, graph_id, start_dt, end_dt,
                      location="", notes="", time_str="", event_id="", field_hash="", subject="",
                      time_zone=DEFAULT_TZ, categories=None):
    url     = f"https://graph.microsoft.com/v1.0/users/{mailbox}/events/{graph_id}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    t_start, t_end = _parse_time_range(time_str)
    body_text = _make_body(notes, event_id, field_hash)
    if t_start and t_end:
        payload = {
            "subject": subject, "isAllDay": False,
            "start": {"dateTime": start_dt.strftime(f"%Y-%m-%dT{t_start}:00"), "timeZone": time_zone},
            "end":   {"dateTime": _timed_end(end_dt, t_start, t_end),         "timeZone": time_zone},
            "location": {"displayName": location or "Not specified"},
            "body": {"contentType": "HTML", "content": body_text},
        }
    else:
        payload = {
            "subject": subject, "isAllDay": True,
            "start": {"dateTime": start_dt.strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"},
            "end":   {"dateTime": (end_dt + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"},
            "location": {"displayName": location or "Not specified"},
            "body": {"contentType": "HTML", "content": body_text},
        }
    payload["categories"] = categories or []
    return requests.patch(url, headers=headers, json=payload, timeout=15)

def delete_graph_event(token, mailbox, graph_id):
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/events/{graph_id}"
    return requests.delete(url, headers={"Authorization": f"Bearer {token}"}, timeout=10)

def _parse_time_range(time_str):
    """Parse '8:00 - 16:30' into ('08:00', '16:30'), or return (None, None)."""
    if not time_str:
        return None, None
    m = re.match(r"(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})", str(time_str).strip())
    if not m:
        return None, None
    def pad(t):
        h, mi = t.split(":")
        return f"{int(h):02d}:{mi}"
    return pad(m.group(1)), pad(m.group(2))

def _timed_end(end_dt, t_start, t_end):
    """End dateTime string for a timed event. Overnight shifts ('22:00 - 02:00', '20:00 - 24:00')
    end on the next day; otherwise Outlook rejects an end that is before the start."""
    if t_end == "24:00":
        t_end = "00:00"
    if t_end <= t_start:
        end_dt = end_dt + timedelta(days=1)
    return end_dt.strftime(f"%Y-%m-%dT{t_end}:00")

def create_graph_event(token, mailbox, subject, start_dt, end_dt,
                       location="", notes="", time_str="", event_id="", field_hash="", time_zone=DEFAULT_TZ,
                       categories=None):
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/events"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    t_start, t_end = _parse_time_range(time_str)
    body_text = _make_body(notes, event_id, field_hash)
    if t_start and t_end:
        payload = {
            "subject": subject, "isAllDay": False,
            "start": {"dateTime": start_dt.strftime(f"%Y-%m-%dT{t_start}:00"), "timeZone": time_zone},
            "end":   {"dateTime": _timed_end(end_dt, t_start, t_end),       "timeZone": time_zone},
            "location": {"displayName": location or "Not specified"},
            "body": {"contentType": "HTML", "content": body_text},
        }
    else:
        payload = {
            "subject": subject, "isAllDay": True,
            "start": {"dateTime": start_dt.strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"},
            "end":   {"dateTime": (end_dt + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"},
            "location": {"displayName": location or "Not specified"},
            "body": {"contentType": "HTML", "content": body_text},
        }
    payload["categories"] = categories or []
    return requests.post(url, headers=headers, json=payload, timeout=15)

# ─── SHAREPOINT HELPERS ──────────────────────────────────────────────
SP_HOSTNAME = "viadcorp.sharepoint.com"
SP_SITE     = "/sites/ErinsTestSandbox"
SP_FOLDER   = "Attachments/ElectricalSchedule"

def _sp_headers(token):
    return {"Authorization": f"Bearer {token}"}

def get_sharepoint_files(token):
    site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{SP_HOSTNAME}:{SP_SITE}",
        headers=_sp_headers(token), timeout=10,
    )
    site.raise_for_status()
    site_id = site.json()["id"]

    drive = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive",
        headers=_sp_headers(token), timeout=10,
    )
    drive.raise_for_status()
    drive_id = drive.json()["id"]

    children = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SP_FOLDER}:/children",
        headers=_sp_headers(token),
        params={"$select": "id,name,lastModifiedDateTime,file", "$top": "100"},
        timeout=10,
    )
    children.raise_for_status()
    return drive_id, [
        {"name": i["name"], "id": i["id"],
         "modified":         i.get("lastModifiedDateTime", ""),        # full ISO for change detection
         "modified_display": dt.fromisoformat(i["lastModifiedDateTime"].replace("Z", "+00:00"))
                               .astimezone(_PT).strftime("%Y-%m-%d %H:%M PT")}  # for UI
        for i in children.json().get("value", [])
        if i.get("name", "").endswith(".xlsx") and "file" in i
    ]

def download_sharepoint_file(token, drive_id, item_id):
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
        headers=_sp_headers(token), timeout=30, allow_redirects=True,
    )
    resp.raise_for_status()
    return resp.content

# ─── LV FORMAT IMPORTER ──────────────────────────────────────────────
def _read_branch_time_zones(xl):
    """Return {BRANCH: Outlook time zone} from the 'Lists' sheet, or {} if the sheet/table is absent."""
    if LISTS_SHEET not in xl.sheet_names:
        return {}
    lists = xl.parse(LISTS_SHEET, header=0, dtype=str)
    lists.columns = [str(c).strip().upper() for c in lists.columns]
    if "BRANCH" not in lists.columns or "TIME ZONE" not in lists.columns:
        return {}
    pairs = lists[["BRANCH", "TIME ZONE"]].dropna()
    return {b.strip().upper(): tz.strip() for b, tz in pairs.itertuples(index=False) if b.strip() and tz.strip()}

def _cell_text(val):
    """Cell value as stripped text; empty cells (NaN/None) become '' instead of 'nan'."""
    s = str(val if val is not None else "").strip()
    return "" if s.lower() in ("nan", "none") else s

def _read_link_urls(src, sheet_name):
    """{ID: URL} for SHAREPOINT LINK cells that are hyperlinks. pandas only sees the text Excel
    displays (often a file name), so the URL behind the cell is read with openpyxl."""
    if hasattr(src, "seek"):
        src.seek(0)
    ws = load_workbook(src, data_only=True)[sheet_name]
    header = [str(c.value).strip().upper() if c.value is not None else "" for c in ws[1]]
    if LINK_COL not in header or "ID" not in header:
        return {}
    li, ii = header.index(LINK_COL), header.index("ID")
    urls = {}
    for row in ws.iter_rows(min_row=2):
        cell = row[li]
        if cell.hyperlink is not None and cell.hyperlink.target:
            id_val = row[ii].value
            if isinstance(id_val, float) and id_val.is_integer():
                id_val = int(id_val)  # match parse_lv_byDay's ID handling
            urls[_cell_text(id_val)] = cell.hyperlink.target
    return urls

def parse_lv_byDay(uploaded_file, sheet_name):
    """Parse the 'By Day' sheet from the LV GES Electrical Show Schedule format.

    Expected columns: DAY, DATE, TIME, SHOW, FUNCTION, LOCATION, NOTES, FOREMAN, 1..7
    Crew columns 1-7 hold names or '+N' (extra unnamed crew).
    Each row's time zone comes from its ID prefix (e.g. 'ORL-0001' -> ORL) looked up in
    the BRANCH / TIME ZONE table on the 'Lists' sheet; no prefix or no table -> DEFAULT_TZ.
    Optional FACILITY column (facility code, picks the Outlook calendar per row): stored upper-cased;
    None on every row when the file has no FACILITY column.
    """
    xl = pd.ExcelFile(uploaded_file)
    # Read crew columns as text, otherwise a column holding only '+1' comes back as 1.0
    raw = xl.parse(sheet_name, header=0, dtype={c: str for c in LV_CREW_COLS})
    branch_tz = _read_branch_time_zones(xl)
    link_urls = _read_link_urls(uploaded_file, sheet_name) if LINK_COL in raw.columns else {}

    required = {"DATE", "SHOW", "FUNCTION", "LOCATION", "NOTES", "FOREMAN"}
    present  = {str(c).strip().upper() for c in raw.columns}
    missing  = required - present
    if missing:
        raise ValueError(f"Sheet '{sheet_name}' is missing expected columns: {', '.join(sorted(missing))}")

    crew_cols = [c for c in raw.columns if isinstance(c, int)]
    has_fac   = FACILITY_COL in raw.columns

    rows = []
    for _, r in raw.iterrows():
        date = pd.to_datetime(r.get("DATE"), errors="coerce")
        if pd.isna(date) or date.year < 2020:
            continue
        show = str(r.get("SHOW", "") or "").strip()
        if not show or show.lower() in ("nan", "none"):
            continue

        foreman_val = str(r.get("FOREMAN", "") or "").strip()
        foreman     = foreman_val if foreman_val and foreman_val.lower() not in ("nan", "none", "") else "TBD"

        crew_list = [foreman] if foreman != "TBD" else []
        extra = 0
        for cc in crew_cols:
            val = str(r.get(cc, "") or "").strip()
            if not val or val.lower() in ("nan", "none", ""):
                continue
            if val.startswith("+"):
                try:
                    extra += int(val[1:])
                except ValueError:
                    extra += 1
            else:
                crew_list.append(val)

        fn = str(r.get("FUNCTION", "") or "").strip().upper() or "SHOW"
        id_val = r.get("ID", "")
        if isinstance(id_val, float) and id_val.is_integer():
            id_val = int(id_val)  # Excel numbers come in as 1.0 — store as "1"
        raw_id = str(id_val or "").strip()
        event_id = raw_id if raw_id and raw_id.lower() not in ("nan", "none") else _make_event_id(show, fn, date)
        m = re.match(r"([A-Za-z]+)-\d", raw_id)
        branch = m.group(1).upper() if m else ""
        if branch and branch_tz and branch not in branch_tz:
            raise ValueError(f"ID '{raw_id}': branch '{branch}' is not in the BRANCH / TIME ZONE table "
                             f"on the '{LISTS_SHEET}' sheet.")
        rows.append({
            "EVENT_ID":      event_id,
            "DATE":          date,
            "END_DATE":      pd.NaT,
            "SHOW":          show,
            "FUNCTION":      fn,
            "LOCATION":      str(r.get("LOCATION", "") or "").strip(),
            "TIME":          str(r.get("TIME", "") or "").strip(),
            "FOREMAN":       foreman,
            "NOTES":         _cell_text(r.get("NOTES")),
            "TOTAL_CREW":    max(len(crew_list) + extra, 1),
            "CREW_LIST":     crew_list,
            "SCHEDULE_TYPE": "shifts",
            "TIME_ZONE":     branch_tz.get(branch, DEFAULT_TZ),
            # LINK = the URL (from the hyperlink behind the cell if there is one); LINK_TEXT = what the cell shows
            "LINK":          link_urls.get(raw_id) or _cell_text(r.get(LINK_COL)),
            "LINK_TEXT":     _cell_text(r.get(LINK_COL)),
            "BRANCH":        branch,
            "FACILITY":      _cell_text(r.get(FACILITY_COL)).upper() if has_fac else None,
            "EXTRA_CREW":    extra,  # unnamed '+N' crew, used by the Overtime page
        })

    if not rows:
        raise ValueError("No valid shifts found in the selected sheet.")
    return pd.DataFrame(rows).reset_index(drop=True)

# ─── DATA HELPERS ────────────────────────────────────────────────────
def build_notifications(df, demo_contacts):
    rows = []
    union_email   = demo_contacts.get("union_email",   "union-dispatch@ibew.org")
    foreman_email = demo_contacts.get("foreman_email", "foreman@ges.com")
    for _, row in df.iterrows():
        date_str    = row["DATE"].strftime("%Y-%m-%d")
        notify_date = (row["DATE"] - timedelta(days=2)).strftime("%Y-%m-%d")
        show = row["SHOW"]; fn = row["FUNCTION"]
        loc  = row["LOCATION"]; time_s = row["TIME"]; foreman = row["FOREMAN"]
        rows.append({
            "Type": "Labor Call", "Recipient": f"Union Dispatch ({union_email})",
            "Show": show, "Date": date_str, "Notify By": notify_date,
            "Message": (f"Labor call for {int(row['TOTAL_CREW'])} electrician(s). "
                        f"Show: {show} | {fn} | {date_str} {time_s} | {loc}. Foreman: {foreman}."),
            "Channel": "Email", "Status": "Ready",
        })
        if foreman not in ("TBD", "nan", ""):
            rows.append({
                "Type": "Foreman Assignment", "Recipient": f"{foreman} ({foreman_email})",
                "Show": show, "Date": date_str, "Notify By": notify_date,
                "Message": f"You are Foreman for {show} | {fn} | {date_str} {time_s} | {loc}.",
                "Channel": "Email", "Status": "Ready",
            })
    return pd.DataFrame(rows)

def _build_ics(df_shifts, cal_name):
    if df_shifts.empty:
        return b""
    lines = [
        "BEGIN:VCALENDAR", "VERSION:2.0",
        f"PRODID:-//GES SmartSchedule//{cal_name}//EN",
        f"X-WR-CALNAME:GES {cal_name}",
    ]
    for _, r in df_shifts.iterrows():
        start_str = r["DATE"].strftime("%Y%m%d")
        end_val   = r["END_DATE"] if pd.notna(r["END_DATE"]) else r["DATE"]
        end_str   = (pd.Timestamp(end_val) + timedelta(days=1)).strftime("%Y%m%d")
        lines += [
            "BEGIN:VEVENT",
            f"DTSTART;VALUE=DATE:{start_str}",
            f"DTEND;VALUE=DATE:{end_str}",
            f"SUMMARY:[GES] {r['SHOW']} - {r['FUNCTION']}",
            f"DESCRIPTION:{r['TIME']} | {r['LOCATION']} | Foreman: {r['FOREMAN']}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines).encode()

# ─── AUTO-SWEEP ENGINE ───────────────────────────────────────────────
# On App Service the app folder is a read-only package; /home is writable and survives restarts.
SWEEP_STATE_FILE = (Path("/home/data/sweep_state.json") if os.environ.get("WEBSITE_SITE_NAME")
                    else Path(".streamlit/sweep_state.json"))

@st.cache_resource
def _get_sweep_status():
    return {"last_check": None, "running": False, "log": [], "error": None}

@st.cache_resource
def _get_sweep_config():
    return {"test_rows": 2}

def _load_sweep_state():
    try:
        return json.loads(SWEEP_STATE_FILE.read_text()) if SWEEP_STATE_FILE.exists() else {}
    except Exception:
        return {}

def _save_sweep_state(s):
    SWEEP_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    SWEEP_STATE_FILE.write_text(json.dumps(s, indent=2))

def _load_folder(token):
    """Download and parse every LV-format file in the SharePoint folder.
    Files without a 'By Day' sheet are not schedules and are ignored.
    Returns (sp_files, {name: df}, {name: error})."""
    drive_id, sp_files = get_sharepoint_files(token)
    parsed, errors = {}, {}
    for f in sp_files:
        try:
            raw = download_sharepoint_file(token, drive_id, f["id"])
            if LV_SHEET not in pd.ExcelFile(io.BytesIO(raw)).sheet_names:
                continue
            parsed[f["name"]] = parse_lv_byDay(io.BytesIO(raw), LV_SHEET)
        except Exception as e:
            errors[f["name"]] = str(e)
    return sp_files, parsed, errors

def _sync_dataframe(token, mailbox, df_file, keep_ids=None):
    """Run create/update sync for one parsed dataframe. Returns stats dict.
    keep_ids: {EVENT_ID: mailbox the row now routes to (None if it can't be routed)}.
    Deletes Outlook events in the date range whose ID is not in keep_ids, or whose row now routes to a
    different mailbox (facility changed); pass keep_ids=None to skip deletes (test mode)."""
    cal_map = build_calendar_map(token, mailbox, df_file["DATE"].min(), df_file["DATE"].max())
    created = updated = deleted = skipped = failed = 0
    for _, row in df_file.iterrows():
        eid     = row["EVENT_ID"]
        subject = f"{row['SHOW']} - {row['FUNCTION']}"
        end_dt  = row["END_DATE"] if pd.notna(row["END_DATE"]) else row["DATE"]
        notes   = _event_notes(row)
        fhash   = _field_hash(row)
        try:
            entry = cal_map.get(eid)
            if entry and entry["hash"] == fhash:
                skipped += 1
            elif entry:
                resp = patch_graph_event(token, mailbox, entry["graph_id"],
                                         row["DATE"], pd.Timestamp(end_dt),
                                         row["LOCATION"], notes,
                                         time_str=row["TIME"], event_id=eid, field_hash=fhash,
                                         time_zone=row.get("TIME_ZONE", DEFAULT_TZ), categories=_event_categories(row),
                                         subject=subject)
                if resp.status_code == 200:
                    updated += 1
                elif resp.status_code == 404:
                    resp2 = create_graph_event(token, mailbox, subject,
                                               row["DATE"], pd.Timestamp(end_dt),
                                               row["LOCATION"], notes,
                                               time_str=row["TIME"], event_id=eid, field_hash=fhash,
                                               time_zone=row.get("TIME_ZONE", DEFAULT_TZ), categories=_event_categories(row))
                    created += 1 if resp2.status_code == 201 else 0
                    failed  += 0 if resp2.status_code == 201 else 1
                else:
                    failed += 1
            else:
                resp = create_graph_event(token, mailbox, subject,
                                          row["DATE"], pd.Timestamp(end_dt),
                                          row["LOCATION"], notes,
                                          time_str=row["TIME"], event_id=eid, field_hash=fhash,
                                          time_zone=row.get("TIME_ZONE", DEFAULT_TZ), categories=_event_categories(row))
                created += 1 if resp.status_code == 201 else 0
                failed  += 0 if resp.status_code == 201 else 1
        except Exception:
            failed += 1

    if keep_ids is not None:
        for eid, entry in cal_map.items():
            if eid not in keep_ids or keep_ids[eid] not in (None, mailbox):
                dresp = delete_graph_event(token, mailbox, entry["graph_id"])
                if dresp.status_code in (200, 204):
                    deleted += 1

    return {"created": created, "updated": updated, "deleted": deleted,
            "skipped": skipped, "failed": failed}

def run_sweep():
    status = _get_sweep_status()
    config = _get_sweep_config()
    status["running"] = True
    status["error"]   = None
    log = []
    try:
        token    = get_graph_token()
        sp_files, parsed, errors = _load_folder(token)
        state    = _load_sweep_state()
        limit    = config.get("test_rows", 0)

        # IDs from every file in the folder, so one file never deletes another's events.
        # If any file failed to parse, skip deletes rather than risk removing its events.
        keep_ids = None
        if limit == 0 and not errors:
            keep_ids = {}
            for d in parsed.values():
                keep_ids.update(zip(d["EVENT_ID"], _route_rows(d)[0]))
        elif limit == 0:
            log.append("⚠️ Deletes skipped — could not read: " + ", ".join(errors))

        cat_checked = set()  # mailboxes whose Outlook categories (event colors) were checked this sweep
        for f in sp_files:
            fname    = f["name"]
            last_mod = f["modified"]
            if state.get(fname, {}).get("last_modified") == last_mod:
                continue  # unchanged — skip
            if fname in errors:
                log.append(f"❌ {fname} — Error: {errors[fname]}")
                continue
            if fname not in parsed:
                continue  # not an LV schedule file
            try:
                df_f   = parsed[fname]
                if limit > 0:
                    df_f = df_f.head(limit)
                # Each row goes to its facility's calendar; rows that can't be routed are not pushed
                mbx_s, err_s = _route_rows(df_f)
                ok     = err_s == ""
                stats  = {"created": 0, "updated": 0, "deleted": 0, "skipped": 0, "failed": 0}
                for mbx, grp in df_f[ok].groupby(mbx_s[ok]):
                    if mbx not in cat_checked:
                        cat_checked.add(mbx)
                        if (cat_err := ensure_outlook_categories(token, mbx)):
                            log.append(f"⚠️ Event colors: {cat_err}")
                    for k, v in _sync_dataframe(token, mbx, grp, keep_ids).items():
                        stats[k] += v
                result = (f"Created:{stats['created']} Updated:{stats['updated']} "
                          f"Deleted:{stats['deleted']} Skipped:{stats['skipped']} Failed:{stats['failed']}")
                if not ok.all():
                    bad = [f"{e}: {m}" for e, m in zip(df_f.loc[~ok, "EVENT_ID"], err_s[~ok])]
                    result += f" Not pushed:{len(bad)} ({'; '.join(bad[:3])}{'; ...' if len(bad) > 3 else ''})"
                # Rows not pushed are retried next sweep (e.g. once their calendar mailbox is configured)
                state[fname] = {"last_modified": last_mod if ok.all() else None,
                                "last_synced":   dt.now(_PT).strftime("%Y-%m-%d %H:%M PT"),
                                "result":        result}
                log.append(f"✅ {fname} — {result}")
            except Exception as e:
                log.append(f"❌ {fname} — Error: {e}")

        _save_sweep_state(state)
    except Exception as e:
        status["error"] = str(e)
        log.append(f"❌ Sweep error: {e}")
    finally:
        status["last_check"] = dt.now(_PT).strftime("%Y-%m-%d %H:%M PT")
        status["running"]    = False
        status["log"]        = log

@st.cache_resource
def _ensure_scheduler():
    # Streamlit re-runs this script on every click; cache_resource keeps one scheduler per server process.
    sched = BackgroundScheduler(timezone="UTC")
    sched.add_job(run_sweep, "interval", minutes=15,
                  id="sp_sweep", next_run_time=dt.utcnow())
    sched.start()
    return sched

if _graph_secrets_ok():
    _ensure_scheduler()

# ─── SESSION STATE ───────────────────────────────────────────────────
for _k, _v in [("df", None), ("file_name", None), ("notifs", None), ("ai_done", False), ("sp_item_id", None)]:
    if _k not in st.session_state:
        st.session_state[_k] = _v
# Re-assign every run so the nav radio keeps its page while hidden during Change File
# (otherwise Streamlit drops its value and the server falls out of sync with the browser).
st.session_state.nav = st.session_state.get("nav", "Dashboard")

# ─── SIDEBAR ─────────────────────────────────────────────────────────
with st.sidebar:
    st.image("GES-logo-white.png", use_container_width=True)
    st.markdown("## GES SmartSchedule")
    st.caption("AI-Powered Electrical Labor Calendar")
    st.divider()

    if st.session_state.df is not None:
        st.success(f"**{st.session_state.file_name}**")
        st.caption(
            f"{len(st.session_state.df):,} rows  "
            f"({st.session_state.df['SCHEDULE_TYPE'].iloc[0]})"
        )
        if st.button("Change File", use_container_width=True):
            for k in ("df", "file_name", "notifs", "sp_item_id"):
                st.session_state[k] = None
            st.session_state.ai_done = False
            st.rerun()
        st.divider()
        page = st.radio(
            "Navigation",
            ["Dashboard", "Schedule & Push", "Calendar View", "Crew Lookup", "Overtime"],
            key="nav", label_visibility="collapsed",
        )
    else:
        st.warning("No file loaded")
        page = "_upload"

    st.divider()
    st.markdown("**Auto-Sync**")
    _cfg = _get_sweep_config()
    _sts = _get_sweep_status()
    _test_on = st.checkbox("Sweep test mode", value=_cfg["test_rows"] > 0)
    if _test_on:
        _n = st.number_input("Rows to sync", min_value=1, max_value=50,
                             value=max(_cfg["test_rows"], 2), step=1)
        _cfg["test_rows"] = int(_n)
    else:
        _cfg["test_rows"] = 0
    if _sts["running"]:
        st.info("Syncing now...")
    elif _sts["last_check"]:
        st.success(f"Last check: {_sts['last_check']}")
        for entry in _sts["log"]:
            st.caption(entry)
        if _sts["error"]:
            st.error(_sts["error"])
    else:
        st.caption("Waiting for first sweep...")

# ─── HEADER ──────────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 5])
with col_logo:
    st.image("GES-logo.webp", width=100)
with col_title:
    st.markdown("## GES SmartSchedule")
    st.caption("AI-Powered Electrical Labor Calendar — replacing manual Union Outlook entry")
st.divider()

# ═══════════════════════════════════════════════════════════════════
# UPLOAD — LV format ("By Day" sheet)
# ═══════════════════════════════════════════════════════════════════
if st.session_state.df is None:
    st.markdown('<div class="section-title">Load Schedule</div>', unsafe_allow_html=True)

    source = st.radio("Source", ["SharePoint", "Upload from computer"], horizontal=True)
    st.divider()

    file_bytes = None
    file_label = None

    if source == "SharePoint":
        if not _graph_secrets_ok():
            st.error("Graph credentials not configured — cannot connect to SharePoint.")
            st.stop()
        with st.spinner("Connecting to SharePoint..."):
            try:
                token    = get_graph_token()
                drive_id, sp_files = get_sharepoint_files(token)
            except Exception as e:
                st.error(f"SharePoint error: {e}")
                st.stop()

        if not sp_files:
            st.warning(f"No .xlsx files found in `{SP_FOLDER}`.")
            st.stop()

        options  = {f"{f['name']}  (modified {f['modified_display']})": f for f in sp_files}
        selected = st.selectbox("Select file", list(options.keys()))
        sp_file  = options[selected]

        if st.button("Load from SharePoint", type="primary", use_container_width=True):
            with st.spinner(f"Downloading {sp_file['name']}..."):
                try:
                    raw = download_sharepoint_file(token, drive_id, sp_file["id"])
                except Exception as e:
                    st.error(f"Download failed: {e}")
                    st.stop()
            try:
                xl    = pd.ExcelFile(io.BytesIO(raw))
                sheet = LV_SHEET if LV_SHEET in xl.sheet_names else xl.sheet_names[0]
                parsed = parse_lv_byDay(io.BytesIO(raw), sheet)
            except Exception as e:
                st.error(f"Could not parse file: {e}")
                st.stop()
            st.session_state.df        = parsed
            st.session_state.file_name = sp_file["name"]
            st.session_state.sp_item_id = sp_file["id"]
            st.session_state.notifs    = None
            st.session_state.ai_done   = False
            st.rerun()
        st.stop()

    else:
        st.markdown(
            "Upload the **GES Electrical Show Schedule** Excel file (.xlsx). "
            "The app reads the **By Day** sheet."
        )
        uploaded = st.file_uploader("Choose Excel file (.xlsx)", type=["xlsx"])
        if not uploaded:
            st.stop()

        file_bytes = uploaded.read()
        file_label = uploaded.name

        try:
            xl = pd.ExcelFile(io.BytesIO(file_bytes))
        except Exception as e:
            st.error(f"Could not open file: {e}")
            st.stop()

        if LV_SHEET in xl.sheet_names:
            sheet = LV_SHEET
            st.caption(f"Sheet auto-selected: **{sheet}**")
        else:
            sheet = st.selectbox(
                f"No '{LV_SHEET}' sheet found — select the scheduling sheet:",
                xl.sheet_names,
            )

        try:
            parsed = parse_lv_byDay(io.BytesIO(file_bytes), sheet)
        except Exception as e:
            st.error(f"Could not parse file: {e}")
            st.stop()

        st.success(f"**{len(parsed):,} shifts** found across **{parsed['SHOW'].nunique()}** shows.")
        _, _errs = _route_rows(parsed)
        if (_errs != "").any():
            st.warning(f"**{int((_errs != '').sum())}** row(s) have a FACILITY problem and will not be pushed "
                       f"to Outlook (details on the Push to Outlook tab).")

        with st.expander("Preview (first 10 rows)"):
            _fac = ["FACILITY"] if has_facility(parsed["FACILITY"].iloc[0]) else []
            p = parsed[["DATE","SHOW"] + _fac + ["FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW"]].head(10).copy()
            p["DATE"] = p["DATE"].dt.strftime("%Y-%m-%d")
            st.dataframe(p, hide_index=True, use_container_width=True)

        if st.button("Load Schedule", type="primary", use_container_width=True):
            st.session_state.df        = parsed
            st.session_state.file_name = file_label
            st.session_state.sp_item_id = None
            st.session_state.notifs    = None
            st.session_state.ai_done   = False
            st.rerun()

        st.stop()

# ─── From here, df is guaranteed loaded ──────────────────────────────
df    = st.session_state.df
stype = df["SCHEDULE_TYPE"].iloc[0]

# ═══════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ═══════════════════════════════════════════════════════════════════
if "Dashboard" in page:
    total  = len(df)
    shows  = df["SHOW"].nunique()
    crew   = int(df["TOTAL_CREW"].sum())
    dr_min = df["DATE"].min().strftime("%b %d, %Y")
    dr_max = df["DATE"].max().strftime("%b %d, %Y")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Total Shifts / Events</div>
            <div class="metric-value">{total}</div>
            <div class="metric-delta">{st.session_state.file_name}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Unique Shows</div>
            <div class="metric-value">{shows}</div>
            <div class="metric-delta">Across all locations</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Date Range</div>
            <div class="metric-value" style="font-size:18px">{dr_min}</div>
            <div class="metric-delta">through {dr_max}</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Crew Calls (Total)</div>
            <div class="metric-value">{crew}</div>
            <div class="metric-delta">From crew count column</div>
        </div>""", unsafe_allow_html=True)

    st.divider()
    st.markdown('<div class="section-title">Loaded Schedule — Recent 20 Rows</div>', unsafe_allow_html=True)
    recent = df.sort_values("DATE").tail(20)[
        ["DATE","SHOW","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW"]
    ].copy()
    recent["DATE"] = recent["DATE"].dt.strftime("%Y-%m-%d")
    st.dataframe(recent, use_container_width=True, hide_index=True)

    st.divider()
    fn_counts = df["FUNCTION"].value_counts().reset_index()
    fn_counts.columns = ["Function", "Count"]
    fig = px.bar(fn_counts, x="Function", y="Count", color="Function",
                 color_discrete_map=FUNCTION_COLORS, title="Shifts by Function Type")
    fig.update_layout(showlegend=False, height=300, margin=dict(t=40, b=10, l=10, r=10))
    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════
# PAGE: SCHEDULE & PUSH
# ═══════════════════════════════════════════════════════════════════
elif "Schedule" in page:
    with st.expander("Notification Settings", expanded=False):
        dc1, dc2 = st.columns(2)
        with dc1:
            union_email   = st.text_input("Union dispatch email", value="union-dispatch@ibew-local.org")
        with dc2:
            foreman_email = st.text_input("Foreman email",        value="foreman@ges.com")
    demo_contacts = {"union_email": union_email, "foreman_email": foreman_email}

    tab_prev, tab_email, tab_push = st.tabs(["Preview", "Emails", "Push to Outlook"])

    # ── Preview ──────────────────────────────────────────────────────
    with tab_prev:
        st.caption(f"Source: **{st.session_state.file_name}** · {len(df):,} rows · {stype}")
        show_cols = [c for c in ["DATE","SHOW","FACILITY","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW"]
                     if c in df.columns and (c != "FACILITY" or has_facility(df["FACILITY"].iloc[0]))]
        d = df[show_cols].copy()
        d["DATE"] = d["DATE"].dt.strftime("%Y-%m-%d")
        st.dataframe(d, use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("""<div class="ai-banner">
            <b>AI Parsing Engine — Ready</b><br><br>
            Click Preview to review what will be created in Outlook before pushing.
        </div>""", unsafe_allow_html=True)

        if st.button("Preview Schedule", type="primary", use_container_width=True):
            st.session_state.notifs  = build_notifications(df, demo_contacts)
            st.session_state.ai_done = True

        if st.session_state.ai_done and st.session_state.notifs is not None:
            df_n = st.session_state.notifs
            c1, c2, c3 = st.columns(3)
            c1.metric("Calendar Events",     len(df))
            c2.metric("Email Notifications", len(df_n))
            c3.metric("Manual Steps",        "0")
            st.info("Switch to the Push to Outlook tab when ready.")

    # ── Emails ───────────────────────────────────────────────────────
    with tab_email:
        st.markdown('<div class="section-title">Email Notifications</div>', unsafe_allow_html=True)
        if st.session_state.notifs is None:
            st.info("Go to the Preview tab and click Preview Schedule first.")
            if st.button("Quick-generate with default settings"):
                st.session_state.notifs  = build_notifications(df, {})
                st.session_state.ai_done = True
                st.rerun()
        else:
            df_n = st.session_state.notifs
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Emails",    len(df_n))
            c2.metric("Labor Calls",     int((df_n["Type"] == "Labor Call").sum()))
            c3.metric("Foreman Notices", int((df_n["Type"] == "Foreman Assignment").sum()))
            st.divider()
            nf1, nf2 = st.columns(2)
            with nf1:
                type_sel = st.multiselect("Type", df_n["Type"].unique().tolist(),
                                          default=df_n["Type"].unique().tolist(), key="n_type")
            with nf2:
                top_n    = df_n["Show"].value_counts().head(15).index.tolist()
                show_sel = st.multiselect("Show", top_n, default=top_n[:5], key="n_show")
            df_nf = df_n[df_n["Type"].isin(type_sel) & df_n["Show"].isin(show_sel)]
            for _, r in df_nf.head(50).iterrows():
                st.markdown(f"""<div class="notif-card">
                    <b>{r['Type']}</b> &nbsp;|&nbsp; To: <b>{r['Recipient']}</b>
                    &nbsp;|&nbsp; {r['Date']} &nbsp; {r['Status']}<br>
                    <small style="color:#64748b;">{r['Message']}</small>
                </div>""", unsafe_allow_html=True)
            if len(df_nf) > 50:
                st.caption(f"Showing 50 of {len(df_nf)}.")
            st.download_button("Export email log (.csv)", df_nf.to_csv(index=False).encode(),
                               "ges_notifications.csv", "text/csv")

    # ── Push to Outlook ──────────────────────────────────────────────
    with tab_push:
        st.markdown('<div class="section-title">Push to Shared Outlook Calendar</div>', unsafe_allow_html=True)
        if not _graph_secrets_ok():
            st.error("Microsoft Graph credentials not configured.")
            st.markdown("""
Fill in `.streamlit/secrets.toml` (or Azure App Service Application Settings):
```toml
[graph]
app_id         = "2fafc3de-feac-47cf-b231-f16de70392c3"
tenant_id      = "286f2b7d-9113-445b-ae51-103e88d4584a"
client_secret  = "your-secret"
shared_mailbox = "AzureAutomationTest-PIT65@viadcorp.onmicrosoft.com"
```
            """)
        else:
            g              = _graph_cfg()
            shared_mailbox = g["shared_mailbox"]
            # Each row goes to its facility's calendar (files without FACILITY: by ID prefix)
            mbx_s, err_s = _route_rows(df)
            ok_s = err_s == ""
            if not has_facility(df["FACILITY"].iloc[0]):
                st.warning(f"This file has no **{FACILITY_COL}** column, so rows are routed by ID prefix "
                           f"(older format). Move the rows to the current template to route by facility.")
                label = df["BRANCH"].replace("", "no prefix")
            else:  # blank FACILITY rows go by city: show them as e.g. 'ORL (no facility)'
                label = df["FACILITY"].where(df["FACILITY"] != "",
                                             df["BRANCH"].replace("", "no prefix") + " (no facility)")
            if ok_s.any():
                dest = label[ok_s].groupby(mbx_s[ok_s]).unique()
                st.info("Pushing to: " + "; ".join(f"**{', '.join(sorted(ks))}** → {m}"
                                                   for m, ks in dest.items()))
            if not ok_s.all():
                st.error(f"**{int((~ok_s).sum())}** row(s) will NOT be pushed until their FACILITY is fixed "
                         f"or its calendar is configured:")
                bad = df.loc[~ok_s, ["EVENT_ID", "DATE", "SHOW", "FACILITY"]].assign(Problem=err_s[~ok_s])
                bad["DATE"] = bad["DATE"].dt.strftime("%Y-%m-%d")
                st.dataframe(bad, hide_index=True, use_container_width=True)

            col_test, _ = st.columns([1, 3])
            with col_test:
                if st.button("Test Connection"):
                    with st.spinner("Authenticating..."):
                        try:
                            tok  = get_graph_token()
                            resp = requests.get(
                                f"https://graph.microsoft.com/v1.0/users/{shared_mailbox}",
                                headers={"Authorization": f"Bearer {tok}"}, timeout=10,
                            )
                            if resp.status_code == 200:
                                st.success(f"Connected: **{resp.json().get('displayName', shared_mailbox)}**")
                            else:
                                st.error(f"Failed ({resp.status_code})")
                        except Exception as e:
                            st.error(f"Auth error: {e}")

            st.divider()
            st.caption(f"**{len(df):,}** events from **{st.session_state.file_name}** ready to push.")
            with st.expander("Preview first 10 events"):
                for _, row in df.head(10).iterrows():
                    start = row["DATE"].strftime("%Y-%m-%d")
                    end   = row["END_DATE"].strftime("%Y-%m-%d") if pd.notna(row["END_DATE"]) else start
                    label = start if start == end else f"{start} to {end}"
                    st.write(f"**{row['SHOW']}** | {label} | {row['LOCATION'] or '—'}")

            test_mode = st.checkbox("Test mode (push a few events only)", value=True)
            test_n    = st.number_input("Number of events to push", min_value=1, max_value=len(df),
                                        value=2, step=1) if test_mode else len(df)
            df_push   = df.head(int(test_n))

            if test_mode:
                st.caption(f"Test mode: will push **{int(test_n)}** of {len(df)} events.")

            if st.button("Push to Shared Calendar", type="primary", use_container_width=True):
                try:
                    token = get_graph_token()
                except Exception as e:
                    st.error(f"Auth failed: {e}")
                    st.stop()

                # A SharePoint file may have been moved or deleted since it was loaded
                if st.session_state.sp_item_id:
                    try:
                        _, sp_now = get_sharepoint_files(token)
                    except Exception as e:
                        st.error(f"Could not check SharePoint: {e}")
                        st.stop()
                    if st.session_state.sp_item_id not in {f["id"] for f in sp_now}:
                        st.error(f"**{st.session_state.file_name}** is no longer in the SharePoint "
                                 f"folder, so nothing was pushed. Click Change File to load another file.")
                        st.stop()

                created = updated = failed = skipped = deleted = 0
                results = []
                stat = st.empty()

                stat.text("Reading shared calendar from Outlook...")
                cal_maps = {m: build_calendar_map(token, m, df["DATE"].min(), df["DATE"].max())
                            for m in set(mbx_s[ok_s])}
                for m in cal_maps:  # SET / STRIKE colors (OUTLOOK_CATEGORIES); push continues either way
                    if (cat_err := ensure_outlook_categories(token, m)):
                        st.warning(f"Event colors: {cat_err}")

                bar = st.progress(0)
                for i, (idx, row) in enumerate(df_push.iterrows()):
                    eid     = row["EVENT_ID"]
                    if err_s[idx]:
                        failed += 1
                        results.append({"ID": eid, "Event": row["SHOW"], "Status": f"Not pushed: {err_s[idx]}"})
                        bar.progress((i + 1) / len(df_push))
                        continue
                    subject = f"{row['SHOW']} - {row['FUNCTION']}"
                    end_dt  = row["END_DATE"] if pd.notna(row["END_DATE"]) else row["DATE"]
                    notes   = _event_notes(row)
                    fhash   = _field_hash(row)
                    mbx     = mbx_s[idx]
                    cal_map = cal_maps[mbx]
                    stat.text(f"Processing {i+1}/{len(df_push)}: {str(row['SHOW'])[:60]}...")
                    try:
                        entry = cal_map.get(eid)
                        if entry and entry["hash"] == fhash:
                            skipped += 1
                            results.append({"ID": eid, "Event": row["SHOW"], "Status": "Skipped (no change)"})
                        elif entry:
                            resp = patch_graph_event(
                                token, mbx, entry["graph_id"],
                                row["DATE"], pd.Timestamp(end_dt),
                                row["LOCATION"], notes,
                                time_str=row["TIME"], event_id=eid, field_hash=fhash,
                                time_zone=row.get("TIME_ZONE", DEFAULT_TZ), categories=_event_categories(row),
                                subject=subject,
                            )
                            if resp.status_code == 200:
                                updated += 1
                                results.append({"ID": eid, "Event": row["SHOW"], "Status": "Updated"})
                            elif resp.status_code == 404:
                                resp2 = create_graph_event(
                                    token, mbx, subject,
                                    row["DATE"], pd.Timestamp(end_dt),
                                    row["LOCATION"], notes,
                                    time_str=row["TIME"], event_id=eid, field_hash=fhash,
                                    time_zone=row.get("TIME_ZONE", DEFAULT_TZ), categories=_event_categories(row),
                                )
                                if resp2.status_code == 201:
                                    created += 1
                                    results.append({"ID": eid, "Event": row["SHOW"], "Status": "Created (restored)"})
                                else:
                                    failed += 1
                                    msg = resp2.json().get("error", {}).get("message", "")[:80]
                                    results.append({"ID": eid, "Event": row["SHOW"], "Status": f"Failed: {msg}"})
                            else:
                                failed += 1
                                msg = resp.json().get("error", {}).get("message", "")[:80]
                                results.append({"ID": eid, "Event": row["SHOW"], "Status": f"Failed: {msg}"})
                        else:
                            resp = create_graph_event(
                                token, mbx, subject,
                                row["DATE"], pd.Timestamp(end_dt),
                                row["LOCATION"], notes,
                                time_str=row["TIME"], event_id=eid, field_hash=fhash,
                                time_zone=row.get("TIME_ZONE", DEFAULT_TZ), categories=_event_categories(row),
                            )
                            if resp.status_code == 201:
                                created += 1
                                results.append({"ID": eid, "Event": row["SHOW"], "Status": "Created"})
                            else:
                                failed += 1
                                msg = resp.json().get("error", {}).get("message", "")[:80]
                                results.append({"ID": eid, "Event": row["SHOW"], "Status": f"Failed: {msg}"})
                    except Exception as e:
                        failed += 1
                        results.append({"ID": eid, "Event": row["SHOW"], "Status": f"Error: {e}"})
                    bar.progress((i + 1) / len(df_push))

                # Delete Outlook events whose ID is no longer in the full Excel, or whose row now goes
                # to a different calendar. Keep IDs from the other SharePoint files so this file
                # doesn't delete their events. excel_ids: {EVENT_ID: mailbox it routes to, or None}
                del_ok = False
                if not test_mode:
                    excel_ids = {}
                    try:
                        _, others, errs = _load_folder(token)
                        if errs:
                            st.warning("Deletes skipped — could not read: " + ", ".join(errs))
                        else:
                            for name, d in others.items():
                                if name != st.session_state.file_name:
                                    excel_ids.update(zip(d["EVENT_ID"], _route_rows(d)[0]))
                            excel_ids.update(zip(df["EVENT_ID"], mbx_s))
                            del_ok = True
                    except Exception as e:
                        st.warning(f"Deletes skipped — could not read SharePoint folder: {e}")
                if del_ok:
                    for eid, mbx, entry in ((e, m, x) for m, cm in cal_maps.items() for e, x in cm.items()):
                        if eid not in excel_ids or excel_ids[eid] not in (None, mbx):
                            stat.text(f"Deleting removed event: {eid}...")
                            try:
                                dresp = delete_graph_event(token, mbx, entry["graph_id"])
                                if dresp.status_code in (204, 200):
                                    deleted += 1
                                    results.append({"ID": eid, "Event": "—", "Status": "Deleted (removed from Excel)"})
                            except Exception as e:
                                results.append({"ID": eid, "Event": "—", "Status": f"Delete error: {e}"})

                stat.empty()
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Created",  created)
                c2.metric("Updated",  updated)
                c3.metric("Deleted",  deleted)
                c4.metric("Skipped",  skipped)
                c5.metric("Failed",   failed)
                st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)
                if failed == 0:
                    st.success("Done. Check the shared Outlook calendar.")
                else:
                    st.warning(f"{failed} event(s) failed. See table above.")


# ═══════════════════════════════════════════════════════════════════
# PAGE: CALENDAR VIEW
# ═══════════════════════════════════════════════════════════════════
elif "Calendar" in page:
    st.markdown('<div class="section-title">Labor Calendar</div>', unsafe_allow_html=True)
    st.caption(f"Source: **{st.session_state.file_name}** · {stype}")

    # ── Filters ──────────────────────────────────────────────────────
    f1, f2, f3, f4 = st.columns([1, 2, 2, 3])
    with f1:
        years   = sorted(df["DATE"].dt.year.dropna().unique().tolist())
        yr_sel  = st.selectbox("Year", ["All"] + [str(y) for y in years])
    with f2:
        fn_opts = sorted(df["FUNCTION"].dropna().unique().tolist())
        fn_sel  = st.multiselect("Function", fn_opts, default=fn_opts)
    with f3:
        locs    = [l for l in sorted(df["LOCATION"].dropna().unique().tolist()) if l]
        loc_sel = st.selectbox("Location", ["All Locations"] + locs)
    with f4:
        shows    = sorted(df["SHOW"].dropna().unique().tolist())
        show_sel = st.selectbox("Show", ["All Shows"] + shows)

    df_filt = df.copy()
    if yr_sel != "All":
        df_filt = df_filt[df_filt["DATE"].dt.year == int(yr_sel)]
    if fn_sel:
        df_filt = df_filt[df_filt["FUNCTION"].isin(fn_sel)]
    if loc_sel != "All Locations":
        df_filt = df_filt[df_filt["LOCATION"] == loc_sel]
    if show_sel != "All Shows":
        df_filt = df_filt[df_filt["SHOW"] == show_sel]

    st.caption(f"**{len(df_filt)}** shifts matching filters")

    # ── Views ────────────────────────────────────────────────────────
    v_gantt, v_grid, v_table, v_monthly = st.tabs(
        ["Gantt", "Calendar Grid", "Table", "Monthly Summary"]
    )

    with v_gantt:
        if df_filt.empty:
            st.warning("No shifts match the current filters.")
        else:
            df_g = df_filt.copy()
            has_end = stype == "events" and df_g["END_DATE"].notna().any()
            if has_end:
                df_g["END_DATE"] = df_g["END_DATE"].fillna(df_g["DATE"])
                fig = px.timeline(
                    df_g.sort_values("DATE"), x_start="DATE", x_end="END_DATE",
                    y="SHOW", color="FUNCTION", color_discrete_map=FUNCTION_COLORS,
                    hover_data={"LOCATION": True, "FOREMAN": True, "TOTAL_CREW": True,
                                "DATE": False, "END_DATE": False},
                    title="Show Timeline",
                )
            else:
                df_g["_END"] = df_g["DATE"] + pd.Timedelta(hours=8)
                fig = px.timeline(
                    df_g.sort_values("DATE"), x_start="DATE", x_end="_END",
                    y="SHOW", color="FUNCTION", color_discrete_map=FUNCTION_COLORS,
                    hover_data={"LOCATION": True, "TIME": True, "FOREMAN": True,
                                "TOTAL_CREW": True, "DATE": False, "_END": False},
                    title="Labor Schedule",
                )
            fig.update_layout(
                height=max(420, len(df_g["SHOW"].unique()) * 22 + 100),
                xaxis_title="", yaxis_title="", legend_title="Function",
                plot_bgcolor="#f8fafc", paper_bgcolor="white",
                font=dict(size=12), margin=dict(l=10, r=10, t=40, b=20),
            )
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(fig, use_container_width=True)

    with v_grid:
        if df_filt.empty:
            st.warning("No shifts match the current filters.")
        else:
            df_g = df_filt.copy()
            df_g["DayOfMonth"] = df_g["DATE"].dt.day
            df_g["MonthYear"]  = df_g["DATE"].dt.strftime("%b %Y")
            month_order = (
                df_g["DATE"].dt.to_period("M").drop_duplicates().sort_values()
                    .apply(lambda p: p.strftime("%b %Y")).tolist()
            )
            pivot = (
                df_g.groupby(["MonthYear", "DayOfMonth"]).size()
                    .reset_index(name="Shifts")
                    .pivot(index="MonthYear", columns="DayOfMonth", values="Shifts")
                    .fillna(0)
                    .reindex([m for m in month_order if m in df_g["MonthYear"].values])
            )
            fig_grid = px.imshow(
                pivot,
                labels=dict(x="Day", y="Month", color="Shifts"),
                color_continuous_scale=[[0, "#f0f9ff"], [0.3, "#93c5fd"], [1, "#1d4ed8"]],
                title="Shift Density — darker = more shifts that day",
                aspect="auto",
            )
            fig_grid.update_layout(
                height=max(260, len(pivot) * 42 + 100),
                margin=dict(l=10, r=10, t=40, b=20),
            )
            fig_grid.update_xaxes(side="top", tickmode="linear", tick0=1, dtick=1)
            st.plotly_chart(fig_grid, use_container_width=True)

            st.markdown("**Breakdown by function:**")
            fn_monthly = (
                df_g.groupby([df_g["DATE"].dt.strftime("%b %Y"), "FUNCTION"])
                    .size().reset_index(name="Shifts")
            )
            fn_monthly.columns = ["Month", "Function", "Shifts"]
            fig_fn = px.bar(fn_monthly, x="Month", y="Shifts", color="Function",
                            color_discrete_map=FUNCTION_COLORS, barmode="group",
                            category_orders={"Month": month_order})
            fig_fn.update_layout(height=300, margin=dict(l=10, r=10, t=20, b=20))
            st.plotly_chart(fig_fn, use_container_width=True)

    with v_table:
        if df_filt.empty:
            st.warning("No shifts match the current filters.")
        else:
            d = df_filt[["DATE","SHOW","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW","NOTES"]].copy()
            d["DATE"] = d["DATE"].dt.strftime("%Y-%m-%d")
            st.dataframe(d.sort_values("DATE"), use_container_width=True, hide_index=True)

    with v_monthly:
        if df_filt.empty:
            st.warning("No shifts match the current filters.")
        else:
            df_m = df_filt.copy()
            df_m["Month"] = df_m["DATE"].dt.to_period("M").astype(str)
            monthly = df_m.groupby(["Month", "FUNCTION"]).size().reset_index(name="Shifts")
            fig2 = px.bar(monthly, x="Month", y="Shifts", color="FUNCTION",
                          color_discrete_map=FUNCTION_COLORS, barmode="stack",
                          title="Shifts per Month by Function")
            fig2.update_layout(height=380, margin=dict(t=40, b=20, l=10, r=10))
            st.plotly_chart(fig2, use_container_width=True)

    # ── Shift Detail ─────────────────────────────────────────────────
    if not df_filt.empty:
        st.divider()
        st.markdown('<div class="section-title">Shift Detail</div>', unsafe_allow_html=True)
        sel = st.selectbox("Select show", sorted(df_filt["SHOW"].unique().tolist()), key="cal_show")
        for _, r in df_filt[df_filt["SHOW"] == sel].sort_values("DATE").iterrows():
            fn  = r["FUNCTION"]
            cls = {"SET": "fn-set", "STRIKE": "fn-strike", "STANDBY": "fn-standby"}.get(fn, "fn-set")
            end_label = (f" → {r['END_DATE'].strftime('%b %d')}"
                         if pd.notna(r["END_DATE"]) else "")
            st.markdown(f"""
**{r['DATE'].strftime('%a %b %d, %Y')}**{end_label} &nbsp;
<span class="fn-badge {cls}">{fn}</span> &nbsp; {r['TIME']} &nbsp; {r['LOCATION']}

Foreman: {r['FOREMAN']} &nbsp; {r['NOTES']}

---""", unsafe_allow_html=True)

    # ── Exports ───────────────────────────────────────────────────────
    st.divider()
    st.markdown('<div class="section-title">Export Calendar</div>', unsafe_allow_html=True)
    st.caption(
        "Each .ics file imports as a separate calendar in Outlook. "
        "Importing by function type keeps SET, STRIKE, and STANDBY events visually distinct."
    )

    e1, e2, e3, e4, e5 = st.columns(5)
    with e1:
        st.download_button("All Shifts (.ics)",
            data=_build_ics(df_filt, "All"), file_name="ges_all.ics",
            mime="text/calendar", use_container_width=True, disabled=df_filt.empty)
    for fn_type, col_w in zip(["SET", "STRIKE", "STANDBY"], [e2, e3, e4]):
        df_fn = df_filt[df_filt["FUNCTION"] == fn_type]
        with col_w:
            st.download_button(f"{fn_type} (.ics)",
                data=_build_ics(df_fn, fn_type), file_name=f"ges_{fn_type.lower()}.ics",
                mime="text/calendar", use_container_width=True, disabled=df_fn.empty)
    with e5:
        csv_out = df_filt[["DATE","SHOW","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW","NOTES"]].copy()
        csv_out["DATE"] = csv_out["DATE"].dt.strftime("%Y-%m-%d")
        st.download_button("Export (.csv)",
            data=csv_out.to_csv(index=False).encode(), file_name="ges_schedule.csv",
            mime="text/csv", use_container_width=True, disabled=df_filt.empty)


# ═══════════════════════════════════════════════════════════════════
# PAGE: CREW LOOKUP
# ═══════════════════════════════════════════════════════════════════
elif "Crew" in page:
    st.markdown('<div class="section-title">Crew Lookup</div>', unsafe_allow_html=True)

    has_crew = df["CREW_LIST"].apply(lambda x: len(x) > 0).any()
    if not has_crew:
        st.info(
            "Crew Lookup requires individual crew member columns in your Excel file "
            "(columns named 1, 2, 3 ... for each crew member slot). "
            "Your current file does not have those columns."
        )
        st.stop()

    all_crew = sorted({
        p.replace(" (Foreman)", "").strip()
        for _, row in df.iterrows()
        for p in row["CREW_LIST"]
        if p.replace(" (Foreman)", "").strip() not in ("TBD", "nan", "")
    })

    if not all_crew:
        st.warning("No named crew members found.")
        st.stop()

    selected = st.selectbox("Select electrician", all_crew)
    person_rows = []
    for _, row in df.iterrows():
        names = [p.replace(" (Foreman)", "").strip() for p in row["CREW_LIST"]]
        if selected in names:
            role = "Foreman" if any(selected in p and "(Foreman)" in p for p in row["CREW_LIST"]) else "Crew"
            person_rows.append({
                "Date": row["DATE"].strftime("%Y-%m-%d"),
                "Show": row["SHOW"], "Function": row["FUNCTION"],
                "Location": row["LOCATION"], "Time": row["TIME"],
                "Role": role, "Notes": row["NOTES"],
            })

    if person_rows:
        df_p = pd.DataFrame(person_rows).sort_values("Date")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Shifts", len(df_p))
        c2.metric("As Foreman",   int((df_p["Role"] == "Foreman").sum()))
        c3.metric("As Crew",      int((df_p["Role"] == "Crew").sum()))
        st.dataframe(df_p, use_container_width=True, hide_index=True)

        crew_df = df[df["CREW_LIST"].apply(lambda lst: any(selected in p for p in lst))]
        st.download_button(
            f"Download {selected}'s schedule (.ics)",
            data=_build_ics(crew_df, f"Crew {selected}"),
            file_name=f"ges_{selected.replace(' ', '_')}.ics",
            mime="text/calendar",
        )
        st.info("Double-click the .ics file to import directly into Outlook.")
    else:
        st.info(f"No shifts found for **{selected}**.")


# ═══════════════════════════════════════════════════════════════════
# PAGE: OVERTIME / DOUBLE-TIME
# ═══════════════════════════════════════════════════════════════════
elif "Overtime" in page:
    st.markdown('<div class="section-title">Overtime</div>', unsafe_allow_html=True)
    st.caption(f"Source: **{st.session_state.file_name}** · rules and holidays are set in schedule_config.py")

    ot_all = compute_overtime(df)  # on the whole file, so daily totals see every shift before filtering
    if ot_all.empty:
        st.info("No crew names or +N crew in this file, so there is nothing to count.")
        st.stop()

    # ── Filters ──────────────────────────────────────────────────────
    f1, f2 = st.columns([2, 3])
    with f1:
        d_min, d_max = df["DATE"].min().date(), df["DATE"].max().date()
        d_rng = st.date_input("Dates", (d_min, d_max), min_value=d_min, max_value=d_max, key="ot_dates")
    with f2:
        fac_opts = sorted(ot_all["FACILITY"].unique().tolist())
        fac_sel  = st.multiselect("Facility", fac_opts, default=fac_opts, key="ot_fac",
                                  format_func=lambda c: f"{c} - {FACILITIES[c]['name']}" if c in FACILITIES
                                                        else (c or "(no facility)"))
    d_from, d_to = (d_rng[0], d_rng[-1]) if d_rng else (d_min, d_max)
    ot = ot_all[ot_all["DATE"].dt.date.between(d_from, d_to) & ot_all["FACILITY"].isin(fac_sel)]

    n_norule = int((ot["STATUS"] == "no OT rules").sum())
    n_notime = int((ot["STATUS"] == "no TIME").sum())
    if n_norule:
        st.warning(f"{n_norule} crew shift(s) are at facilities/branches with no OT rules configured yet; "
                   f"their hours are counted as regular.")
    if n_notime:
        st.caption(f"{n_notime} crew shift(s) have no TIME, so their hours are unknown (counted as 0).")

    t_team, t_ind, t_dt_team, t_dt_ind = st.tabs(
        ["OT - Team", "OT - Individual", "Double-Time - Team", "Double-Time - Individual"])

    with t_team:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Crew Shifts", len(ot))
        c2.metric("Total Hours", f"{ot['HOURS'].sum():,.2f}")
        c3.metric("OT Shifts",   int(ot["OT_SHIFT"].sum()))
        c4.metric("OT Hours",    f"{ot['OT_HOURS'].sum():,.2f}")
        st.markdown("**By electrician**")
        st.dataframe(ot_summary(ot, "PERSON").round(2), hide_index=True, use_container_width=True)
        st.caption("Unnamed crew = all +N slots combined; each is counted per shift (they can't be tracked "
                   "across shifts).")
        st.markdown("**By facility**")
        st.dataframe(ot_summary(ot, "FACILITY").round(2), hide_index=True, use_container_width=True)
        st.download_button("Export OT detail (.csv)", ot.to_csv(index=False).encode(),
                           "ges_overtime.csv", "text/csv")

    with t_ind:
        names = sorted(ot.loc[ot["NAMED"], "PERSON"].unique().tolist())
        if not names:
            st.info("No named electricians match the filters.")
        else:
            who = st.selectbox("Select electrician", names, key="ot_person")
            mine = ot[ot["PERSON"] == who]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Shifts",      len(mine))
            c2.metric("Total Hours", f"{mine['HOURS'].sum():,.2f}")
            c3.metric("OT Shifts",   int(mine["OT_SHIFT"].sum()))
            c4.metric("OT Hours",    f"{mine['OT_HOURS'].sum():,.2f}")
            d = mine[["DATE", "SHOW", "FACILITY", "FUNCTION", "TIME", "HOURS", "REG_HOURS", "OT_HOURS",
                      "OT_REASON", "STATUS"]].copy()
            d.insert(1, "DAY", d["DATE"].dt.strftime("%a"))
            d["DATE"] = d["DATE"].dt.strftime("%Y-%m-%d")
            st.dataframe(d.round(2), hide_index=True, use_container_width=True)

    for tab in (t_dt_team, t_dt_ind):
        with tab:
            if not DT_RULES:
                st.info("Double-Time rules have not been provided yet, so double-time is not calculated. "
                        "This view will show double-time shifts and hours once the rules are added "
                        "(DT_RULES in schedule_config.py).")

    with st.expander("OT rules and holidays in use"):
        for region, rules in OT_RULES.items():
            facs = ", ".join(c for c, f in FACILITIES.items() if f["ot_region"] == region)
            st.markdown(f"**{region}** ({facs}): over {rules['daily_hours']} hrs on a normal workday"
                        + (", Saturday" if rules.get("saturday") else "")
                        + (", Sunday" if rules.get("sunday") else "")
                        + (", holidays" if rules.get("holidays") else "") + " = OT")
            if rules.get("holidays"):
                st.caption("Holidays: " + ", ".join(pd.to_datetime(HOLIDAYS.get(region, []))
                                                    .strftime("%a %b %d, %Y")))
