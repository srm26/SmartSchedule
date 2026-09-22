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
import threading
from pathlib import Path
from datetime import datetime as dt
from zoneinfo import ZoneInfo

_PT = ZoneInfo("America/Los_Angeles")
import msal
import requests
from apscheduler.schedulers.background import BackgroundScheduler

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
  [data-testid="stSidebar"] .stRadio label { font-size: 15px; padding: 6px 0; }
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

# ─── OUTLOOK BODY MARKERS ────────────────────────────────────────────
# EVENT_ID and field hash are stored in the first line of the event body.
# Outlook is the source of truth — no local file needed.

def _field_hash(row):
    key = "|".join(str(row.get(f, "")) for f in ["SHOW", "FUNCTION", "TIME", "LOCATION", "FOREMAN", "NOTES"])
    return hashlib.md5(key.encode()).hexdigest()[:8]

def _make_body(notes, event_id, field_hash):
    marker = f"GES_ID:{event_id}|GES_HASH:{field_hash}"
    return f"{marker}\n{notes}".strip() if notes else marker

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
                      location="", notes="", time_str="", event_id="", field_hash=""):
    url     = f"https://graph.microsoft.com/v1.0/users/{mailbox}/events/{graph_id}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    t_start, t_end = _parse_time_range(time_str)
    body_text = _make_body(notes, event_id, field_hash)
    if t_start and t_end:
        payload = {
            "isAllDay": False,
            "start": {"dateTime": start_dt.strftime(f"%Y-%m-%dT{t_start}:00"), "timeZone": "Pacific Standard Time"},
            "end":   {"dateTime": end_dt.strftime(f"%Y-%m-%dT{t_end}:00"),     "timeZone": "Pacific Standard Time"},
            "location": {"displayName": location or "Not specified"},
            "body": {"contentType": "Text", "content": body_text},
        }
    else:
        payload = {
            "isAllDay": True,
            "start": {"dateTime": start_dt.strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"},
            "end":   {"dateTime": (end_dt + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"},
            "location": {"displayName": location or "Not specified"},
            "body": {"contentType": "Text", "content": body_text},
        }
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

def create_graph_event(token, mailbox, subject, start_dt, end_dt,
                       location="", notes="", time_str="", event_id="", field_hash=""):
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/events"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    t_start, t_end = _parse_time_range(time_str)
    body_text = _make_body(notes, event_id, field_hash)
    if t_start and t_end:
        payload = {
            "subject": subject, "isAllDay": False,
            "start": {"dateTime": start_dt.strftime(f"%Y-%m-%dT{t_start}:00"), "timeZone": "Pacific Standard Time"},
            "end":   {"dateTime": end_dt.strftime(f"%Y-%m-%dT{t_end}:00"),   "timeZone": "Pacific Standard Time"},
            "location": {"displayName": location or "Not specified"},
            "body": {"contentType": "Text", "content": body_text},
        }
    else:
        payload = {
            "subject": subject, "isAllDay": True,
            "start": {"dateTime": start_dt.strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"},
            "end":   {"dateTime": (end_dt + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"},
            "location": {"displayName": location or "Not specified"},
            "body": {"contentType": "Text", "content": body_text},
        }
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
         "modified_display": i.get("lastModifiedDateTime", "")[:16].replace("T", " ")}  # for UI
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
def parse_lv_byDay(uploaded_file, sheet_name):
    """Parse the 'By Day' sheet from the LV GES Electrical Show Schedule format.

    Expected columns: DAY, DATE, TIME, SHOW, FUNCTION, LOCATION, NOTES, FOREMAN, 1..7
    Crew columns 1-7 hold names or '+N' (extra unnamed crew).
    """
    raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)

    required = {"DATE", "SHOW", "FUNCTION", "LOCATION", "NOTES", "FOREMAN"}
    present  = {str(c).strip().upper() for c in raw.columns}
    missing  = required - present
    if missing:
        raise ValueError(f"Sheet '{sheet_name}' is missing expected columns: {', '.join(sorted(missing))}")

    crew_cols = [c for c in raw.columns if isinstance(c, int)]

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
        raw_id = str(r.get("ID", "") or "").strip()
        event_id = raw_id if raw_id and raw_id.lower() not in ("nan", "none") else _make_event_id(show, fn, date)
        rows.append({
            "EVENT_ID":      event_id,
            "DATE":          date,
            "END_DATE":      pd.NaT,
            "SHOW":          show,
            "FUNCTION":      fn,
            "LOCATION":      str(r.get("LOCATION", "") or "").strip(),
            "TIME":          str(r.get("TIME", "") or "").strip(),
            "FOREMAN":       foreman,
            "NOTES":         str(r.get("NOTES", "") or "").strip(),
            "TOTAL_CREW":    max(len(crew_list) + extra, 1),
            "CREW_LIST":     crew_list,
            "SCHEDULE_TYPE": "shifts",
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
SWEEP_STATE_FILE = Path(".streamlit/sweep_state.json")
_sched      = None
_sched_lock = threading.Lock()

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
    SWEEP_STATE_FILE.write_text(json.dumps(s, indent=2))

def _sync_dataframe(token, mailbox, df_file):
    """Run create/update/delete sync for one parsed dataframe. Returns stats dict."""
    cal_map = build_calendar_map(token, mailbox, df_file["DATE"].min(), df_file["DATE"].max())
    created = updated = deleted = skipped = failed = 0
    for _, row in df_file.iterrows():
        eid     = row["EVENT_ID"]
        subject = f"{row['SHOW']} - {row['FUNCTION']}"
        end_dt  = row["END_DATE"] if pd.notna(row["END_DATE"]) else row["DATE"]
        notes   = f"Foreman: {row['FOREMAN']}. {row['NOTES']}".strip()
        fhash   = _field_hash(row)
        try:
            entry = cal_map.get(eid)
            if entry and entry["hash"] == fhash:
                skipped += 1
            elif entry:
                resp = patch_graph_event(token, mailbox, entry["graph_id"],
                                         row["DATE"], pd.Timestamp(end_dt),
                                         row["LOCATION"], notes,
                                         time_str=row["TIME"], event_id=eid, field_hash=fhash)
                if resp.status_code == 200:
                    updated += 1
                elif resp.status_code == 404:
                    resp2 = create_graph_event(token, mailbox, subject,
                                               row["DATE"], pd.Timestamp(end_dt),
                                               row["LOCATION"], notes,
                                               time_str=row["TIME"], event_id=eid, field_hash=fhash)
                    created += 1 if resp2.status_code == 201 else 0
                    failed  += 0 if resp2.status_code == 201 else 1
                else:
                    failed += 1
            else:
                resp = create_graph_event(token, mailbox, subject,
                                          row["DATE"], pd.Timestamp(end_dt),
                                          row["LOCATION"], notes,
                                          time_str=row["TIME"], event_id=eid, field_hash=fhash)
                created += 1 if resp.status_code == 201 else 0
                failed  += 0 if resp.status_code == 201 else 1
        except Exception:
            failed += 1

    excel_ids = set(df_file["EVENT_ID"])
    for eid, entry in cal_map.items():
        if eid not in excel_ids:
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
        mailbox  = _graph_cfg()["shared_mailbox"]
        drive_id, sp_files = get_sharepoint_files(token)
        state    = _load_sweep_state()

        for f in sp_files:
            fname    = f["name"]
            last_mod = f["modified"]
            if state.get(fname, {}).get("last_modified") == last_mod:
                continue  # unchanged — skip
            try:
                raw    = download_sharepoint_file(token, drive_id, f["id"])
                xl     = pd.ExcelFile(io.BytesIO(raw))
                sheet  = LV_SHEET if LV_SHEET in xl.sheet_names else xl.sheet_names[0]
                df_f   = parse_lv_byDay(io.BytesIO(raw), sheet)
                limit  = config.get("test_rows", 0)
                if limit > 0:
                    df_f = df_f.head(limit)
                stats  = _sync_dataframe(token, mailbox, df_f)
                result = (f"Created:{stats['created']} Updated:{stats['updated']} "
                          f"Deleted:{stats['deleted']} Skipped:{stats['skipped']} Failed:{stats['failed']}")
                state[fname] = {"last_modified": last_mod,
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

def _ensure_scheduler():
    global _sched
    with _sched_lock:
        if _sched is None or not _sched.running:
            _sched = BackgroundScheduler(timezone="UTC")
            _sched.add_job(run_sweep, "interval", minutes=15,
                           id="sp_sweep", next_run_time=dt.utcnow())
            _sched.start()

if _graph_secrets_ok():
    _ensure_scheduler()

# ─── SESSION STATE ───────────────────────────────────────────────────
for _k, _v in [("df", None), ("file_name", None), ("notifs", None), ("ai_done", False)]:
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ─── SIDEBAR ─────────────────────────────────────────────────────────
with st.sidebar:
    st.image("GES-logo.webp", use_container_width=True)
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
            for k in ("df", "file_name", "notifs"):
                st.session_state[k] = None
            st.session_state.ai_done = False
            st.rerun()
        st.divider()
        page = st.radio(
            "Navigation",
            ["Dashboard", "Schedule & Push", "Calendar View", "Crew Lookup"],
            label_visibility="collapsed",
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

        with st.expander("Preview (first 10 rows)"):
            p = parsed[["DATE","SHOW","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW"]].head(10).copy()
            p["DATE"] = p["DATE"].dt.strftime("%Y-%m-%d")
            st.dataframe(p, hide_index=True, use_container_width=True)

        if st.button("Load Schedule", type="primary", use_container_width=True):
            st.session_state.df        = parsed
            st.session_state.file_name = file_label
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
        show_cols = [c for c in ["DATE","SHOW","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW"] if c in df.columns]
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
            st.info(f"Pushing to: **{shared_mailbox}**")

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

                created = updated = failed = skipped = deleted = 0
                results = []
                stat = st.empty()

                stat.text("Reading shared calendar from Outlook...")
                cal_map = build_calendar_map(
                    token, shared_mailbox,
                    df["DATE"].min(), df["DATE"].max(),
                )

                bar = st.progress(0)
                for i, (_, row) in enumerate(df_push.iterrows()):
                    eid     = row["EVENT_ID"]
                    subject = f"{row['SHOW']} - {row['FUNCTION']}"
                    end_dt  = row["END_DATE"] if pd.notna(row["END_DATE"]) else row["DATE"]
                    notes   = f"Foreman: {row['FOREMAN']}. {row['NOTES']}".strip()
                    fhash   = _field_hash(row)
                    stat.text(f"Processing {i+1}/{len(df_push)}: {str(row['SHOW'])[:60]}...")
                    try:
                        entry = cal_map.get(eid)
                        if entry and entry["hash"] == fhash:
                            skipped += 1
                            results.append({"ID": eid, "Event": row["SHOW"], "Status": "Skipped (no change)"})
                        elif entry:
                            resp = patch_graph_event(
                                token, shared_mailbox, entry["graph_id"],
                                row["DATE"], pd.Timestamp(end_dt),
                                row["LOCATION"], notes,
                                time_str=row["TIME"], event_id=eid, field_hash=fhash,
                            )
                            if resp.status_code == 200:
                                updated += 1
                                results.append({"ID": eid, "Event": row["SHOW"], "Status": "Updated"})
                            elif resp.status_code == 404:
                                resp2 = create_graph_event(
                                    token, shared_mailbox, subject,
                                    row["DATE"], pd.Timestamp(end_dt),
                                    row["LOCATION"], notes,
                                    time_str=row["TIME"], event_id=eid, field_hash=fhash,
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
                                token, shared_mailbox, subject,
                                row["DATE"], pd.Timestamp(end_dt),
                                row["LOCATION"], notes,
                                time_str=row["TIME"], event_id=eid, field_hash=fhash,
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

                # Delete Outlook events whose ID is no longer in the full Excel
                if not test_mode:
                    excel_ids = set(df["EVENT_ID"])
                    for eid, entry in cal_map.items():
                        if eid not in excel_ids:
                            stat.text(f"Deleting removed event: {eid}...")
                            try:
                                dresp = delete_graph_event(token, shared_mailbox, entry["graph_id"])
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
