# -*- coding: utf-8 -*-
"""
SmartSchedule sync engine: reads the schedule files from SharePoint and keeps the Outlook calendars in step.
No Streamlit here, so it runs on its own:  python sync_engine.py   syncs every 15 minutes (started by startup.sh
when the App Service starts, so auto-sync no longer waits for someone to open the app). app.py uses the same code.

- One SharePoint folder per calendar (admin workbook, Calendars sheet). Every row in a file goes to its
  folder's calendar; FACILITY only picks colors (and warns when a facility belongs to another calendar).
- Rows dated before today are never created or updated. Past events in Outlook are kept as history.
- A shift is identified by DATE + TIME + SHOW + FUNCTION + FACILITY + LOCATION (no ID column), so sorting,
  filtering or inserting rows changes nothing in Outlook. Changing one of those fields replaces the event.
  Rows with the same key (in one file or across files) are held back with a message.
- Shifts removed from Excel are deleted from Outlook (today to +18 months), unless a file can't be read.
- New problems found by auto-sync are emailed to the admins on the admin workbook's Alerts sheet.
"""
import io
import json
import os
import re
import sys
import time
import hashlib
import tomllib
from datetime import timedelta, datetime as dt
from html import escape
from pathlib import Path
from urllib.parse import unquote, urlparse
from zoneinfo import ZoneInfo

import msal
import pandas as pd
import requests
from openpyxl import load_workbook

import schedule_config
from schedule_config import CALENDARS, FACILITIES, CONFIG_WARNINGS, outlook_categories, event_category

_PT = ZoneInfo("America/Los_Angeles")
LV_SHEET      = "By Day"
LV_CREW_COLS  = [1, 2, 3, 4, 5, 6, 7]   # numeric header columns for crew slots
DEFAULT_TZ    = "Pacific Standard Time"
LINK_COL      = "SHAREPOINT LINK"        # optional By Day column; shown on the Outlook event notes
FACILITY_COL  = "FACILITY"
SYNC_MINUTES  = 15
DELETE_DAYS_AHEAD = 548                  # removed shifts are deleted from today through about 18 months ahead

SP_HOSTNAME = "viadcorp.sharepoint.com"
SP_SITE     = "/sites/ErinsTestSandbox"
SP_FOLDER   = "Attachments/ElectricalSchedule"
ROOT        = Path(__file__).resolve().parent

# On App Service the app folder is a read-only package; /home is writable and survives restarts.
DATA_DIR = Path("/home/data") if os.environ.get("WEBSITE_SITE_NAME") else ROOT / ".streamlit"
SWEEP_STATE_FILE  = DATA_DIR / "sweep_state.json"    # last synced version of each file
SWEEP_STATUS_FILE = DATA_DIR / "sweep_status.json"   # last check / log, shown in the app sidebar
SWEEP_CONFIG_FILE = DATA_DIR / "sweep_config.json"   # sweep test mode, set from the app sidebar
ALERTS_SENT_FILE  = DATA_DIR / "alerts_sent.json"    # problems already emailed to the admins
SYNC_FORMAT = 2   # bump when the way events are stored changes: every file is then synced again


def today():
    return pd.Timestamp(dt.now(_PT).date())


# ─── OUTLOOK TRACKING ────────────────────────────────────────────────
# Each app event carries the shift key and a hash of its fields in two hidden Outlook properties (not shown
# in the event). Outlook is the source of truth. Older events kept 'GES_ID:..|GES_HASH:..' in the first body
# line; those are read as 'legacy:<id>' so the new events replace them.
_PROP_SET = "String {6f1c2a7e-9b3d-4c58-a1e4-3d7b9c0e5f21} Name "
KEY_PROP, HASH_PROP = _PROP_SET + "GES_KEY", _PROP_SET + "GES_HASH"
_EXPAND = f"singleValueExtendedProperties($filter=id eq '{KEY_PROP}' or id eq '{HASH_PROP}')"


def _norm(v):
    return " ".join(str(v or "").split()).upper()


def shift_key(date, time, show, function, facility, location):
    """The shift's identity: DATE | TIME | SHOW | FUNCTION | FACILITY | LOCATION (spacing and case ignored)."""
    t_start, t_end = _parse_time_range(time)
    t = f"{t_start} - {t_end}" if t_start else _norm(time)
    return " | ".join([date.strftime("%Y-%m-%d"), t, _norm(show), _norm(function), _norm(facility), _norm(location)])


def _event_categories(row):
    cat = event_category(row.get("CALENDAR"), row.get("FUNCTION"), row.get("FACILITY"))
    return [cat] if cat else []


def _event_subject(row):
    return " - ".join(p for p in (row.get("FACILITY"), row["SHOW"], row["FUNCTION"]) if p)


def _field_hash(row):
    key = "|".join([_event_subject(row), _event_notes(row), ",".join(_event_categories(row))]
                   + [str(row.get(f, "")) for f in ["DATE", "TIME", "LOCATION"]])
    return hashlib.md5(key.encode()).hexdigest()[:8]


def _tracking(e):
    """(key, hash) of an Outlook event: from the hidden properties, else the legacy body line, else ('', '')."""
    props = {p["id"].rsplit(" ", 1)[-1].upper(): p.get("value", "") for p in e.get("singleValueExtendedProperties") or []}
    if props.get("GES_KEY"):
        return props["GES_KEY"], props.get("GES_HASH", "")
    eid, _ = _parse_body((e.get("body") or {}).get("content", ""))
    return (f"legacy:{eid}", "") if eid else ("", "")


_PRESET_NAMES = {v: k for k, v in schedule_config.COLORS.items() if k != "grey"}


def ensure_outlook_categories(token, mailbox):
    """Create the color categories used by the calendars on this mailbox. Returns '' or an error message."""
    wanted = {}
    for key, c in CALENDARS.items():
        if c["mailbox"] == mailbox:
            wanted.update(outlook_categories(key))
    if not wanted:
        return ""
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/outlook/masterCategories"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, timeout=15)
    if resp.status_code != 200:
        return (f"could not read Outlook categories in {mailbox} ({resp.status_code}). Events are still tagged; "
                f"to show colors, grant the app MailboxSettings.ReadWrite or create the categories by hand: "
                + ", ".join(f"{n} ({_PRESET_NAMES.get(p, p)})" for n, p in wanted.items()))
    have = {c["displayName"].lower() for c in resp.json().get("value", [])}
    for name, preset in wanted.items():
        if name.lower() not in have:
            r = requests.post(url, headers=headers, json={"displayName": name, "color": preset}, timeout=15)
            if r.status_code not in (200, 201):
                return f"could not create Outlook category {name} in {mailbox} ({r.status_code})"
    return ""


def _event_notes(row):
    """Outlook notes text: foreman, the NOTES column (all its lines), and each SharePoint link on its own line."""
    notes = f"Foreman: {row['FOREMAN']}. {row['NOTES']}".strip()
    for label, url in row.get("LINKS") or []:
        notes += f"\nSharePoint: [{label}]({url})" if url else f"\nSharePoint: {label}"  # link made by _make_body
    return notes


def _make_body(notes):
    """HTML event body from the notes; '[label](url)' becomes a clickable link."""
    html = escape(notes.strip())
    html = re.sub(r"\[([^\]\n]+)\]\((\S+?)\)", r'<a href="\2">\1</a>', html)
    return html.replace("\n", "<br>")


def _parse_body(body_content):
    """Legacy events: (event_id, field_hash) from the 'GES_ID:..|GES_HASH:..' body line, or ('', '')."""
    for line in (body_content or "").splitlines():
        line = line.strip()
        if line.startswith("GES_ID:") and "|GES_HASH:" in line:
            parts = line.split("|GES_HASH:")
            return parts[0].replace("GES_ID:", "").strip(), parts[1].strip() if len(parts) > 1 else ""
    return "", ""


def build_calendar_map(token, mailbox, start_date, end_date):
    """Query Outlook for the date range and return {shift key: {graph_id, hash}} for the app's events.
    Raises if Outlook can't be read: a partial map would make the app create duplicates or skip deletes."""
    url    = f"https://graph.microsoft.com/v1.0/users/{mailbox}/calendarView"
    params = {
        "startDateTime": start_date.strftime("%Y-%m-%dT00:00:00Z"),
        # +2 days: an evening shift on the last day starts after midnight UTC (US time zones are behind UTC)
        "endDateTime":   (end_date + timedelta(days=2)).strftime("%Y-%m-%dT00:00:00Z"),
        "$select": "id,body",
        "$expand": _EXPAND,
        "$top": "500",
    }
    headers = {"Authorization": f"Bearer {token}", "Prefer": "outlook.body-content-type='text'"}
    mapping, next_url, next_params = {}, url, params
    while next_url:
        resp = requests.get(next_url, headers=headers, params=next_params, timeout=15)
        if resp.status_code != 200:
            raise RuntimeError(f"could not read the Outlook calendar of {mailbox} ({resp.status_code})")
        data = resp.json()
        for e in data.get("value", []):
            eid, fhash = _tracking(e)
            if eid:
                mapping[eid] = {"graph_id": e["id"], "hash": fhash}
        next_url, next_params = data.get("@odata.nextLink"), None
    return mapping


def read_calendar_events(token, mailbox, start_date, end_date, time_zone):
    """Read-only: every Outlook event from start_date through end_date, times in `time_zone`.
    Returns (events, error). Each event: SUBJECT, START, END (end is exclusive), ALL_DAY, LOCATION,
    CATEGORIES, EVENT_ID (shift key if the app made it, else ''), NOTES (body text without a legacy ID line),
    MODIFIED (UTC)."""
    params = {
        "startDateTime": (start_date - timedelta(days=1)).strftime("%Y-%m-%dT00:00:00Z"),  # pad for time zones
        "endDateTime":   (end_date + timedelta(days=2)).strftime("%Y-%m-%dT00:00:00Z"),
        "$select": "subject,start,end,isAllDay,location,categories,body,lastModifiedDateTime",
        "$expand": _EXPAND,
        "$orderby": "start/dateTime",
        "$top": "500",
    }
    headers = {"Authorization": f"Bearer {token}",
               "Prefer": f"outlook.timezone=\"{time_zone}\", outlook.body-content-type=\"text\""}
    events, next_url = [], f"https://graph.microsoft.com/v1.0/users/{mailbox}/calendarView"
    while next_url:
        resp = requests.get(next_url, headers=headers, params=params, timeout=15)
        if resp.status_code != 200:
            return events, f"could not read the Outlook calendar of {mailbox} ({resp.status_code})"
        data = resp.json()
        for e in data.get("value", []):
            body = (e.get("body") or {}).get("content", "")
            eid, _ = _tracking(e)
            events.append({
                "SUBJECT":    e.get("subject") or "(no title)",
                "START":      pd.Timestamp(e["start"]["dateTime"][:19]),
                "END":        pd.Timestamp(e["end"]["dateTime"][:19]),
                "ALL_DAY":    bool(e.get("isAllDay")),
                "LOCATION":   (e.get("location") or {}).get("displayName", ""),
                "CATEGORIES": e.get("categories") or [],
                "EVENT_ID":   eid,
                "NOTES":      "\n".join(l for l in body.splitlines() if not l.strip().startswith("GES_ID:")).strip(),
                "MODIFIED":   pd.Timestamp(e["lastModifiedDateTime"]) if e.get("lastModifiedDateTime") else None,
            })
        next_url, params = data.get("@odata.nextLink"), None
    return events, ""


def events_on_day(events, day):
    """Events that show on `day` (an event spanning several days shows on each), all-day first, then by time."""
    day = pd.Timestamp(day).normalize()
    hits = [e for e in events
            if e["START"] < day + timedelta(days=1) and (e["END"] > day or e["START"] == e["END"] == day)]
    return sorted(hits, key=lambda e: (not e["ALL_DAY"], e["START"]))


# ─── GRAPH HELPERS ───────────────────────────────────────────────────
def _graph_cfg():
    """Graph credentials from env vars (App Service) or .streamlit/secrets.toml [graph] (local)."""
    env = {"app_id":        os.environ.get("GRAPH_APP_ID"),
           "tenant_id":     os.environ.get("GRAPH_TENANT_ID"),
           "client_secret": os.environ.get("GRAPH_CLIENT_SECRET")}
    if all(env.values()):
        return env
    try:
        return dict(tomllib.loads((ROOT / ".streamlit" / "secrets.toml").read_text(encoding="utf8"))["graph"])
    except Exception:
        return env


def graph_secrets_ok():
    return bool(_graph_cfg().get("client_secret"))


def get_graph_token():
    g = _graph_cfg()
    app = msal.ConfidentialClientApplication(
        client_id=g["app_id"], client_credential=g["client_secret"],
        authority=f"https://login.microsoftonline.com/{g['tenant_id']}")
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    raise RuntimeError(result.get("error_description", "Graph auth failed"))


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


def _event_payload(row, event_id, field_hash):
    """Graph event JSON for one row (timed when TIME is 'HH:MM - HH:MM', otherwise all-day)."""
    start_dt = row["DATE"]
    end_dt   = pd.Timestamp(row["END_DATE"] if pd.notna(row["END_DATE"]) else row["DATE"])
    t_start, t_end = _parse_time_range(row["TIME"])
    tz = row.get("TIME_ZONE") or DEFAULT_TZ
    payload = {"subject": _event_subject(row),
               "location": {"displayName": row["LOCATION"] or "Not specified"},
               "body": {"contentType": "HTML", "content": _make_body(_event_notes(row))},
               "categories": _event_categories(row),
               "singleValueExtendedProperties": [{"id": KEY_PROP, "value": event_id},
                                                 {"id": HASH_PROP, "value": field_hash}]}
    if t_start and t_end:
        payload.update(isAllDay=False,
                       start={"dateTime": start_dt.strftime(f"%Y-%m-%dT{t_start}:00"), "timeZone": tz},
                       end={"dateTime": _timed_end(end_dt, t_start, t_end), "timeZone": tz})
    else:
        payload.update(isAllDay=True,
                       start={"dateTime": start_dt.strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"},
                       end={"dateTime": (end_dt + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"})
    return payload


def create_graph_event(token, mailbox, payload):
    return requests.post(f"https://graph.microsoft.com/v1.0/users/{mailbox}/events",
                         headers={"Authorization": f"Bearer {token}"}, json=payload, timeout=15)


def patch_graph_event(token, mailbox, graph_id, payload):
    return requests.patch(f"https://graph.microsoft.com/v1.0/users/{mailbox}/events/{graph_id}",
                          headers={"Authorization": f"Bearer {token}"}, json=payload, timeout=15)


def delete_graph_event(token, mailbox, graph_id):
    return requests.delete(f"https://graph.microsoft.com/v1.0/users/{mailbox}/events/{graph_id}",
                           headers={"Authorization": f"Bearer {token}"}, timeout=10)


# ─── SHAREPOINT ──────────────────────────────────────────────────────
def _sp_headers(token):
    return {"Authorization": f"Bearer {token}"}


def _drive_id(token):
    site = requests.get(f"https://graph.microsoft.com/v1.0/sites/{SP_HOSTNAME}:{SP_SITE}",
                        headers=_sp_headers(token), timeout=10)
    site.raise_for_status()
    drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site.json()['id']}/drive",
                         headers=_sp_headers(token), timeout=10)
    drive.raise_for_status()
    return drive.json()["id"]


def load_admin(token, drive_id=None):
    """Download the admin workbook and load it into schedule_config. Returns its last-modified time ('' if not read)."""
    drive_id = drive_id or _drive_id(token)
    path = f"{SP_FOLDER}/{schedule_config.ADMIN_FOLDER}/{schedule_config.ADMIN_FILE}"
    try:
        item = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{path}",
                            headers=_sp_headers(token), timeout=10)
        item.raise_for_status()
        data = download_sharepoint_file(token, drive_id, item.json()["id"])
    except Exception as e:
        CONFIG_WARNINGS[:] = [f"Could not download the admin workbook {path} ({e}); using the last settings that loaded."]
        return ""
    schedule_config.apply(io.BytesIO(data))
    return item.json().get("lastModifiedDateTime", "")


def get_sharepoint_files(token, drive_id=None):
    """Schedule files (.xlsx, name without 'TEMPLATE') in every calendar's folder.
    Returns (drive_id, files, folder_errors); each file has name, id, path, calendar, modified."""
    drive_id = drive_id or _drive_id(token)
    files, folder_errors = [], {}
    for key, c in CALENDARS.items():
        resp = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SP_FOLDER}/{c['folder']}:/children",
                            headers=_sp_headers(token),
                            params={"$select": "id,name,lastModifiedDateTime,file", "$top": "200"}, timeout=10)
        if resp.status_code != 200:
            folder_errors[c["folder"]] = f"folder for calendar {key} could not be read ({resp.status_code})"
            continue
        for i in resp.json().get("value", []):
            if i.get("name", "").lower().endswith(".xlsx") and "file" in i and "TEMPLATE" not in i["name"].upper():
                files.append({"name": i["name"], "id": i["id"], "path": f"{c['folder']}/{i['name']}", "calendar": key,
                              "modified": i.get("lastModifiedDateTime", ""),
                              "modified_display": dt.fromisoformat(i["lastModifiedDateTime"].replace("Z", "+00:00"))
                                                    .astimezone(_PT).strftime("%Y-%m-%d %H:%M PT")})
    return drive_id, files, folder_errors


def download_sharepoint_file(token, drive_id, item_id):
    resp = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
                        headers=_sp_headers(token), timeout=30, allow_redirects=True)
    resp.raise_for_status()
    return resp.content


# ─── LV FORMAT IMPORTER ──────────────────────────────────────────────
def _cell_text(val):
    """Cell value as stripped text; empty cells (NaN/None) become '' instead of 'nan'."""
    s = str(val if val is not None else "").strip()
    return "" if s.lower() in ("nan", "none") else s


def _is_url(s):
    return bool(re.match(r"https?://\S+$", s, re.I))


def _url_label(url):
    """A short name for a link: the file or page name at the end of the URL."""
    name = unquote(urlparse(url).path.rstrip("/").rsplit("/", 1)[-1])
    return name or url


def _links(text, url):
    """[(label, url)] for a SHAREPOINT LINK cell: the hyperlink behind the cell (labelled by the cell's first
    line that isn't a URL), then each URL on its own line. Other text lines are kept as (text, '')."""
    lines = [l.strip() for l in str(text).splitlines() if l.strip()]
    out = []
    if url:
        label = next((l for l in lines if not _is_url(l)), "") or _url_label(url)
        out.append((label, url))
        if label in lines:
            lines.remove(label)
    for l in lines:
        if l != url:
            out.append((_url_label(l), l) if _is_url(l) else (l, ""))
    return out


def _read_link_urls(src, sheet_name):
    """{Excel row number: URL} for SHAREPOINT LINK cells that are hyperlinks. pandas only sees the text Excel
    displays (often a file name), so the URL behind the cell is read with openpyxl."""
    if hasattr(src, "seek"):
        src.seek(0)
    ws = load_workbook(src)[sheet_name]
    header = [str(c.value).strip().upper() if c.value is not None else "" for c in ws[1]]
    if LINK_COL not in header:
        return {}
    li = header.index(LINK_COL)
    return {row[li].row: row[li].hyperlink.target for row in ws.iter_rows(min_row=2)
            if row[li].hyperlink is not None and row[li].hyperlink.target}


def parse_lv_byDay(src, sheet_name, calendar):
    """Parse the 'By Day' sheet of the GES Electrical Show Schedule template; every row goes to `calendar`.

    Columns: DAY, DATE, TIME, SHOW, FUNCTION, FACILITY, LOCATION, NOTES, FOREMAN, SHAREPOINT LINK, 1..7.
    Crew columns 1-7 hold names or '+N' (extra unnamed crew). An old ID column is ignored.
    Each row's EVENT_ID is its shift key (see shift_key).
    """
    xl = pd.ExcelFile(src)
    # Read crew columns as text, otherwise a column holding only '+1' comes back as 1.0.
    # In an Excel Table the crew headers are text ('1'), elsewhere numbers (1).
    raw = xl.parse(sheet_name, header=0, dtype={c: str for c in LV_CREW_COLS + [str(c) for c in LV_CREW_COLS]})
    raw.columns = [int(c) if str(c).strip().isdigit() else c for c in raw.columns]
    link_urls = _read_link_urls(src, sheet_name) if LINK_COL in raw.columns else {}
    time_zone = (CALENDARS.get(calendar) or {}).get("time_zone") or DEFAULT_TZ

    required = {"DATE", "SHOW", "FUNCTION", "LOCATION", "NOTES", "FOREMAN"}
    missing  = required - {str(c).strip().upper() for c in raw.columns}
    if missing:
        raise ValueError(f"Sheet '{sheet_name}' is missing expected columns: {', '.join(sorted(missing))}")

    crew_cols = [c for c in raw.columns if isinstance(c, int)]
    rows = []
    for i, r in raw.iterrows():
        date = pd.to_datetime(r.get("DATE"), errors="coerce")
        if pd.isna(date) or date.year < 2020:
            continue
        show = _cell_text(r.get("SHOW"))
        if not show:
            continue
        foreman = _cell_text(r.get("FOREMAN")) or "TBD"
        crew_list = [foreman] if foreman != "TBD" else []
        extra = 0
        for cc in crew_cols:
            val = _cell_text(r.get(cc))
            if not val:
                continue
            if val.startswith("+"):
                try:
                    extra += int(val[1:])
                except ValueError:
                    extra += 1
            else:
                crew_list.append(val)

        fn = _cell_text(r.get("FUNCTION")).upper() or "SHOW"
        time, location = _cell_text(r.get("TIME")), _cell_text(r.get("LOCATION"))
        facility = _cell_text(r.get(FACILITY_COL)).upper()
        rows.append({
            "EVENT_ID":      shift_key(date, time, show, fn, facility, location),
            "DATE":          date,
            "END_DATE":      pd.NaT,
            "SHOW":          show,
            "FUNCTION":      fn,
            "LOCATION":      location,
            "TIME":          time,
            "FOREMAN":       foreman,
            "NOTES":         _cell_text(r.get("NOTES")),
            "TOTAL_CREW":    max(len(crew_list) + extra, 1),
            "CREW_LIST":     crew_list,
            "SCHEDULE_TYPE": "shifts",
            "TIME_ZONE":     time_zone,
            # [(label, url)]: the hyperlink behind the cell (matched by Excel row) and URLs typed one per line
            "LINKS":         _links(_cell_text(r.get(LINK_COL)), link_urls.get(i + 2, "")),
            "FACILITY":      facility,
            "CALENDAR":      calendar,
            "EXTRA_CREW":    extra,  # unnamed '+N' crew, used by the Overtime page
        })
    if not rows:
        raise ValueError("No valid shifts found in the selected sheet.")
    return pd.DataFrame(rows).reset_index(drop=True)


def load_all(token):
    """Admin settings, then every schedule file in every calendar folder.
    Returns (files, {path: df}, {path or folder: error}, admin_modified). Each df has a FILE column."""
    drive_id = _drive_id(token)
    admin_modified = load_admin(token, drive_id)
    _, files, errors = get_sharepoint_files(token, drive_id)
    parsed = {}
    for f in files:
        try:
            raw = download_sharepoint_file(token, drive_id, f["id"])
            if LV_SHEET not in pd.ExcelFile(io.BytesIO(raw)).sheet_names:
                continue  # not a schedule file
            parsed[f["path"]] = parse_lv_byDay(io.BytesIO(raw), LV_SHEET, f["calendar"]).assign(FILE=f["path"])
        except Exception as e:
            errors[f["path"]] = str(e)
    return files, parsed, errors, admin_modified


# ─── ROUTING ─────────────────────────────────────────────────────────
def duplicate_ids(dfs):
    """Shift keys used by more than one row (in one file or across files)."""
    ids = pd.concat([d["EVENT_ID"] for d in dfs]) if dfs else pd.Series(dtype=object)
    return set(ids[ids.duplicated()])


def route_rows(df, dup_ids=()):
    """(mailbox, error) per row: the row's calendar mailbox, or None with the reason it is not pushed."""
    def route(r):
        if r["EVENT_ID"] in dup_ids:
            return None, ("another row has the same DATE, TIME, SHOW, FUNCTION, FACILITY and LOCATION; "
                          "remove the extra row or make them differ")
        c = CALENDARS.get(r["CALENDAR"])
        if not c:
            return None, f"calendar {r['CALENDAR']} is not on the admin Calendars sheet"
        if not c["mailbox"]:
            return None, f"calendar {r['CALENDAR']} has no MAILBOX on the admin Calendars sheet"
        return c["mailbox"], ""
    routes = [route(r) for _, r in df.iterrows()]
    return (pd.Series([m for m, _ in routes], index=df.index, dtype=object),
            pd.Series([e for _, e in routes], index=df.index, dtype=object))


def facility_warnings(df):
    """Per row: '' or a note about the FACILITY (unknown code, or a facility of another calendar).
    These rows are still pushed to their folder's calendar."""
    def warn(r):
        code = r["FACILITY"]
        if not code:
            return ""
        f = FACILITIES.get(code)
        if f is None:
            return f"'{code}' is not on the admin Facilities sheet"
        if f["calendar"] and f["calendar"] != r["CALENDAR"]:
            return f"{code} belongs to the {f['calendar']} calendar, but the file is in the {r['CALENDAR']} folder"
        return ""
    return pd.Series([warn(r) for _, r in df.iterrows()], index=df.index, dtype=object)


# ─── SYNC ────────────────────────────────────────────────────────────
def _all_calendar_mailboxes():
    return {c["mailbox"] for c in CALENDARS.values() if c["mailbox"]}


def _delete_removed(token, keep_ids, mailboxes):
    """Delete app events, from today on, in every calendar, whose ID is in no schedule file, or whose row now
    routes to a different calendar. keep_ids: {EVENT_ID: mailbox it routes to, or None}.
    One mailbox can be listed under two addresses; it returns the same event id under both, so an event seen
    under its row's address is not 'moved'. Returns [(EVENT_ID, mailbox, deleted_ok)]."""
    start = today()
    end   = start + timedelta(days=DELETE_DAYS_AHEAD)
    maps  = {m: build_calendar_map(token, m, start, end) for m in sorted(mailboxes)}
    done, gone = [], set()
    for mbx, cal_map in maps.items():
        for eid, entry in cal_map.items():
            if eid in keep_ids:
                target = keep_ids[eid]
                if target in (None, mbx):
                    continue
                if maps.get(target, {}).get(eid, {}).get("graph_id") == entry["graph_id"]:
                    continue  # same mailbox under its other address
            if entry["graph_id"] in gone:
                continue      # already deleted through the mailbox's other address
            try:
                ok = delete_graph_event(token, mbx, entry["graph_id"]).status_code in (200, 204)
            except Exception:
                ok = False
            if ok:
                gone.add(entry["graph_id"])
            done.append((eid, mbx, ok))
    return done


def remove_deleted_shifts(token, parsed, errors):
    """Delete Outlook events (today on) whose shift is no longer in Excel. Returns (results, problem message).
    Skipped when a file or folder can't be read (its events would look removed). A calendar whose folder has
    no schedule files is skipped too (protection against a SharePoint glitch), with every calendar on its mailbox."""
    if errors:
        return [], ("Deleted shifts were NOT removed from Outlook — could not read: " + ", ".join(errors)
                    + ". Fix or remove that file.")
    have = {d["CALENDAR"].iloc[0] for d in parsed.values()}
    unsafe = {c["mailbox"] for k, c in CALENDARS.items() if k not in have}
    mailboxes = _all_calendar_mailboxes() - unsafe
    dups = duplicate_ids(list(parsed.values()))
    keep_ids = {}
    for d in parsed.values():
        future = d[d["DATE"] >= today()]
        keep_ids.update(zip(future["EVENT_ID"], route_rows(future, dups)[0]))
    return _delete_removed(token, keep_ids, mailboxes), ""


def sync_rows(token, df, dup_ids=()):
    """Create / update Outlook events for the rows in df dated today or later. Past rows are left alone.
    Returns (stats, results, not_pushed) - results: [{ID, Show, Status}]; not_pushed: [(ID, reason)]."""
    stats = {"created": 0, "updated": 0, "skipped": 0, "failed": 0, "past": int((df["DATE"] < today()).sum())}
    results, not_pushed = [], []
    df = df[df["DATE"] >= today()]
    if df.empty:
        return stats, results, not_pushed
    mbx_s, err_s = route_rows(df, dup_ids)
    ok = err_s == ""
    not_pushed = list(zip(df.loc[~ok, "EVENT_ID"], err_s[~ok]))
    for mbx, grp in df[ok].groupby(mbx_s[ok]):
        cal_map = build_calendar_map(token, mbx, grp["DATE"].min(), grp["DATE"].max())
        for _, row in grp.iterrows():
            eid, fhash = row["EVENT_ID"], _field_hash(row)
            entry, status = cal_map.get(eid), None
            try:
                if entry and entry["hash"] == fhash:
                    stats["skipped"] += 1
                    continue
                payload = _event_payload(row, eid, fhash)
                if entry:
                    resp = patch_graph_event(token, mbx, entry["graph_id"], payload)
                    if resp.status_code == 200:
                        stats["updated"] += 1
                        status = "Updated"
                if status is None and (not entry or resp.status_code == 404):  # new, or deleted by hand
                    resp = create_graph_event(token, mbx, payload)
                    if resp.status_code == 201:
                        stats["created"] += 1
                        status = "Created"
                if status is None:
                    stats["failed"] += 1
                    status = "Failed: " + resp.json().get("error", {}).get("message", str(resp.status_code))[:80]
            except Exception as e:
                stats["failed"] += 1
                status = f"Error: {e}"
            results.append({"Shift": eid, "Show": row["SHOW"], "Status": status})
    return stats, results, not_pushed


# ─── AUTO-SYNC (every 15 minutes) ────────────────────────────────────
def _read_json(path, default):
    try:
        return json.loads(path.read_text()) if path.exists() else default
    except Exception:
        return default


def _write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2))
    tmp.replace(path)


def sweep_status():
    return _read_json(SWEEP_STATUS_FILE, {"last_check": None, "running": False, "log": [], "error": None})


def sweep_test_rows():
    return int(_read_json(SWEEP_CONFIG_FILE, {}).get("test_rows", 0))


def set_sweep_test_rows(n):
    _write_json(SWEEP_CONFIG_FILE, {"test_rows": int(n)})


def send_admin_email(token, problems):
    """Email problems to the admins on the admin Alerts sheet. Returns '' or why it was not sent."""
    alerts = schedule_config.ALERTS
    if not alerts["to"] or not alerts["sender"]:
        return "the admin workbook's Alerts sheet needs an ADMIN EMAIL and a SEND FROM mailbox"
    items = "".join(f"<li>{escape(p)}</li>" for p in problems)
    html = (f"<p>SmartSchedule auto-sync found {len(problems)} new problem(s):</p><ul>{items}</ul>"
            f"<p>Fix them in the Excel file or the admin workbook. Auto-sync tries again every {SYNC_MINUTES} "
            f"minutes. You get one email per new problem; it is not repeated while the problem stays.</p>")
    msg = {"subject": f"SmartSchedule: {len(problems)} new problem(s)", "body": {"contentType": "HTML", "content": html},
           "toRecipients": [{"emailAddress": {"address": a}} for a in alerts["to"]]}
    try:
        r = requests.post(f"https://graph.microsoft.com/v1.0/users/{alerts['sender']}/sendMail",
                          headers={"Authorization": f"Bearer {token}"},
                          json={"message": msg, "saveToSentItems": False}, timeout=15)
    except Exception as e:
        return str(e)
    if r.status_code == 202:
        return ""
    return f"Outlook said {r.status_code}" + (" - the app needs the Mail.Send permission" if r.status_code == 403 else "")


def alert_admins(token, problems):
    """Email the problems not emailed before. Fixed problems are forgotten, so they are emailed again if they
    come back. Returns a log line ('' when there is nothing new)."""
    current = sorted(set(problems))
    new = [p for p in current if p not in set(_read_json(ALERTS_SENT_FILE, []))]
    if new and (err := send_admin_email(token, new)):
        return f"⚠️ Admin email not sent: {err}"   # not marked as sent, so it is tried again next sweep
    _write_json(ALERTS_SENT_FILE, current)
    return f"📧 Emailed {len(new)} new problem(s) to the admins" if new else ""


def run_sweep():
    """One auto-sync: push changed files, remove deleted shifts, email new problems to the admins.
    Writes the result to SWEEP_STATUS_FILE."""
    status = {**sweep_status(), "running": True, "error": None}
    _write_json(SWEEP_STATUS_FILE, status)
    log, problems, token = [], [], None
    try:
        token = get_graph_token()
        files, parsed, errors, admin_modified = load_all(token)
        state = _read_json(SWEEP_STATE_FILE, {})
        if state.get("_admin_modified") != admin_modified or state.get("_format") != SYNC_FORMAT:
            # settings changed, or a new app version stores events differently: re-check every file
            state = {"_admin_modified": admin_modified, "_format": SYNC_FORMAT}
        limit = sweep_test_rows()
        dups  = duplicate_ids(list(parsed.values()))
        log += [f"⚠️ {w}" for w in CONFIG_WARNINGS]
        problems += [f"Admin workbook: {w}" for w in CONFIG_WARNINGS]
        for m in sorted(_all_calendar_mailboxes()):
            if (cat_err := ensure_outlook_categories(token, m)):
                log.append(f"⚠️ Event colors: {cat_err}")
                problems.append(f"Event colors: {cat_err}")
        for path, df_p in parsed.items():
            fut = df_p[df_p["DATE"] >= today()]
            notes = facility_warnings(fut)
            problems += [f"{path}: {k} - {n} (still pushed)" for k, n in zip(fut["EVENT_ID"], notes) if n]
        for f in files:
            path, last_mod = f["path"], f["modified"]
            if path in errors:
                log.append(f"❌ {path} — Error: {errors[path]}")
                problems.append(f"{path}: could not be read - {errors[path]}")
                continue
            if path not in parsed or state.get(path, {}).get("last_modified") == last_mod:
                continue  # not a schedule file, or unchanged
            df_f = parsed[path]
            if limit > 0:
                df_f = df_f[df_f["DATE"] >= today()].head(limit)
            stats, results, bad = sync_rows(token, df_f, dups)
            result = (f"Created:{stats['created']} Updated:{stats['updated']} Skipped:{stats['skipped']} "
                      f"Failed:{stats['failed']} Past (not pushed):{stats['past']}")
            if bad:
                result += f" Not pushed:{len(bad)} ({'; '.join(f'{e}: {m}' for e, m in bad[:3])}{'; ...' if len(bad) > 3 else ''})"
            problems += [f"{path}: {k} - not pushed: {m}" for k, m in bad]
            problems += [f"{path}: {r['Shift']} - {r['Status']}" for r in results
                         if r["Status"] not in ("Created", "Updated")]
            # Files with rows not pushed are retried next sweep (e.g. once the duplicate row is fixed)
            state[path] = {"last_modified": None if bad or stats["failed"] else last_mod,
                           "last_synced": dt.now(_PT).strftime("%Y-%m-%d %H:%M PT"), "result": result}
            log.append(f"✅ {path} — {result}")
        for folder, err in errors.items():
            if folder not in {f["path"] for f in files}:
                log.append(f"❌ {folder} — {err}")
                problems.append(f"{folder}: {err}")
        _write_json(SWEEP_STATE_FILE, state)

        removed, problem = remove_deleted_shifts(token, parsed, errors)
        if problem:
            status["error"] = problem
            problems.append(problem)
        ok_ids  = [e for e, _, ok in removed if ok]
        bad_ids = [e for e, _, ok in removed if not ok]
        if ok_ids:
            log.append(f"🗑️ Removed {len(ok_ids)} event(s) no longer in Excel: "
                       f"{', '.join(ok_ids[:10])}{' ...' if len(ok_ids) > 10 else ''}")
        if bad_ids:
            status["error"] = f"Could not delete {len(bad_ids)} removed event(s): {', '.join(bad_ids[:10])}"
            problems += [f"Could not delete the Outlook event of removed shift {e} ({m})" for e, m, ok in removed if not ok]
    except Exception as e:
        status["error"] = str(e)
        log.append(f"❌ Sweep error: {e}")
        problems.append(f"Auto-sync stopped with an error: {e}")
    finally:
        if token:
            try:
                if (line := alert_admins(token, problems)):
                    log.append(line)
            except Exception as e:
                log.append(f"⚠️ Admin email not sent: {e}")
        status.update(last_check=dt.now(_PT).strftime("%Y-%m-%d %H:%M PT"), running=False, log=log)
        _write_json(SWEEP_STATUS_FILE, status)


if __name__ == "__main__":
    # Background worker: started by startup.sh next to the Streamlit app.
    if not graph_secrets_ok():
        sys.exit("Graph credentials are not configured; auto-sync not started.")
    while True:
        started = time.time()
        run_sweep()
        print(f"{dt.now(_PT):%Y-%m-%d %H:%M} sweep done: {sweep_status().get('error') or 'ok'}", flush=True)
        time.sleep(max(60, SYNC_MINUTES * 60 - (time.time() - started)))
