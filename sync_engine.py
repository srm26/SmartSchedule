# -*- coding: utf-8 -*-
"""
SmartSchedule sync engine: reads the schedule files from SharePoint and keeps the Outlook calendars in step.
No Streamlit here, so it runs on its own:  python sync_engine.py   syncs every 15 minutes (started by startup.sh
when the App Service starts, so auto-sync no longer waits for someone to open the app). app.py uses the same code.

- One SharePoint folder per calendar (admin workbook, Calendars sheet). Every row in a file goes to its
  folder's calendar; FACILITY only picks colors (and warns when a facility belongs to another calendar).
- Rows dated before today are never created or updated. Past events in Outlook are kept as history.
- Duplicate IDs (in one file or across files) are held back with a message.
- Shifts removed from Excel are deleted from Outlook (today to +18 months), unless a file can't be read.
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
LISTS_SHEET   = "Lists"                  # holds the file's ID PREFIX
ID_PREFIX_COL = "ID PREFIX"              # Lists column: prefix used by the ID formula (e.g. LV -> LV-0001)
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


def today():
    return pd.Timestamp(dt.now(_PT).date())


# ─── OUTLOOK BODY MARKERS ────────────────────────────────────────────
# EVENT_ID and field hash are stored in the first line of the event body. Outlook is the source of truth.
def _event_categories(row):
    cat = event_category(row.get("CALENDAR"), row.get("FUNCTION"), row.get("FACILITY"))
    return [cat] if cat else []


def _field_hash(row):
    key = "|".join(str(row.get(f, "")) for f in ["DATE", "SHOW", "FUNCTION", "TIME", "LOCATION", "FOREMAN", "NOTES"])
    if row.get("LINK"):  # only when present, so rows without a link keep their existing hash
        key += "|" + str(row["LINK"])
    if _event_categories(row):  # same idea: only colored rows get a new hash
        key += "|cat:" + ",".join(_event_categories(row))
    return hashlib.md5(key.encode()).hexdigest()[:8]


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
            eid, fhash = _parse_body(e.get("body", {}).get("content", ""))
            if eid:
                mapping[eid] = {"graph_id": e["id"], "hash": fhash}
        next_url, next_params = data.get("@odata.nextLink"), None
    return mapping


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


def _make_event_id(show, function, date):
    slug = re.sub(r"[^A-Za-z0-9]", "", str(show))[:20].upper()
    return f"{slug}-{function}-{date.strftime('%Y%m%d')}"


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
    payload = {"subject": f"{row['SHOW']} - {row['FUNCTION']}",
               "location": {"displayName": row["LOCATION"] or "Not specified"},
               "body": {"contentType": "HTML", "content": _make_body(_event_notes(row), event_id, field_hash)},
               "categories": _event_categories(row)}
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


def _read_id_prefix(xl):
    """The file's ID PREFIX (Lists sheet), used by the ID formula; '' if absent."""
    if LISTS_SHEET not in xl.sheet_names:
        return ""
    lists = xl.parse(LISTS_SHEET, header=0, dtype=str)
    lists.columns = [str(c).strip().upper() for c in lists.columns]
    if ID_PREFIX_COL not in lists.columns:
        return ""
    vals = [_cell_text(v) for v in lists[ID_PREFIX_COL]]
    return next((v for v in vals if v), "")


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

    Columns: ID, DAY, DATE, TIME, SHOW, FUNCTION, FACILITY, LOCATION, NOTES, FOREMAN, SHAREPOINT LINK, 1..7.
    Crew columns 1-7 hold names or '+N' (extra unnamed crew).
    ID is a formula: <ID PREFIX>-<row number - 1>. When the cell has no saved value (a file saved by a script),
    the same ID is worked out here, so the ID never depends on how the file was saved.
    """
    xl = pd.ExcelFile(src)
    # Read crew columns as text, otherwise a column holding only '+1' comes back as 1.0
    raw = xl.parse(sheet_name, header=0, dtype={c: str for c in LV_CREW_COLS})
    link_urls = _read_link_urls(src, sheet_name) if LINK_COL in raw.columns else {}
    prefix = _read_id_prefix(xl)
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
        id_val = r.get("ID", "")
        if isinstance(id_val, float) and id_val.is_integer():
            id_val = int(id_val)  # Excel numbers come in as 1.0 — store as "1"
        raw_id = _cell_text(id_val)
        if not raw_id and prefix:
            raw_id = f"{prefix}-{i + 1:04d}"  # same as the ID formula: row number - 1 (row i + 2 in Excel)
        rows.append({
            "EVENT_ID":      raw_id or _make_event_id(show, fn, date),
            "DATE":          date,
            "END_DATE":      pd.NaT,
            "SHOW":          show,
            "FUNCTION":      fn,
            "LOCATION":      _cell_text(r.get("LOCATION")),
            "TIME":          _cell_text(r.get("TIME")),
            "FOREMAN":       foreman,
            "NOTES":         _cell_text(r.get("NOTES")),
            "TOTAL_CREW":    max(len(crew_list) + extra, 1),
            "CREW_LIST":     crew_list,
            "SCHEDULE_TYPE": "shifts",
            "TIME_ZONE":     time_zone,
            # LINK = the URL (from the hyperlink behind the cell if there is one); LINK_TEXT = what the cell shows
            "LINK":          link_urls.get(i + 2) or _cell_text(r.get(LINK_COL)),
            "LINK_TEXT":     _cell_text(r.get(LINK_COL)),
            "FACILITY":      _cell_text(r.get(FACILITY_COL)).upper(),
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
    """IDs used by more than one row (in one file or across files)."""
    ids = pd.concat([d["EVENT_ID"] for d in dfs]) if dfs else pd.Series(dtype=object)
    return set(ids[ids.duplicated()])


def route_rows(df, dup_ids=()):
    """(mailbox, error) per row: the row's calendar mailbox, or None with the reason it is not pushed."""
    def route(r):
        if r["EVENT_ID"] in dup_ids:
            return None, f"ID {r['EVENT_ID']} is used by more than one row"
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
            results.append({"ID": eid, "Show": row["SHOW"], "Status": status})
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


def run_sweep():
    """One auto-sync: push changed files, then remove deleted shifts. Writes the result to SWEEP_STATUS_FILE."""
    status = {**sweep_status(), "running": True, "error": None}
    _write_json(SWEEP_STATUS_FILE, status)
    log = []
    try:
        token = get_graph_token()
        files, parsed, errors, admin_modified = load_all(token)
        state = _read_json(SWEEP_STATE_FILE, {})
        if state.get("_admin_modified") != admin_modified:
            state = {"_admin_modified": admin_modified}  # settings changed: re-check every file
        limit = sweep_test_rows()
        dups  = duplicate_ids(list(parsed.values()))
        log += [f"⚠️ {w}" for w in CONFIG_WARNINGS]
        for m in sorted(_all_calendar_mailboxes()):
            if (cat_err := ensure_outlook_categories(token, m)):
                log.append(f"⚠️ Event colors: {cat_err}")
        for f in files:
            path, last_mod = f["path"], f["modified"]
            if path in errors:
                log.append(f"❌ {path} — Error: {errors[path]}")
                continue
            if path not in parsed or state.get(path, {}).get("last_modified") == last_mod:
                continue  # not a schedule file, or unchanged
            df_f = parsed[path]
            if limit > 0:
                df_f = df_f[df_f["DATE"] >= today()].head(limit)
            stats, _, bad = sync_rows(token, df_f, dups)
            result = (f"Created:{stats['created']} Updated:{stats['updated']} Skipped:{stats['skipped']} "
                      f"Failed:{stats['failed']} Past (not pushed):{stats['past']}")
            if bad:
                result += f" Not pushed:{len(bad)} ({'; '.join(f'{e}: {m}' for e, m in bad[:3])}{'; ...' if len(bad) > 3 else ''})"
            # Files with rows not pushed are retried next sweep (e.g. once the duplicate ID is fixed)
            state[path] = {"last_modified": None if bad or stats["failed"] else last_mod,
                           "last_synced": dt.now(_PT).strftime("%Y-%m-%d %H:%M PT"), "result": result}
            log.append(f"✅ {path} — {result}")
        for folder, err in errors.items():
            if folder not in {f["path"] for f in files}:
                log.append(f"❌ {folder} — {err}")
        _write_json(SWEEP_STATE_FILE, state)

        removed, problem = remove_deleted_shifts(token, parsed, errors)
        if problem:
            status["error"] = problem
        ok_ids  = [e for e, _, ok in removed if ok]
        bad_ids = [e for e, _, ok in removed if not ok]
        if ok_ids:
            log.append(f"🗑️ Removed {len(ok_ids)} event(s) no longer in Excel: "
                       f"{', '.join(ok_ids[:10])}{' ...' if len(ok_ids) > 10 else ''}")
        if bad_ids:
            status["error"] = f"Could not delete {len(bad_ids)} removed event(s): {', '.join(bad_ids[:10])}"
    except Exception as e:
        status["error"] = str(e)
        log.append(f"❌ Sweep error: {e}")
    finally:
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
