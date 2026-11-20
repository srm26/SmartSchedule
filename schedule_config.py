# -*- coding: utf-8 -*-
"""
SmartSchedule settings, kept by an admin in the Excel workbook ADMIN_FILE on SharePoint
(ElectricalSchedule / z_Electrical Scheduling App Admin). One sheet per table:

  Calendars        CALENDAR | SHAREPOINT FOLDER | MAILBOX | TIME ZONE | COLOR BY | OT REGION
      Every schedule file in a calendar's SharePoint folder goes to that calendar's mailbox.
      COLOR BY: FUNCTION (Function Colors sheet), FACILITY (Facilities sheet) or blank (no colors).
  Facilities       FACILITY CODE | FACILITY NAME | CITY | STATE | CALENDAR | OUTLOOK COLOR | EXCEL COLOR
      EXCEL COLOR is the template's row fill (6 hex digits); OUTLOOK COLOR is used by COLOR BY = FACILITY.
  Function Colors  FUNCTION | OUTLOOK COLOR
  OT Rules         OT REGION | DAILY HOURS | SATURDAY | SUNDAY | HOLIDAYS      (Yes / No)
  Holidays         OT REGION | DATE | HOLIDAY
  Alerts           ADMIN EMAIL | SEND FROM
      One admin email per row; new auto-sync problems are emailed to all of them. SEND FROM (first filled
      cell) is the mailbox the email is sent from; the app needs the Mail.Send permission for it.

Outlook colors: red, orange, brown, yellow, green, teal, olive, blue, purple, cranberry, gray.
The app re-reads the workbook on every sync. Rows that can't be read are skipped and listed in CONFIG_WARNINGS.
apply() updates the dicts below in place, so modules that imported them see the new values.
"""
import re
import pandas as pd

ADMIN_FOLDER = "z_Electrical Scheduling App Admin"
ADMIN_FILE   = "Electrical Scheduling Admin.xlsx"
SHEETS = {
    "Calendars":       ["CALENDAR", "SHAREPOINT FOLDER", "MAILBOX", "TIME ZONE", "COLOR BY", "OT REGION"],
    "Facilities":      ["FACILITY CODE", "FACILITY NAME", "CITY", "STATE", "CALENDAR", "OUTLOOK COLOR", "EXCEL COLOR"],
    "Function Colors": ["FUNCTION", "OUTLOOK COLOR"],
    "OT Rules":        ["OT REGION", "DAILY HOURS", "SATURDAY", "SUNDAY", "HOLIDAYS"],
    "Holidays":        ["OT REGION", "DATE", "HOLIDAY"],
    "Alerts":          ["ADMIN EMAIL", "SEND FROM"],
}
COLORS = {"red": "preset0", "orange": "preset1", "brown": "preset2", "yellow": "preset3",
          "green": "preset4", "teal": "preset5", "olive": "preset6", "blue": "preset7",
          "purple": "preset8", "cranberry": "preset9", "gray": "preset12", "grey": "preset12"}

CALENDARS, FACILITIES, FUNCTION_COLORS, OT_RULES, HOLIDAYS, CONFIG_WARNINGS = {}, {}, {}, {}, {}, []
ALERTS = {"to": [], "sender": ""}
# Double-time: TBD - business rules not provided yet, so double-time is not calculated.
DT_RULES = {}


def _text(v):
    s = "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _color(value):
    """Outlook preset for a color name, None when blank; ValueError when unknown."""
    if not value:
        return None
    preset = COLORS.get(value.lower())
    if not preset:
        raise ValueError(f"unknown Outlook color '{value}' (use {', '.join(list(COLORS)[:-1])})")
    return preset


def _yes(v):
    return v.lower() in ("yes", "y", "true", "x", "1")


def load(tables):
    """tables: {sheet name: DataFrame}.
    Returns (calendars, facilities, function_colors, ot_rules, holidays, warnings, alerts)."""
    cal, fac, fn, ot, hol, warn = {}, {}, {}, {}, {}, []
    alerts = {"to": [], "sender": ""}

    def rows(sheet):
        df = tables.get(sheet)
        if df is None:
            warn.append(f"Admin workbook has no '{sheet}' sheet.")
            return []
        df = df.rename(columns=lambda c: str(c).strip().upper())
        missing = [c for c in SHEETS[sheet] if c not in df.columns]
        if missing:
            warn.append(f"Admin '{sheet}' sheet is missing column(s): {', '.join(missing)}.")
            return []
        out = []
        for i, r in df.iterrows():
            vals = {c: _text(r[c]) for c in SHEETS[sheet]}
            if any(vals.values()):
                out.append((i + 2, vals))  # Excel row number, for messages
        return out

    def each(sheet, fill):
        for n, r in rows(sheet):
            try:
                fill(r)
            except ValueError as e:
                warn.append(f"Admin '{sheet}' row {n} skipped: {e}")

    def calendar(r):
        key = r["CALENDAR"].upper()
        if not key or not r["SHAREPOINT FOLDER"] or not r["TIME ZONE"]:
            raise ValueError("needs CALENDAR, SHAREPOINT FOLDER and TIME ZONE")
        color_by = r["COLOR BY"].upper()
        if color_by not in ("", "FUNCTION", "FACILITY"):
            raise ValueError(f"COLOR BY '{r['COLOR BY']}' should be FUNCTION, FACILITY or blank")
        cal[key] = {"key": key, "folder": r["SHAREPOINT FOLDER"], "mailbox": r["MAILBOX"] or None,
                    "time_zone": r["TIME ZONE"], "color_by": color_by or None,
                    "ot_region": r["OT REGION"].upper() or None}

    def facility(r):
        code = r["FACILITY CODE"].upper()
        if not code or not r["FACILITY NAME"]:
            raise ValueError("needs FACILITY CODE and FACILITY NAME")
        if r["STATE"] and not re.fullmatch(r"[A-Za-z]{2}", r["STATE"]):
            raise ValueError(f"STATE '{r['STATE']}' should be 2 letters, e.g. FL")
        if r["EXCEL COLOR"] and not re.fullmatch(r"#?[0-9A-Fa-f]{6}", r["EXCEL COLOR"]):
            raise ValueError(f"EXCEL COLOR '{r['EXCEL COLOR']}' should be 6 hex digits, e.g. FF7C80")
        fac[code] = {"name": r["FACILITY NAME"], "city": r["CITY"] or None, "state": r["STATE"].upper() or None,
                     "calendar": r["CALENDAR"].upper() or None, "outlook_color": _color(r["OUTLOOK COLOR"]),
                     "excel_color": r["EXCEL COLOR"].lstrip("#").upper() or None}

    def function_color(r):
        if not r["FUNCTION"] or not r["OUTLOOK COLOR"]:
            raise ValueError("needs FUNCTION and OUTLOOK COLOR")
        fn[r["FUNCTION"].upper()] = _color(r["OUTLOOK COLOR"])

    def ot_rule(r):
        if not r["OT REGION"] or not re.fullmatch(r"\d+(\.\d+)?", r["DAILY HOURS"]):
            raise ValueError("needs OT REGION and DAILY HOURS as a number, e.g. 8")
        hours = float(r["DAILY HOURS"])
        ot[r["OT REGION"].upper()] = {"daily_hours": int(hours) if hours.is_integer() else hours,
                                      **{k.lower(): _yes(r[k]) for k in ("SATURDAY", "SUNDAY", "HOLIDAYS")}}

    def holiday(r):
        # An Excel date cell reads as '2027-01-01 00:00:00'; typed text must be 2027-01-01
        if not r["OT REGION"] or not re.fullmatch(r"\d{4}-\d{2}-\d{2}( 00:00:00)?", r["DATE"]):
            raise ValueError(f"needs OT REGION and a DATE like 2027-01-01 (got '{r['DATE']}')")
        hol.setdefault(r["OT REGION"].upper(), []).append(r["DATE"][:10])

    def alert(r):
        for col in ("ADMIN EMAIL", "SEND FROM"):
            if r[col] and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", r[col]):
                raise ValueError(f"{col} '{r[col]}' is not an email address")
        if r["ADMIN EMAIL"]:
            alerts["to"].append(r["ADMIN EMAIL"])
        alerts["sender"] = alerts["sender"] or r["SEND FROM"]

    for sheet, fill in [("Calendars", calendar), ("Facilities", facility), ("Function Colors", function_color),
                        ("OT Rules", ot_rule), ("Holidays", holiday), ("Alerts", alert)]:
        each(sheet, fill)

    if not cal:
        warn.append("No calendars in the admin workbook, so nothing can be pushed to Outlook.")
    for key, c in cal.items():
        if not c["mailbox"]:
            warn.append(f"Calendar {key} has no MAILBOX yet, so its rows are not pushed.")
        if c["ot_region"] and c["ot_region"] not in ot:
            warn.append(f"Calendar {key} uses OT region {c['ot_region']}, which is not on the OT Rules sheet.")
    for code, f in fac.items():
        if f["calendar"] and f["calendar"] not in cal:
            warn.append(f"Facility {code} names calendar {f['calendar']}, which is not on the Calendars sheet.")
    if "Alerts" in tables and not (alerts["to"] and alerts["sender"]):
        warn.append("Problems are not emailed: the Alerts sheet needs at least one ADMIN EMAIL and a SEND FROM mailbox.")
    return cal, fac, fn, ot, hol, warn, alerts


def read(src):
    """{sheet name: DataFrame} from the admin workbook (path, bytes buffer or file)."""
    return pd.read_excel(src, sheet_name=None, dtype=str)


def apply(src):
    """Load the admin workbook into the module dicts. If it can't be read, the previous settings are kept."""
    try:
        new = load(read(src))
    except Exception as e:
        CONFIG_WARNINGS[:] = [f"Could not read the admin workbook ({e}); using the last settings that loaded."]
        return
    for target, value in zip((CALENDARS, FACILITIES, FUNCTION_COLORS, OT_RULES, HOLIDAYS), new):
        target.clear()
        target.update(value)
    CONFIG_WARNINGS[:] = new[5]
    ALERTS.clear()
    ALERTS.update(new[6])


def outlook_categories(calendar):
    """{category name: Outlook preset} used by a calendar (by its COLOR BY)."""
    c = CALENDARS.get(calendar) or {}
    if c.get("color_by") == "FACILITY":
        return {f["name"]: f["outlook_color"] for f in FACILITIES.values() if f["outlook_color"]}
    if c.get("color_by") == "FUNCTION":
        return dict(FUNCTION_COLORS)
    return {}


def event_category(calendar, function, facility):
    """Outlook category (color) name for one row, or None."""
    c = CALENDARS.get(calendar) or {}
    if c.get("color_by") == "FACILITY":
        f = FACILITIES.get(str(facility or "").upper())
        return f["name"] if f and f["outlook_color"] else None
    if c.get("color_by") == "FUNCTION":
        fn = str(function or "").upper()
        return fn if fn in FUNCTION_COLORS else None
    return None
