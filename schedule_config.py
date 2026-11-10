# -*- coding: utf-8 -*-
"""
SmartSchedule configuration, read from App Service app settings (environment variables).
One setting per item; values are separated by ';' (blank = not provided yet):

  FACILITY_<CODE>         = name; city; state (2 letters); calendar key; Excel color (hex); OT region
      e.g. FACILITY_HARROC = Hard Rock; Orlando; FL; LOEWS_ELECTRICAL; FF7C80; ORL
      - calendar key: its mailbox is set in GRAPH_CALENDAR_MAILBOXES. Blank = rows are not pushed.
      - Excel color: row fill in the template. Blank = no color.   OT region: blank = no OT rules.
  BRANCH_<ID PREFIX>      = state; state code; Outlook time zone
      e.g. BRANCH_LV = Nevada; NV; Pacific Standard Time
      Each spreadsheet uses one ID prefix (LV-0001). The state decides the time zone, and a spreadsheet
      may only hold facilities in its own state.
  OT_RULES_<REGION>       = daily hours; then any of Saturday, Sunday, Holidays
      e.g. OT_RULES_ORL = 8; Saturday; Sunday; Holidays
      Each OT hour is counted once: holiday, then Sunday, then Saturday, then over daily hours.
  HOLIDAYS_<REGION>       = dates (YYYY-MM-DD) separated by commas
  OUTLOOK_COLOR_<FUNCTION> = color of that function's Outlook events (use _ for spaces: BOOTH_CHECK)
      e.g. OUTLOOK_COLOR_SET = green.   Colors: red, orange, brown, yellow, green, teal, olive, blue,
      purple, cranberry, gray.   Functions without a setting keep the normal calendar color.

Locally (tests, update_template.py) the same NAME=VALUE lines are read from settings.local.env when no
such environment variables are set. Settings that can't be read are skipped and listed in CONFIG_WARNINGS.
After changing a facility or Excel color, rebuild the Excel files:  python update_template.py --from-azure
"""
import os
import re
from pathlib import Path

LOCAL_FILE = Path(__file__).with_name("settings.local.env")
PREFIXES   = ("FACILITY_", "BRANCH_", "OT_RULES_", "HOLIDAYS_", "OUTLOOK_COLOR_")
COLORS     = {"red": "preset0", "orange": "preset1", "brown": "preset2", "yellow": "preset3",
              "green": "preset4", "teal": "preset5", "olive": "preset6", "blue": "preset7",
              "purple": "preset8", "cranberry": "preset9", "gray": "preset12", "grey": "preset12"}


def _read_settings():
    """{NAME: value} for our settings: environment variables, else settings.local.env."""
    found = {k.upper(): v for k, v in os.environ.items() if k.upper().startswith(PREFIXES)}
    if not found and LOCAL_FILE.exists():
        for line in LOCAL_FILE.read_text(encoding="utf8").splitlines():
            name, sep, value = line.partition("=")
            if sep and not line.lstrip().startswith("#") and name.strip().upper().startswith(PREFIXES):
                found[name.strip().upper()] = value
    return found


def _parts(value, n):
    """The first n ';'-separated values, stripped; missing or blank ones become None."""
    p = [x.strip() or None for x in str(value).split(";")]
    return (p + [None] * n)[:n]


def load(settings):
    """Build (FACILITIES, BRANCHES, OT_RULES, HOLIDAYS, OUTLOOK_CATEGORIES, warnings) from settings."""
    fac, br, ot, hol, cat, warn = {}, {}, {}, {}, {}, []
    for name, value in sorted(settings.items()):
        try:
            if name.startswith("FACILITY_"):
                fname, city, state, calendar, color, region = _parts(value, 6)
                if not fname:
                    raise ValueError("facility name is blank")
                if state and not re.fullmatch(r"[A-Za-z]{2}", state):
                    raise ValueError(f"state '{state}' should be 2 letters, e.g. FL")
                if color and not re.fullmatch(r"#?[0-9A-Fa-f]{6}", color):
                    raise ValueError(f"Excel color '{color}' should be 6 hex digits, e.g. FF7C80")
                fac[name[len("FACILITY_"):]] = {
                    "name": fname, "city": city, "state": state.upper() if state else None,
                    "calendar": calendar.upper() if calendar else None,
                    "excel_color": color.lstrip("#").upper() if color else None,
                    "ot_region": region.upper() if region else None}
            elif name.startswith("BRANCH_"):
                state, code, tz = _parts(value, 3)
                if not (state and code and tz):
                    raise ValueError("needs state; state code; time zone")
                br[name[len("BRANCH_"):]] = {"state": state, "state_code": code.upper(), "time_zone": tz}
            elif name.startswith("OT_RULES_"):
                words = [w.strip().lower() for w in str(value).split(";") if w.strip()]
                if not words or not re.fullmatch(r"\d+(\.\d+)?", words[0]):
                    raise ValueError("should start with the daily hours, e.g. 8; Saturday; Sunday; Holidays")
                unknown = set(words[1:]) - {"saturday", "sunday", "holidays"}
                if unknown:
                    raise ValueError(f"unknown rule(s): {', '.join(sorted(unknown))}")
                hours = float(words[0])
                ot[name[len("OT_RULES_"):]] = {"daily_hours": int(hours) if hours.is_integer() else hours,
                                               **{w: w in words for w in ("saturday", "sunday", "holidays")}}
            elif name.startswith("HOLIDAYS_"):
                dates = [d.strip() for d in str(value).split(",") if d.strip()]
                bad = [d for d in dates if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d)]
                if bad:
                    raise ValueError(f"dates should look like 2027-01-01: {', '.join(bad)}")
                hol[name[len("HOLIDAYS_"):]] = dates
            elif name.startswith("OUTLOOK_COLOR_"):
                color = str(value).strip().lower()
                preset = COLORS.get(color) or (color if re.fullmatch(r"preset\d+", color) else None)
                if not preset:
                    raise ValueError(f"unknown color '{value.strip()}' (use {', '.join(list(COLORS)[:-1])})")
                fn = name[len("OUTLOOK_COLOR_"):].replace("_", " ")
                cat[fn] = {"name": fn, "color": preset}
        except ValueError as e:
            warn.append(f"Setting {name} skipped: {e}")
    if not fac:
        warn.append("No FACILITY_ settings found, so no facility can be routed to Outlook.")
    for code, f in fac.items():
        if f["ot_region"] and f["ot_region"] not in ot:
            warn.append(f"FACILITY_{code} uses OT region {f['ot_region']}, which has no OT_RULES_{f['ot_region']}.")
    return fac, br, ot, hol, cat, warn


FACILITIES, BRANCHES, OT_RULES, HOLIDAYS, OUTLOOK_CATEGORIES, CONFIG_WARNINGS = load(_read_settings())

# Calendar keys in use (shown as-is in messages: it's the name to put in GRAPH_CALENDAR_MAILBOXES).
CALENDARS = {f["calendar"]: f["calendar"] for f in FACILITIES.values() if f["calendar"]}

# Double-time: TBD - business rules not provided yet, so double-time is not calculated.
DT_RULES = {}
