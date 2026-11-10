# -*- coding: utf-8 -*-
"""
Facility routing and overtime calculation. No Streamlit here, so it can be tested on its own.
All facility / calendar / rule data comes from schedule_config.py.
"""
import re
import pandas as pd
from schedule_config import FACILITIES, CALENDARS, OT_RULES, HOLIDAYS

# ─── FACILITY ROUTING ────────────────────────────────────────────────
def facility_route(facility, calendar_mailboxes):
    """(mailbox, '') for a row's facility, or (None, reason) when it can't be routed.
    calendar_mailboxes: {CALENDAR KEY: mailbox} from the GRAPH_CALENDAR_MAILBOXES setting."""
    code = str(facility or "").strip().upper()
    if not code:
        return None, "FACILITY is blank"
    fac = FACILITIES.get(code)
    if fac is None:
        return None, f"'{code}' is not a known facility code"
    if not fac["calendar"]:
        return None, f"no Outlook calendar is configured for {code} yet"
    mailbox = calendar_mailboxes.get(fac["calendar"])
    if not mailbox:
        return None, (f"the {CALENDARS.get(fac['calendar'], fac['calendar'])} calendar has no mailbox "
                      f"set (GRAPH_CALENDAR_MAILBOXES)")
    return mailbox, ""

def has_facility(value):
    """False for rows from files without a FACILITY column (value None/NaN)."""
    return value is not None and not (isinstance(value, float) and pd.isna(value))

# ─── OVERTIME ────────────────────────────────────────────────────────
_TIME_RE = re.compile(r"(\d{1,2}):(\d{2})\s*[-–]\s*(\d{1,2}):(\d{2})")  # same TIME format the importer accepts

def shift_span(time_str):
    """(start minute, hours) for 'HH:MM - HH:MM', or (None, 0.0) for blank/untimed shifts.
    Overnight shifts ('22:00 - 02:00') end the next day."""
    m = _TIME_RE.match(str(time_str or "").strip())
    if not m:
        return None, 0.0
    start = int(m[1]) * 60 + int(m[2])
    end   = int(m[3]) * 60 + int(m[4])
    if end <= start:
        end += 24 * 60
    return start, (end - start) / 60

def _ot_region(row):
    """Rules region for a row: from its facility, or from the ID prefix when FACILITY is blank or absent."""
    fac = row.get("FACILITY")
    if has_facility(fac) and fac:
        return FACILITIES.get(str(fac), {}).get("ot_region")
    return row.get("BRANCH") or None

def _premium_day(day, region, rules):
    """Reason the whole day is OT ('Holiday', 'Sunday', 'Saturday'), or ''."""
    if rules.get("holidays") and day.strftime("%Y-%m-%d") in HOLIDAYS.get(region, []):
        return "Holiday"
    if rules.get("sunday") and day.weekday() == 6:
        return "Sunday"
    if rules.get("saturday") and day.weekday() == 5:
        return "Saturday"
    return ""

def compute_overtime(df):
    """One row per person per shift with HOURS, REG_HOURS, OT_HOURS, OT_REASON and OT_SHIFT.

    - A shift counts on its start date.
    - Holiday / Sunday / Saturday: every hour of the shift is OT.
    - Other days: a person's hours past daily_hours (summed over that day's shifts in start-time order) are OT.
    - Each hour is counted once, under one reason.
    - Unnamed crew (+N) can't be tracked across shifts, so each one is treated as its own person per shift.
    - STATUS is 'ok', 'no OT rules' (region has no rules configured) or 'no TIME' (hours unknown).
    """
    people = []
    for _, r in df.iterrows():
        start, hours = shift_span(r["TIME"])
        base = {"EVENT_ID": r["EVENT_ID"], "DATE": r["DATE"], "SHOW": r["SHOW"], "FUNCTION": r["FUNCTION"],
                "TIME": r["TIME"], "FACILITY": r.get("FACILITY") if has_facility(r.get("FACILITY")) else "",
                "REGION": _ot_region(r), "_START": start, "HOURS": hours, "_TIMED": start is not None}
        for name in r["CREW_LIST"]:
            people.append({**base, "PERSON": name, "NAMED": True, "_KEY": name})
        for k in range(int(r.get("EXTRA_CREW", 0) or 0)):
            people.append({**base, "PERSON": "Unnamed crew", "NAMED": False, "_KEY": f"~{r['EVENT_ID']}#{k}"})

    cols = ["PERSON", "NAMED", "DATE", "SHOW", "FUNCTION", "FACILITY", "TIME", "HOURS",
            "REG_HOURS", "OT_HOURS", "OT_REASON", "OT_SHIFT", "STATUS", "EVENT_ID"]
    if not people:
        return pd.DataFrame(columns=cols)

    ps = pd.DataFrame(people)
    ps["REG_HOURS"], ps["OT_HOURS"], ps["OT_REASON"], ps["OT_SHIFT"], ps["STATUS"] = 0.0, 0.0, "", False, "ok"
    for (_, day), grp in ps.groupby(["_KEY", "DATE"]):
        worked = 0.0  # normal-workday hours so far for this person on this day
        for idx, s in grp.sort_values("_START", na_position="last").iterrows():
            rules = OT_RULES.get(s["REGION"])
            if not rules:
                ps.loc[idx, ["REG_HOURS", "STATUS"]] = [s["HOURS"], "no OT rules"]
                continue
            reason = _premium_day(day, s["REGION"], rules)
            if reason:
                ot = s["HOURS"]
            else:
                ot = max(0.0, min(s["HOURS"], worked + s["HOURS"] - rules["daily_hours"]))
                worked += s["HOURS"]
                reason = f"Over {rules['daily_hours']} hrs/day" if ot > 0 else ""
            ps.loc[idx, ["REG_HOURS", "OT_HOURS", "OT_REASON"]] = [s["HOURS"] - ot, ot, reason]
            # An untimed shift on a premium day is still an OT shift, even though its hours are unknown
            ps.loc[idx, "OT_SHIFT"] = ot > 0 or (bool(reason) and not s["_TIMED"])
            if not s["_TIMED"]:
                ps.loc[idx, "STATUS"] = "no TIME"
    return ps[cols].sort_values(["DATE", "PERSON"]).reset_index(drop=True)

def ot_summary(ps, by):
    """Totals per `by` column: shifts, hours, OT shifts, OT hours, and OT hours per reason (no double counting)."""
    if ps.empty:
        return pd.DataFrame(columns=[by, "Shifts", "Hours", "Regular Hours", "OT Shifts", "OT Hours"])
    out = ps.groupby(by).agg(**{"Shifts": ("EVENT_ID", "size"), "Hours": ("HOURS", "sum"),
                                "Regular Hours": ("REG_HOURS", "sum"), "OT Shifts": ("OT_SHIFT", "sum"),
                                "OT Hours": ("OT_HOURS", "sum")})
    reasons = ps[ps["OT_HOURS"] > 0].pivot_table(index=by, columns="OT_REASON", values="OT_HOURS",
                                                 aggfunc="sum", fill_value=0.0)
    reasons.columns = [f"OT Hours: {c}" for c in reasons.columns]
    out = out.join(reasons).fillna(0.0)
    out["OT Shifts"] = out["OT Shifts"].astype(int)
    return out.reset_index().sort_values("OT Hours", ascending=False)
