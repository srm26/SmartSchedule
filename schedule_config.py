# -*- coding: utf-8 -*-
"""
Central configuration for SmartSchedule: facilities, Excel row colors, Outlook calendars,
overtime / double-time rules and holiday dates.

Edit this file to add a facility, change a color, or add holidays. After changing a facility or
a color, rebuild the Excel template so its dropdown and row colors match:
    python update_template.py
"""

# ─── FACILITIES ──────────────────────────────────────────────────────
# code: the value users pick in the FACILITY column on the 'By Day' sheet.
#   name        - facility name
#   city/state  - where the facility is (None = not provided yet)
#   excel_color - row fill in the Excel template, hex RGB (None = not provided yet, row stays uncolored)
#   calendar    - key into CALENDARS below (None = not provided yet; rows for this facility are NOT pushed)
#   ot_region   - key into OT_RULES / HOLIDAYS below (None = not provided yet; OT is not calculated)
FACILITIES = {
    "HARROC": {"name": "Hard Rock",       "city": "Orlando", "state": "FL",
               "excel_color": "FF7C80", "calendar": "LOEWS_ELECTRICAL", "ot_region": "ORL"},   # red
    "HELIOS": {"name": "Helios",          "city": "Orlando", "state": "FL",
               "excel_color": "CC99FF", "calendar": "LOEWS_ELECTRICAL", "ot_region": "ORL"},   # purple
    "PORBAY": {"name": "Portofino",       "city": "Orlando", "state": "FL",
               "excel_color": "FFFF66", "calendar": "LOEWS_ELECTRICAL", "ot_region": "ORL"},   # yellow
    "ROYPAC": {"name": "Royal Pacific",   "city": "Orlando", "state": "FL",
               "excel_color": "FFB366", "calendar": "LOEWS_ELECTRICAL", "ot_region": "ORL"},   # orange
    "LSAPP":  {"name": "Sapphire Falls",  "city": "Orlando", "state": "FL",
               "excel_color": "99CCFF", "calendar": "LOEWS_ELECTRICAL", "ot_region": "ORL"},   # blue
    "MARWCR": {"name": "Marriott",        "city": None,      "state": None,
               "excel_color": None,     "calendar": None,               "ot_region": None},    # TBD
    "HILLV":  {"name": "Westgate Las Vegas Resort & Casino", "city": "Las Vegas", "state": "NV",
               "excel_color": None,     "calendar": None,               "ot_region": None},    # TBD
    "SHECHI": {"name": "Sheraton Grand Chicago",             "city": "Chicago",   "state": "IL",
               "excel_color": None,     "calendar": None,               "ot_region": None},    # TBD
}

# ─── OUTLOOK CALENDARS ───────────────────────────────────────────────
# key -> display name. The mailbox address for each key is NOT stored here; it is set per environment
# in the app setting GRAPH_CALENDAR_MAILBOXES (or calendar_mailboxes in secrets.toml), e.g.
#     GRAPH_CALENDAR_MAILBOXES = "LOEWS_ELECTRICAL=loews-electrical@example.com"
# A calendar without a mailbox setting is treated as unmapped: its rows are not pushed.
CALENDARS = {
    "LOEWS_ELECTRICAL": "Loews Electrical",
}

# ─── OUTLOOK EVENT COLORS ────────────────────────────────────────────
# FUNCTION -> Outlook category put on the event; the category's color shows on the calendar.
# Functions not listed get no category (normal calendar color).
# color = Outlook preset: preset0 red, preset1 orange, preset2 brown, preset3 yellow, preset4 green,
#         preset5 teal, preset6 olive, preset7 blue, preset8 purple, preset9 cranberry, preset12 gray
# The app creates missing categories in each calendar (needs the MailboxSettings.ReadWrite Graph
# permission); an existing category keeps whatever color it already has in that calendar.
OUTLOOK_CATEGORIES = {
    "SET":    {"name": "SET",    "color": "preset4"},   # green
    "STRIKE": {"name": "STRIKE", "color": "preset0"},   # red
}

# ─── OVERTIME RULES ──────────────────────────────────────────────────
# region -> rules. Each OT hour is counted once, with the first matching reason in this order:
# holiday, Sunday, Saturday, over daily_hours on a normal workday.
OT_RULES = {
    "ORL": {"daily_hours": 8, "saturday": True, "sunday": True, "holidays": True},
}

# region -> holiday dates (YYYY-MM-DD) that count as OT.
HOLIDAYS = {
    "ORL": [
        "2027-01-01", "2027-01-18", "2027-02-15", "2027-05-28", "2027-05-31", "2027-06-18",
        "2027-07-02", "2027-07-05", "2027-09-03", "2027-09-06", "2027-11-25", "2027-11-26",
        "2027-12-24", "2027-12-27",
    ],
}

# ─── DOUBLE-TIME RULES ───────────────────────────────────────────────
# TBD: business rules not provided yet. Double-time is not calculated until rules are added here.
DT_RULES = {}
