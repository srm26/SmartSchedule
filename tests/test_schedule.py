# -*- coding: utf-8 -*-
"""Run:  python -m unittest discover -s tests -v   (from the project folder)"""
import hashlib
import shutil
import sys
import tempfile
import types
import unittest
import warnings
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import schedule_config                                                      # noqa: E402
import sync_engine as se                                                    # noqa: E402
from schedule_config import CALENDARS, FACILITIES, HOLIDAYS                 # noqa: E402
from schedule_logic import compute_overtime, ot_summary, shift_span        # noqa: E402

TEMPLATE = ROOT / "templates" / "GES Electrical Show Schedule TEMPLATE.xlsx"
warnings.simplefilter("ignore")

# Same content as the admin workbook on SharePoint (test mailboxes)
ADMIN = {
    "Calendars": (["CALENDAR", "SHAREPOINT FOLDER", "MAILBOX", "TIME ZONE", "COLOR BY", "OT REGION"], [
        ["CHI", "Chicago", "test@x.com", "Central Standard Time", "FUNCTION", ""],
        ["LV", "Las Vegas", "test@x.com", "Pacific Standard Time", "FUNCTION", ""],
        ["ORL LOEWS", "Orlando - Lowes", "loews@x.com", "Eastern Standard Time", "FACILITY", "ORL"],
        ["ORL MARRIOTT", "Orlando Marriott", "test@x.com", "Eastern Standard Time", "FUNCTION", "ORL"]]),
    "Facilities": (["FACILITY CODE", "FACILITY NAME", "CITY", "STATE", "CALENDAR", "OUTLOOK COLOR", "EXCEL COLOR"], [
        ["HARROC", "Hard Rock", "Orlando", "FL", "ORL LOEWS", "red", "FF7C80"],
        ["HELIOS", "Helios", "Orlando", "FL", "ORL LOEWS", "purple", "CC99FF"],
        ["PORBAY", "Portofino", "Orlando", "FL", "ORL LOEWS", "yellow", "FFFF66"],
        ["ROYPAC", "Royal Pacific", "Orlando", "FL", "ORL LOEWS", "orange", "FFB366"],
        ["LSAPP", "Sapphire Falls", "Orlando", "FL", "ORL LOEWS", "blue", "99CCFF"],
        ["MARWCR", "Marriott", "Orlando", "FL", "ORL MARRIOTT", "", ""],
        ["HILLV", "Westgate Las Vegas Resort & Casino", "Las Vegas", "NV", "LV", "", ""],
        ["SHECHI", "Sheraton Grand Chicago", "Chicago", "IL", "CHI", "", ""]]),
    "Function Colors": (["FUNCTION", "OUTLOOK COLOR"], [["SET", "green"], ["STRIKE", "red"]]),
    "OT Rules": (["OT REGION", "DAILY HOURS", "SATURDAY", "SUNDAY", "HOLIDAYS"], [["ORL", "8", "Yes", "Yes", "Yes"]]),
    "Holidays": (["OT REGION", "DATE", "HOLIDAY"], [["ORL", d, ""] for d in
                 ["2027-01-01", "2027-01-18", "2027-02-15", "2027-05-31", "2027-07-05", "2027-11-25", "2027-12-24"]]),
}


def _tables(admin=ADMIN):
    return {name: pd.DataFrame(rows, columns=cols, dtype=str) for name, (cols, rows) in admin.items()}


def _install(admin=ADMIN):
    """Load test settings into schedule_config's dicts, the way apply() does for the real workbook."""
    new = schedule_config.load(_tables(admin))
    for target, value in zip((schedule_config.CALENDARS, schedule_config.FACILITIES, schedule_config.FUNCTION_COLORS,
                              schedule_config.OT_RULES, schedule_config.HOLIDAYS), new):
        target.clear()
        target.update(value)
    schedule_config.CONFIG_WARNINGS[:] = new[5]


_install()

FUTURE = (se.today() + pd.Timedelta(days=30)).strftime("%Y-%m-%d")
PAST   = (se.today() - pd.Timedelta(days=30)).strftime("%Y-%m-%d")


def _row(date, time, crew, extra=0, facility="HARROC", eid=None, calendar="ORL LOEWS"):
    return {"EVENT_ID": eid or f"T-{date}-{time}", "DATE": pd.Timestamp(date), "SHOW": "S", "FUNCTION": "SET",
            "TIME": time, "CREW_LIST": crew, "EXTRA_CREW": extra, "FACILITY": facility, "CALENDAR": calendar}


def _shift(eid, date=FUTURE, calendar="LV", function="SET", facility=""):
    return {"EVENT_ID": eid, "DATE": pd.Timestamp(date), "END_DATE": pd.NaT, "SHOW": "S", "FUNCTION": function,
            "LOCATION": "", "TIME": "08:00 - 12:00", "FOREMAN": "TBD", "NOTES": "", "LINK": "", "LINK_TEXT": "",
            "FACILITY": facility, "CALENDAR": calendar, "TIME_ZONE": "Pacific Standard Time"}


class AdminSettings(unittest.TestCase):
    def test_loads_all_sheets(self):
        cal, fac, fn, ot, hol, warn = schedule_config.load(_tables())
        self.assertEqual(list(cal), ["CHI", "LV", "ORL LOEWS", "ORL MARRIOTT"])
        self.assertEqual(cal["ORL LOEWS"], {"key": "ORL LOEWS", "folder": "Orlando - Lowes", "mailbox": "loews@x.com",
                                            "time_zone": "Eastern Standard Time", "color_by": "FACILITY",
                                            "ot_region": "ORL"})
        self.assertEqual(fac["HARROC"]["outlook_color"], "preset0")
        self.assertEqual(fac["HARROC"]["excel_color"], "FF7C80")
        self.assertIsNone(fac["MARWCR"]["outlook_color"])
        self.assertEqual(fn, {"SET": "preset4", "STRIKE": "preset0"})
        self.assertEqual(ot["ORL"], {"daily_hours": 8, "saturday": True, "sunday": True, "holidays": True})
        self.assertEqual(hol["ORL"][0], "2027-01-01")
        self.assertEqual(warn, [])

    def test_bad_rows_skipped_with_warnings(self):
        bad = {k: (cols, [list(r) for r in rows]) for k, (cols, rows) in ADMIN.items()}
        bad["Calendars"][1].append(["NOPE", "", "", "", "", ""])                   # no folder / time zone
        bad["Calendars"][1].append(["X", "X", "", "Eastern Standard Time", "SHOW", ""])  # bad COLOR BY
        bad["Facilities"][1].append(["BAD", "Bad", "", "Florida", "ORL LOEWS", "", ""])  # state not 2 letters
        bad["Facilities"][1].append(["BAD2", "Bad", "", "FL", "ORL LOEWS", "pink", ""])  # unknown color
        bad["Function Colors"][1].append(["PREP", "sparkly"])
        bad["OT Rules"][1].append(["LV", "eight", "Yes", "No", "No"])
        bad["Holidays"][1].append(["ORL", "Jan 1st", ""])
        cal, fac, fn, ot, hol, warn = schedule_config.load(_tables(bad))
        self.assertNotIn("NOPE", cal)
        self.assertNotIn("X", cal)
        self.assertNotIn("BAD", fac)
        self.assertNotIn("PREP", fn)
        self.assertNotIn("LV", ot)
        self.assertEqual(len(warn), 7)
        self.assertTrue(any("Facilities' row 10" in w and "2 letters" in w for w in warn))

    def test_missing_sheet_and_checks(self):
        t = _tables()
        del t["Holidays"]
        t["Calendars"].loc[0, "MAILBOX"] = ""
        t["Calendars"].loc[1, "OT REGION"] = "LV"
        t["Facilities"].loc[0, "CALENDAR"] = "GAYLORD"
        warn = schedule_config.load(t)[5]
        self.assertTrue(any("no 'Holidays' sheet" in w for w in warn))
        self.assertTrue(any("CHI has no MAILBOX" in w for w in warn))
        self.assertTrue(any("OT region LV" in w for w in warn))
        self.assertTrue(any("calendar GAYLORD" in w for w in warn))

    def test_unreadable_workbook_keeps_last_settings(self):
        schedule_config.apply(str(ROOT / "no such file.xlsx"))
        try:
            self.assertEqual(len(CALENDARS), 4)
            self.assertIn("Could not read the admin workbook", schedule_config.CONFIG_WARNINGS[0])
        finally:
            _install()

    def test_colors_by_function_or_facility(self):
        cat = schedule_config.event_category
        self.assertEqual(cat("LV", "SET", "HARROC"), "SET")          # LV / CHI: by FUNCTION
        self.assertEqual(cat("CHI", "PREP", ""), None)
        self.assertEqual(cat("ORL LOEWS", "SET", "HARROC"), "Hard Rock")  # Orlando: by FACILITY
        self.assertEqual(cat("ORL LOEWS", "STRIKE", "LSAPP"), "Sapphire Falls")
        self.assertEqual(cat("ORL LOEWS", "SET", ""), None)
        self.assertEqual(cat("ORL MARRIOTT", "SET", "MARWCR"), "SET")  # one facility: SET / STRIKE (FR-021)
        self.assertEqual(schedule_config.outlook_categories("ORL LOEWS")["Helios"], "preset8")
        self.assertEqual(schedule_config.outlook_categories("LV"), {"SET": "preset4", "STRIKE": "preset0"})


class Routing(unittest.TestCase):
    def test_folder_decides_calendar(self):
        df = pd.DataFrame([_shift("A", calendar="ORL LOEWS", facility="HARROC"),
                           _shift("B", calendar="ORL LOEWS", facility=""),        # blank facility: still pushed
                           _shift("C", calendar="ORL LOEWS", facility="XYZ"),     # unknown facility: still pushed
                           _shift("D", calendar="ORL LOEWS", facility="MARWCR"),  # other calendar's facility
                           _shift("E", calendar="CHI")])
        mbx, err = se.route_rows(df)
        self.assertEqual(list(mbx), ["loews@x.com"] * 4 + ["test@x.com"])
        self.assertEqual(list(err), [""] * 5)
        notes = se.facility_warnings(df)
        self.assertEqual(notes[0], "")
        self.assertEqual(notes[1], "")
        self.assertIn("not on the admin Facilities sheet", notes[2])
        self.assertIn("belongs to the ORL MARRIOTT calendar", notes[3])

    def test_no_mailbox_unknown_calendar_and_duplicates_held_back(self):
        CALENDARS["CHI"]["mailbox"] = None
        try:
            df = pd.DataFrame([_shift("A", calendar="CHI"), _shift("B", calendar="GONE"), _shift("LV-1")])
            mbx, err = se.route_rows(df, {"LV-1"})
        finally:
            _install()
        self.assertEqual(list(mbx), [None, None, None])
        self.assertIn("no MAILBOX", err[0])
        self.assertIn("not on the admin Calendars sheet", err[1])
        self.assertIn("used by more than one row", err[2])

    def test_duplicate_ids_in_and_across_files(self):
        a = pd.DataFrame([_shift("LV-1"), _shift("LV-2"), _shift("LV-2")])
        b = pd.DataFrame([_shift("LV-1"), _shift("CHI-1")])
        self.assertEqual(se.duplicate_ids([a, b]), {"LV-1", "LV-2"})
        self.assertEqual(se.duplicate_ids([]), set())


class Parser(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.path = self.tmp / "t.xlsx"
        shutil.copy(TEMPLATE, self.path)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)  # pandas may still hold the file open on Windows

    def test_formula_ids_without_saved_values(self):
        # The template's ID column is a formula; saved by a script, Excel hasn't stored its values yet
        wb = load_workbook(self.path)
        ws = wb["By Day"]
        self.assertTrue(str(ws["A2"].value).startswith("=IF("))
        ws["A5"] = "TYPED-1"   # a typed ID still wins
        ws["G2"], ws["G3"] = "HARROC", "helios"
        wb["Lists"]["P2"] = "ORL"
        wb.save(self.path)
        df = se.parse_lv_byDay(str(self.path), "By Day", "ORL LOEWS")
        self.assertEqual(list(df["EVENT_ID"][:5]), ["ORL-0001", "ORL-0002", "ORL-0003", "TYPED-1", "ORL-0005"])
        self.assertEqual(list(df["FACILITY"][:3]), ["HARROC", "HELIOS", ""])
        self.assertTrue((df["CALENDAR"] == "ORL LOEWS").all())
        self.assertTrue((df["TIME_ZONE"] == "Eastern Standard Time").all())   # from the calendar
        self.assertEqual(df.loc[2, "EXTRA_CREW"], 2)                          # row 4 has '+2'
        self.assertTrue(df.loc[0, "LINK"].startswith("https://"))             # URL behind the cell, by row

    def test_inserted_blank_row_keeps_row_numbers(self):
        wb = load_workbook(self.path)
        wb["By Day"].insert_rows(3)   # empty row: its ID formula is missing too
        wb.save(self.path)
        df = se.parse_lv_byDay(str(self.path), "By Day", "LV")
        self.assertEqual(list(df["EVENT_ID"][:3]), ["LV-0001", "LV-0003", "LV-0004"])  # Excel rows 2, 4, 5


class Template(unittest.TestCase):
    def test_dropdowns_row_colors_and_id_formula(self):
        wb = load_workbook(TEMPLATE)
        ws = wb["By Day"]
        self.assertEqual([c.value for c in ws[1]][6], "FACILITY")
        lists = {str(d.sqref): d.formula1 for d in ws.data_validations.dataValidation}
        self.assertIn("G2:G2000", lists)
        rules = {r.formula[0]: r.dxf.fill.end_color.rgb[-6:]
                 for rs in ws.conditional_formatting._cf_rules.values() for r in rs}
        expected = {f'$G2="{c}"': f["excel_color"] for c, f in FACILITIES.items() if f["excel_color"]}
        self.assertEqual(rules, expected)
        self.assertEqual(ws["A2000"].value, '=IF(C2000="","",Lists!$P$2&"-"&TEXT(ROW()-1,"0000"))')
        self.assertEqual((wb["Lists"]["P1"].value, wb["Lists"]["P2"].value), ("ID PREFIX", "LV"))
        text = " ".join(str(c.value or "") for c in wb["Instructions"]["A"])
        self.assertIn("ID fills in automatically", text)
        self.assertIn("The folder decides the Outlook calendar", text)
        self.assertNotIn("BRANCH / STATE / TIME ZONE", text)


class Sync(unittest.TestCase):
    """sync_rows / remove_deleted_shifts with Graph replaced by fakes."""

    def setUp(self):
        self.cal, self.calls, self.deleted = {}, [], []
        self.orig = {n: getattr(se, n) for n in ("build_calendar_map", "create_graph_event", "patch_graph_event",
                                                 "delete_graph_event")}
        se.build_calendar_map = lambda t, m, s, e: {k: dict(v) for k, v in self.cal.get(m, {}).items()}
        se.create_graph_event = lambda t, m, p: (self.calls.append(("create", m, p)),
                                                 types.SimpleNamespace(status_code=201))[1]
        se.patch_graph_event = lambda t, m, g, p: (self.calls.append(("patch", m, p)),
                                                   types.SimpleNamespace(status_code=200))[1]
        se.delete_graph_event = lambda t, m, g: (self.deleted.append((m, g)), types.SimpleNamespace(status_code=204))[1]

    def tearDown(self):
        for n, f in self.orig.items():
            setattr(se, n, f)

    def test_past_rows_are_not_pushed(self):
        df = pd.DataFrame([_shift("LV-1", PAST), _shift("LV-2"), _shift("LV-3", se.today().strftime("%Y-%m-%d"))])
        stats, results, bad = se.sync_rows("t", df)
        self.assertEqual((stats["created"], stats["past"]), (2, 1))
        self.assertEqual([r["ID"] for r in results], ["LV-2", "LV-3"])

    def test_update_skip_and_duplicates(self):
        df = pd.DataFrame([_shift("LV-1"), _shift("LV-2"), _shift("LV-3")])
        self.cal = {"test@x.com": {"LV-1": {"graph_id": "g1", "hash": se._field_hash(df.iloc[0])},
                                   "LV-2": {"graph_id": "g2", "hash": "old"}}}
        stats, _, bad = se.sync_rows("t", df, {"LV-3"})
        self.assertEqual((stats["skipped"], stats["updated"], stats["created"]), (1, 1, 0))
        self.assertEqual(bad, [("LV-3", "ID LV-3 is used by more than one row")])

    def test_payload_colors_and_time_zone(self):
        df = pd.DataFrame([_shift("ORL-1", calendar="ORL LOEWS", facility="HARROC"),
                           _shift("LV-1", function="STRIKE"), _shift("LV-2", function="PREP")])
        se.sync_rows("t", df)
        payloads = {p["subject"] + m: p for _, m, p in self.calls}
        cats = [p["categories"] for _, _, p in self.calls]
        self.assertIn(["Hard Rock"], cats)
        self.assertIn(["STRIKE"], cats)
        self.assertIn([], cats)
        self.assertTrue(all(p["start"]["timeZone"] == "Pacific Standard Time" for p in payloads.values()))

    def test_hash_unchanged_for_uncolored_rows(self):
        r = _shift("LV-1", function="PREP")
        old = hashlib.md5("|".join(str(r.get(f, "")) for f in
                                   ["DATE", "SHOW", "FUNCTION", "TIME", "LOCATION", "FOREMAN", "NOTES"]).encode()).hexdigest()[:8]
        self.assertEqual(se._field_hash(r), old)       # no needless Outlook updates
        colored = _shift("LV-1")                       # SET on LV: colored, so its hash includes the color
        plain = hashlib.md5("|".join(str(colored.get(f, "")) for f in
                                     ["DATE", "SHOW", "FUNCTION", "TIME", "LOCATION", "FOREMAN", "NOTES"]).encode()).hexdigest()[:8]
        self.assertNotEqual(se._field_hash(colored), plain)

    def files(self):
        return {"Las Vegas/a.xlsx": pd.DataFrame([_shift("LV-1"), _shift("LV-OLD", PAST)]),
                "Chicago/c.xlsx": pd.DataFrame([_shift("CHI-1", calendar="CHI")]),
                "Orlando - Lowes/o.xlsx": pd.DataFrame([_shift("ORL-1", calendar="ORL LOEWS")]),
                "Orlando Marriott/m.xlsx": pd.DataFrame([_shift("ORLM-1", calendar="ORL MARRIOTT")])}

    def test_removed_and_moved_to_past_are_deleted(self):
        self.cal = {"test@x.com": {"LV-1": {"graph_id": "g1"}, "LV-GONE": {"graph_id": "g2"},
                                   "LV-OLD": {"graph_id": "g3"}},   # row now dated in the past
                    "loews@x.com": {"ORL-1": {"graph_id": "g4"}}}
        removed, problem = se.remove_deleted_shifts("t", self.files(), {})
        self.assertEqual(problem, "")
        self.assertEqual(sorted(self.deleted), [("test@x.com", "g2"), ("test@x.com", "g3")])

    def test_no_deletes_when_a_file_cannot_be_read(self):
        self.cal = {"test@x.com": {"LV-GONE": {"graph_id": "g2"}}}
        removed, problem = se.remove_deleted_shifts("t", self.files(), {"Chicago/c.xlsx": "bad"})
        self.assertEqual((removed, self.deleted), ([], []))
        self.assertIn("could not read: Chicago/c.xlsx", problem)

    def test_calendar_without_files_protects_its_mailbox(self):
        files = self.files()
        del files["Chicago/c.xlsx"]   # Chicago folder empty: test@x.com (shared with LV) is not cleaned
        self.cal = {"test@x.com": {"LV-GONE": {"graph_id": "g2"}}, "loews@x.com": {"ORL-GONE": {"graph_id": "g5"}}}
        se.remove_deleted_shifts("t", files, {})
        self.assertEqual(self.deleted, [("loews@x.com", "g5")])

    def test_same_mailbox_under_two_addresses(self):
        ev = {"LV-1": {"graph_id": "g1"}, "LV-2": {"graph_id": "g2"}}
        self.cal = {"a@x.com": ev, "b@x.com": ev}
        se._delete_removed("t", {"LV-1": "b@x.com"}, {"a@x.com", "b@x.com"})
        self.assertEqual(self.deleted, [("a@x.com", "g2")])   # LV-1 kept; LV-2 deleted once, not twice

    def test_missing_permission_reports_but_does_not_raise(self):
        orig = se.requests
        se.requests = types.SimpleNamespace(get=lambda *a, **k: types.SimpleNamespace(status_code=403))
        try:
            err = se.ensure_outlook_categories("t", "loews@x.com")
        finally:
            se.requests = orig
        self.assertIn("MailboxSettings.ReadWrite", err)
        self.assertIn("Hard Rock (red)", err)


class OutlookView(unittest.TestCase):
    PAGE1 = {"value": [
        {"subject": "Show A - SET", "isAllDay": False, "categories": ["SET"],
         "start": {"dateTime": "2027-01-05T07:00:00.0000000"}, "end": {"dateTime": "2027-01-05T15:00:00.0000000"},
         "location": {"displayName": "Hall C"}, "lastModifiedDateTime": "2027-01-02T18:30:00.1234567Z",
         "body": {"content": "GES_ID:LV-0001|GES_HASH:abc\r\nForeman: Bob. Bring lifts"}}],
        "@odata.nextLink": "next"}
    PAGE2 = {"value": [
        {"subject": "Show B - STRIKE", "isAllDay": True, "categories": [],
         "start": {"dateTime": "2027-01-05T00:00:00.0000000"}, "end": {"dateTime": "2027-01-07T00:00:00.0000000"},
         "body": {"content": "typed in Outlook"}}]}

    def read(self, pages):
        calls, orig = [], se.requests
        def get(url, headers=None, params=None, timeout=None):
            calls.append((url, headers, params))
            p = pages.pop(0)
            return types.SimpleNamespace(status_code=p.get("status", 200), json=lambda: p)
        se.requests = types.SimpleNamespace(get=get)
        try:
            out = se.read_calendar_events("t", "lv@x.com", pd.Timestamp("2027-01-01"), pd.Timestamp("2027-01-31"),
                                          "Pacific Standard Time")
        finally:
            se.requests = orig
        return out, calls

    def test_reads_all_pages_in_the_calendar_time_zone(self):
        (events, err), calls = self.read([dict(self.PAGE1), dict(self.PAGE2)])
        self.assertEqual(err, "")
        self.assertEqual(len(calls), 2)
        self.assertIn('outlook.timezone="Pacific Standard Time"', calls[0][1]["Prefer"])
        self.assertIsNone(calls[1][2])  # the next page link already has the query
        a, b = events
        self.assertEqual((a["EVENT_ID"], a["NOTES"], a["LOCATION"]), ("LV-0001", "Foreman: Bob. Bring lifts", "Hall C"))
        self.assertEqual(a["START"], pd.Timestamp("2027-01-05 07:00"))
        self.assertEqual(str(a["MODIFIED"].tz), "UTC")
        self.assertEqual((b["EVENT_ID"], b["NOTES"], b["MODIFIED"]), ("", "typed in Outlook", None))

    def test_error_is_reported(self):
        (events, err), _ = self.read([{"status": 403}])
        self.assertEqual(events, [])
        self.assertIn("403", err)

    def test_events_on_day(self):
        (events, _), _ = self.read([dict(self.PAGE1), dict(self.PAGE2)])
        day5 = se.events_on_day(events, "2027-01-05")
        self.assertEqual([e["SUBJECT"] for e in day5], ["Show B - STRIKE", "Show A - SET"])  # all-day first
        self.assertEqual([e["SUBJECT"] for e in se.events_on_day(events, "2027-01-06")], ["Show B - STRIKE"])
        self.assertEqual(se.events_on_day(events, "2027-01-07"), [])  # all-day end is exclusive


class Overtime(unittest.TestCase):
    def ot(self, rows):
        return compute_overtime(pd.DataFrame(rows))

    def test_shift_span(self):
        self.assertEqual(shift_span("08:00 - 16:45"), (480, 8.75))
        self.assertEqual(shift_span("22:00 - 02:00")[1], 4.0)
        self.assertEqual(shift_span(""), (None, 0.0))

    def test_weekday_over_8(self):  # Wed 2027-01-06
        r = self.ot([_row("2027-01-06", "06:00 - 16:00", ["Ann"])]).iloc[0]
        self.assertEqual((r.HOURS, r.REG_HOURS, r.OT_HOURS), (10.0, 8.0, 2.0))
        self.assertTrue(r.OT_SHIFT)

    def test_weekday_exactly_8_no_ot(self):
        r = self.ot([_row("2027-01-06", "08:00 - 16:00", ["Ann"])]).iloc[0]
        self.assertEqual(r.OT_HOURS, 0.0)
        self.assertFalse(r.OT_SHIFT)

    def test_two_shifts_same_day_ot_on_later_shift(self):
        ps = self.ot([_row("2027-01-06", "14:00 - 18:00", ["Ann"], facility="HELIOS"),
                      _row("2027-01-06", "07:00 - 13:00", ["Ann"], facility="HARROC")])
        by_fac = ps.set_index("FACILITY")["OT_HOURS"]
        self.assertEqual(by_fac["HARROC"], 0.0)   # first 6 hrs
        self.assertEqual(by_fac["HELIOS"], 2.0)   # hrs 7-10 -> 2 OT
        self.assertEqual(ps["OT_HOURS"].sum(), 2.0)

    def test_saturday_sunday_all_hours(self):
        sat = self.ot([_row("2027-01-09", "06:00 - 16:00", ["Ann"])]).iloc[0]
        sun = self.ot([_row("2027-01-10", "08:00 - 12:00", ["Ann"])]).iloc[0]
        self.assertEqual((sat.OT_HOURS, sat.OT_REASON), (10.0, "Saturday"))  # not 10 + 2
        self.assertEqual((sun.OT_HOURS, sun.OT_REASON), (4.0, "Sunday"))

    def test_holidays(self):
        for d in HOLIDAYS["ORL"]:
            r = self.ot([_row(d, "08:00 - 12:00", ["Ann"])]).iloc[0]
            self.assertEqual((r.OT_HOURS, r.OT_REASON), (4.0, "Holiday"), d)

    def test_holiday_on_weekend_counted_once(self):
        HOLIDAYS["ORL"].append("2027-01-09")  # a Saturday
        try:
            r = self.ot([_row("2027-01-09", "06:00 - 16:00", ["Ann"])]).iloc[0]
        finally:
            _install()
        self.assertEqual((r.OT_HOURS, r.OT_REASON), (10.0, "Holiday"))

    def test_marriott_calendar_uses_orlando_rules(self):
        r = self.ot([_row("2027-01-09", "06:00 - 16:00", ["Ann"], facility="", calendar="ORL MARRIOTT")]).iloc[0]
        self.assertEqual((r.OT_HOURS, r.OT_REASON), (10.0, "Saturday"))

    def test_calendar_without_rules(self):
        r = self.ot([_row("2027-01-09", "06:00 - 16:00", ["Ann"], facility="SHECHI", calendar="CHI")]).iloc[0]
        self.assertEqual((r.OT_HOURS, r.STATUS), (0.0, "no OT rules"))

    def test_untimed_saturday_is_ot_shift_without_hours(self):
        r = self.ot([_row("2027-01-09", "", ["Ann"])]).iloc[0]
        self.assertTrue(r.OT_SHIFT)
        self.assertEqual((r.OT_HOURS, r.STATUS), (0.0, "no TIME"))

    def test_unnamed_crew_counted_per_shift(self):
        ps = self.ot([_row("2027-01-06", "07:00 - 13:00", [], extra=2),
                      _row("2027-01-06", "14:00 - 18:00", [], extra=1)])
        self.assertEqual(len(ps), 3)
        self.assertEqual(ps["OT_HOURS"].sum(), 0.0)  # not merged into one 10-hr day

    def test_team_summary(self):
        ps = self.ot([_row("2027-01-06", "06:00 - 16:00", ["Ann", "Bob"]),
                      _row("2027-01-09", "08:00 - 12:00", ["Ann"])])
        s = ot_summary(ps, "PERSON").set_index("PERSON")
        self.assertEqual(s.loc["Ann", "Shifts"], 2)
        self.assertEqual(s.loc["Ann", "OT Hours"], 6.0)
        self.assertEqual(s.loc["Ann", "OT Hours: Saturday"], 4.0)
        self.assertEqual(s.loc["Bob", "OT Hours"], 2.0)
        self.assertEqual(s["OT Hours"].sum(), ps["OT_HOURS"].sum())


if __name__ == "__main__":
    unittest.main()
