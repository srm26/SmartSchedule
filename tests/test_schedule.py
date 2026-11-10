# -*- coding: utf-8 -*-
"""Run:  python -m unittest discover -s tests -v   (from the project folder)"""
import ast
import os
import re
import sys
import shutil
import tempfile
import unittest
import warnings
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from schedule_config import FACILITIES, HOLIDAYS                                   # noqa: E402
from schedule_logic import facility_route, compute_overtime, ot_summary, shift_span  # noqa: E402

TEMPLATE = ROOT / "templates" / "GES Electrical Show Schedule TEMPLATE.xlsx"
OLD_TEMPLATE = ROOT / "templates" / "_backup" / "GES Electrical Show Schedule TEMPLATE (v1.0).xlsx"
warnings.simplefilter("ignore")


def _load_app_functions(names, extra_ns=None):
    """Load selected functions/constants from app.py without running the Streamlit app
    (importing app.py would start the SharePoint sweep)."""
    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf8"))
    ns = {"pd": pd, "re": re, "os": os, "load_workbook": load_workbook,
          "facility_route": facility_route, **(extra_ns or {})}
    from schedule_logic import has_facility
    ns["has_facility"] = has_facility
    for node in tree.body:
        if isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name) and node.targets[0].id.isupper():
            try:
                exec(compile(ast.Module([node], []), "app.py", "exec"), ns)
            except Exception:
                pass
        if isinstance(node, ast.FunctionDef) and node.name in names:
            exec(compile(ast.Module([node], []), "app.py", "exec"), ns)
    return ns


def _row(date, time, crew, extra=0, facility="HARROC", eid=None, branch=""):
    return {"EVENT_ID": eid or f"T-{date}-{time}", "DATE": pd.Timestamp(date), "SHOW": "S", "FUNCTION": "SET",
            "TIME": time, "CREW_LIST": crew, "EXTRA_CREW": extra, "FACILITY": facility, "BRANCH": branch}


class FacilityRouting(unittest.TestCase):
    MBX = {"LOEWS_ELECTRICAL": "loews@x.com"}

    def test_all_loews_facilities_go_to_loews_calendar(self):
        for code in ("HARROC", "HELIOS", "PORBAY", "ROYPAC", "LSAPP"):
            self.assertEqual(facility_route(code, self.MBX), ("loews@x.com", ""))

    def test_lowercase_and_spaces_accepted(self):
        self.assertEqual(facility_route(" harroc ", self.MBX)[0], "loews@x.com")

    def test_blank_unknown_unmapped_are_rejected(self):
        for code, msg in [("", "blank"), (None, "blank"), ("XYZ", "not a known facility"),
                          ("MARWCR", "no mailbox set"), ("HILLV", "no mailbox set"),
                          ("SHECHI", "no mailbox set")]:
            mbx, err = facility_route(code, self.MBX)
            self.assertIsNone(mbx, code)
            self.assertIn(msg, err)

    def test_each_facility_routes_to_its_calendar_mailbox(self):
        mbx = {"LOEWS_ELECTRICAL": "loews@x.com", "ORL_MARRIOTT": "orl@x.com",
               "LV_ELECTRICAL": "lv@x.com", "CHI_ELECTRICAL": "chi@x.com"}
        self.assertEqual([facility_route(c, mbx)[0] for c in ("HARROC", "MARWCR", "HILLV", "SHECHI")],
                         ["loews@x.com", "orl@x.com", "lv@x.com", "chi@x.com"])

    def test_state_check_uses_branch(self):
        mbx = {"LOEWS_ELECTRICAL": "loews@x.com", "LV_ELECTRICAL": "lv@x.com"}
        self.assertEqual(facility_route("HARROC", mbx, "ORL"), ("loews@x.com", ""))
        self.assertEqual(facility_route("HILLV", mbx, "LV"), ("lv@x.com", ""))
        self.assertIsNone(facility_route("HARROC", mbx, "LV")[0])
        self.assertIsNone(facility_route("SHECHI", mbx, "ORLM")[0])
        self.assertEqual(facility_route("HARROC", mbx, ""), ("loews@x.com", ""))  # no prefix: can't check state

    def test_calendar_without_mailbox_setting_is_rejected(self):
        mbx, err = facility_route("HARROC", {})
        self.assertIsNone(mbx)
        self.assertIn("GRAPH_CALENDAR_MAILBOXES", err)


class AppRouteRows(unittest.TestCase):
    """_route_rows in app.py: per-row routing, no silent fallback, legacy files unchanged."""

    def setUp(self):
        self.ns = _load_app_functions(
            {"_mailbox_setting", "_mailbox_for", "_route_rows"},
            {"_graph_cfg": lambda: {"shared_mailbox": "default@x.com"}})
        os.environ["GRAPH_CALENDAR_MAILBOXES"] = "LOEWS_ELECTRICAL=loews@x.com"
        os.environ["GRAPH_BRANCH_MAILBOXES"] = "ORL=orl@x.com"

    def tearDown(self):
        os.environ.pop("GRAPH_CALENDAR_MAILBOXES", None)
        os.environ.pop("GRAPH_BRANCH_MAILBOXES", None)

    def test_mixed_facilities_in_one_file(self):
        df = pd.DataFrame([{"FACILITY": "HARROC", "BRANCH": "ORL"}, {"FACILITY": "SHECHI", "BRANCH": "ORL"},
                           {"FACILITY": "XYZ", "BRANCH": "ORL"}, {"FACILITY": "LSAPP", "BRANCH": "ORL"}])
        mbx, err = self.ns["_route_rows"](df)
        self.assertEqual(list(mbx), ["loews@x.com", None, None, "loews@x.com"])
        self.assertEqual([bool(e) for e in err], [False, True, True, False])

    def test_facility_from_another_state_is_rejected(self):
        df = pd.DataFrame([{"FACILITY": "HARROC", "BRANCH": "LV"},    # Florida facility in the LV sheet
                           {"FACILITY": "HILLV", "BRANCH": "ORL"},    # Nevada facility in the ORL sheet
                           {"FACILITY": "MARWCR", "BRANCH": "ORLM"}])  # Florida in a Florida sheet: OK
        mbx, err = self.ns["_route_rows"](df)
        self.assertEqual(list(mbx)[:2], [None, None])
        self.assertIn("HARROC is in FL, but this LV spreadsheet is for Nevada", err[0])
        self.assertIn("HILLV is in NV", err[1])
        self.assertNotIn("spreadsheet is for", err[2])  # same state: no state error

    def test_blank_facility_goes_to_listed_city_calendar_only(self):
        df = pd.DataFrame([{"FACILITY": "", "BRANCH": "ORL"},   # ORL listed -> city calendar
                           {"FACILITY": "", "BRANCH": "LV"},    # LV not listed -> exception, NOT default mailbox
                           {"FACILITY": "", "BRANCH": ""}])     # no prefix -> exception
        mbx, err = self.ns["_route_rows"](df)
        self.assertEqual(list(mbx), ["orl@x.com", None, None])
        self.assertEqual(err[0], "")
        self.assertIn("branch LV has no city calendar", err[1])
        self.assertIn("no branch prefix", err[2])
        self.assertNotIn("default@x.com", list(mbx))

    def test_file_without_facility_column_keeps_branch_routing(self):
        df = pd.DataFrame([{"FACILITY": None, "BRANCH": "ORL"}, {"FACILITY": None, "BRANCH": "LV"}])
        mbx, err = self.ns["_route_rows"](df)
        self.assertEqual(list(mbx), ["orl@x.com", "default@x.com"])
        self.assertEqual(list(err), ["", ""])


class Parser(unittest.TestCase):
    def setUp(self):
        self.ns = _load_app_functions({"_read_branch_time_zones", "_cell_text", "_read_link_urls",
                                       "parse_lv_byDay", "_make_event_id"})
        self.tmp = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)  # pandas may still hold the file open on Windows

    def test_template_with_facility(self):
        path = self.tmp / "t.xlsx"
        shutil.copy(TEMPLATE, path)
        wb = load_workbook(path)
        ws = wb["By Day"]
        for r, code in zip(range(2, 7), ["HARROC", "helios", "", "SHECHI", "LSAPP"]):
            ws.cell(r, 7).value = code or None
        wb.save(path)
        df = self.ns["parse_lv_byDay"](str(path), "By Day")
        self.assertEqual(list(df["FACILITY"]), ["HARROC", "HELIOS", "", "SHECHI", "LSAPP"])
        self.assertEqual(df.loc[2, "EXTRA_CREW"], 2)          # LV-0003 has '+2'
        self.assertTrue(df.loc[0, "LINK"].startswith("https://"))  # hyperlink still read after the column insert

    def test_marriott_and_chicago_prefixes_get_their_time_zones(self):
        path = self.tmp / "t2.xlsx"
        shutil.copy(TEMPLATE, path)
        wb = load_workbook(path)
        ws = wb["By Day"]
        ws["A2"], ws["G2"] = "ORLM-0001", "MARWCR"
        ws["A3"], ws["G3"] = "CHI-0001", "SHECHI"
        wb.save(path)
        df = self.ns["parse_lv_byDay"](str(path), "By Day").set_index("EVENT_ID")
        self.assertEqual(df.loc["ORLM-0001", "TIME_ZONE"], "Eastern Standard Time")
        self.assertEqual(df.loc["CHI-0001", "TIME_ZONE"], "Central Standard Time")

    @unittest.skipUnless(OLD_TEMPLATE.exists(), "v1.0 template backup not present")
    def test_old_template_without_facility_column(self):
        df = self.ns["parse_lv_byDay"](str(OLD_TEMPLATE), "By Day")
        self.assertTrue(df["FACILITY"].isna().all())


class Template(unittest.TestCase):
    def test_dropdown_and_row_colors(self):
        ws = load_workbook(TEMPLATE)["By Day"]
        header = [c.value for c in ws[1]]
        self.assertEqual(header[6], "FACILITY")
        lists = {str(d.sqref): d.formula1 for d in ws.data_validations.dataValidation}
        self.assertIn("G2:G2000", lists)
        rules = {r.formula[0]: r.dxf.fill.end_color.rgb[-6:]
                 for rs in ws.conditional_formatting._cf_rules.values() for r in rs}
        expected = {f'$G2="{c}"': f["excel_color"] for c, f in FACILITIES.items() if f["excel_color"]}
        self.assertEqual(rules, expected)
        for tbd in ("MARWCR", "HILLV", "SHECHI"):
            self.assertNotIn(f'$G2="{tbd}"', rules)


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

    def test_holiday_on_weekend_counted_once(self):  # 2027-12-25 isn't configured; use a fake weekend holiday
        import schedule_logic
        orig = schedule_logic.HOLIDAYS
        schedule_logic.HOLIDAYS = {"ORL": ["2027-01-09"]}  # a Saturday
        try:
            r = self.ot([_row("2027-01-09", "06:00 - 16:00", ["Ann"])]).iloc[0]
        finally:
            schedule_logic.HOLIDAYS = orig
        self.assertEqual((r.OT_HOURS, r.OT_REASON), (10.0, "Holiday"))

    def test_marriott_uses_orlando_rules(self):
        r = self.ot([_row("2027-01-09", "06:00 - 16:00", ["Ann"], facility="MARWCR")]).iloc[0]
        self.assertEqual((r.OT_HOURS, r.OT_REASON), (10.0, "Saturday"))

    def test_no_rules_region(self):
        r = self.ot([_row("2027-01-09", "06:00 - 16:00", ["Ann"], facility="SHECHI")]).iloc[0]
        self.assertEqual((r.OT_HOURS, r.STATUS), (0.0, "no OT rules"))

    def test_blank_facility_uses_branch_rules(self):
        r = self.ot([_row("2027-01-09", "06:00 - 16:00", ["Ann"], facility="", branch="ORL")]).iloc[0]
        self.assertEqual((r.OT_HOURS, r.OT_REASON), (10.0, "Saturday"))

    def test_legacy_file_uses_branch(self):
        r = self.ot([_row("2027-01-09", "06:00 - 16:00", ["Ann"], facility=None, branch="ORL")]).iloc[0]
        self.assertEqual(r.OT_HOURS, 10.0)

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


class OutlookColors(unittest.TestCase):
    """SET / STRIKE events get an Outlook category (color); other functions don't."""

    def setUp(self):
        import hashlib, types
        from schedule_config import OUTLOOK_CATEGORIES
        self.sent = []
        fake = types.SimpleNamespace(
            post=lambda url, headers=None, json=None, timeout=None: self.sent.append(("post", url, json)),
            patch=lambda url, headers=None, json=None, timeout=None: self.sent.append(("patch", url, json)),
            get=lambda url, headers=None, timeout=None: types.SimpleNamespace(status_code=403, json=lambda: {}))
        self.ns = _load_app_functions(
            {"_field_hash", "_event_categories", "create_graph_event", "patch_graph_event", "_parse_time_range",
             "_timed_end", "_make_body", "ensure_outlook_categories"},
            {"hashlib": hashlib, "OUTLOOK_CATEGORIES": OUTLOOK_CATEGORIES, "requests": fake,
             "timedelta": __import__("datetime").timedelta, "escape": __import__("html").escape})

    def row(self, fn):
        return {"DATE": "2027-01-06", "SHOW": "S", "FUNCTION": fn, "TIME": "08:00 - 12:00", "LOCATION": "",
                "FOREMAN": "TBD", "NOTES": ""}

    def test_categories_by_function(self):
        cats = self.ns["_event_categories"]
        self.assertEqual(cats(self.row("SET")), ["SET"])
        self.assertEqual(cats(self.row("STRIKE")), ["STRIKE"])
        self.assertEqual(cats(self.row("BOOTH CHECK")), [])
        self.assertEqual(cats(self.row("SET/STRIKE")), [])

    def test_hash_unchanged_for_uncolored_functions(self):
        import hashlib
        r = self.row("PREP")
        old = hashlib.md5("|".join(str(r.get(f, "")) for f in
                                   ["DATE", "SHOW", "FUNCTION", "TIME", "LOCATION", "FOREMAN", "NOTES"]).encode()).hexdigest()[:8]
        self.assertEqual(self.ns["_field_hash"](r), old)           # no needless Outlook updates
        self.assertNotEqual(self.ns["_field_hash"](self.row("SET")),
                            hashlib.md5("|".join(str(self.row("SET").get(f, "")) for f in
                                                 ["DATE", "SHOW", "FUNCTION", "TIME", "LOCATION", "FOREMAN", "NOTES"]).encode()).hexdigest()[:8])

    def test_payload_carries_categories(self):
        ts = pd.Timestamp("2027-01-06")
        self.ns["create_graph_event"]("t", "m@x.com", "S - SET", ts, ts, time_str="08:00 - 12:00", categories=["SET"])
        self.ns["patch_graph_event"]("t", "m@x.com", "gid", ts, ts, time_str="08:00 - 12:00")
        self.assertEqual(self.sent[0][2]["categories"], ["SET"])
        self.assertEqual(self.sent[1][2]["categories"], [])        # function changed away from SET: color removed

    def test_missing_permission_reports_but_does_not_raise(self):
        err = self.ns["ensure_outlook_categories"]("t", "m@x.com")
        self.assertIn("MailboxSettings.ReadWrite", err)
        self.assertIn("SET (green)", err)
        self.assertIn("STRIKE (red)", err)


if __name__ == "__main__":
    unittest.main()
