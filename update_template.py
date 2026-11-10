# -*- coding: utf-8 -*-
"""
Add / refresh the FACILITY column on the Excel scheduling template, using schedule_config.py:
  - FACILITY column on 'By Day' (inserted before LOCATION if missing) with a dropdown of facility codes
  - facility code table on the 'Lists' sheet (the dropdown's source)
  - conditional formatting that colors the whole row by facility (FACILITIES[...]['excel_color'])
Safe to run again after editing schedule_config.py (e.g. adding a color).

Usage:  python update_template.py ["path\\to\\file.xlsx"]     (default: the template in templates\\)
The file is overwritten; keep a copy first.
"""
import sys
from copy import copy
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from schedule_config import FACILITIES, BRANCHES

DEFAULT_PATH = r"templates\GES Electrical Show Schedule TEMPLATE.xlsx"
FAC_HEADER   = "FACILITY"
INSERT_AT    = "LOCATION"           # FACILITY goes just before this column
LISTS_COL    = 11                   # Lists!K: FACILITY CODE / NAME / CITY / STATE (K:N)
LISTS_ROWS   = 60                   # rows reserved for the facility table / dropdown source

def _header(ws):
    return [str(c.value).strip().upper() if c.value is not None else "" for c in ws[1]]

def _insert_facility_column(ws):
    """Insert an empty FACILITY column before LOCATION, keeping widths and hyperlinks in place."""
    col = _header(ws).index(INSERT_AT) + 1
    widths = {i: ws.column_dimensions[get_column_letter(i)].width for i in range(col, ws.max_column + 1)}
    ws.insert_cols(col)
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    ws.column_dimensions[get_column_letter(col)].width = 12
    for row in ws.iter_rows():               # openpyxl doesn't move a hyperlink's cell reference
        for cell in row:
            if cell.hyperlink is not None:
                cell.hyperlink.ref = cell.coordinate
    ws.cell(1, col, FAC_HEADER)
    return col

def _write_lists(lists):
    hdr = lists["C1"]                        # copy the existing BRANCH header style
    for r in range(1, LISTS_ROWS + 2):
        for c in range(LISTS_COL, LISTS_COL + 4):
            lists.cell(r, c).value = None
    for c, title in enumerate(["FACILITY CODE", "FACILITY NAME", "CITY", "STATE"]):
        cell = lists.cell(1, LISTS_COL + c, title)
        cell._style = copy(hdr._style)
    for r, (code, f) in enumerate(FACILITIES.items(), start=2):
        for c, val in enumerate([code, f["name"], f["city"] or "TBD", f["state"] or "TBD"]):
            lists.cell(r, LISTS_COL + c, val)
    lists.column_dimensions[get_column_letter(LISTS_COL)].width = 15
    lists.column_dimensions[get_column_letter(LISTS_COL + 1)].width = 36

def _add_branches(lists):
    """Add BRANCHES missing from the BRANCH / STATE / TIME ZONE table (Lists C:E); existing rows are kept."""
    have = {str(lists.cell(r, 3).value).strip().upper() for r in range(2, LISTS_ROWS + 2) if lists.cell(r, 3).value}
    row = 2
    for code, b in BRANCHES.items():
        if code in have:
            continue
        while lists.cell(row, 3).value:
            row += 1
        for c, val in enumerate([code, b["state"], b["time_zone"]]):
            lists.cell(row, 3 + c, val)._style = copy(lists.cell(2, 3 + c)._style)

def _set_validations(ws, fac_col, last_row):
    """Re-create the FUNCTION / LOCATION / FACILITY dropdowns. openpyxl can't keep the template's
    original FUNCTION / LOCATION dropdowns (an Excel extension), so they are added back here."""
    h = _header(ws)
    letters = {name: get_column_letter(h.index(name) + 1) for name in ("FUNCTION", "LOCATION", FAC_HEADER)}
    ours = set(letters.values())
    ws.data_validations.dataValidation = [
        dv for dv in ws.data_validations.dataValidation
        if not any(get_column_letter(r.min_col) in ours for r in dv.sqref.ranges)]
    k = get_column_letter(LISTS_COL)
    specs = [
        ("FUNCTION", "Lists!$A$2:$A$30", True,
         "Pick a FUNCTION from the list (add new ones on the Lists sheet)."),
        ("LOCATION", "Lists!$G$2:$G$60", False, None),
        (FAC_HEADER, f"Lists!${k}$2:${k}${LISTS_ROWS + 1}", True,
         "Pick a FACILITY code from the list (codes are on the Lists sheet)."),
    ]
    for name, src, strict, msg in specs:
        dv = DataValidation(type="list", formula1=src, allow_blank=True, showErrorMessage=strict,
                            errorTitle=name if strict else None, error=msg)
        dv.add(f"{letters[name]}2:{letters[name]}{last_row}")
        ws.add_data_validation(dv)

def _set_row_colors(ws, fac_col, last_row):
    """Whole-row fill by facility code; facilities without an excel_color stay uncolored."""
    ws.conditional_formatting._cf_rules.clear()
    rng = f"A2:{get_column_letter(ws.max_column)}{last_row}"
    letter = get_column_letter(fac_col)
    for code, f in FACILITIES.items():
        if f["excel_color"]:
            fill = PatternFill(start_color=f["excel_color"], end_color=f["excel_color"], fill_type="solid")
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${letter}2="{code}"'], fill=fill))

FAC_HELP = ("FACILITY - pick the facility code from the dropdown (list is on the 'Lists' sheet). It picks "
            "the Outlook calendar and colours the row. Leave it blank only if your city has no facilities: "
            "the row then goes to your city's calendar (from the ID prefix). If your city has no calendar "
            "set up, the row is not pushed to Outlook.")

def _update_instructions(ws):
    """Document the FACILITY column on the Instructions sheet (added once, text refreshed every run)."""
    for c in ws["A"]:
        if str(c.value or "").startswith(("FACILITY (", "FACILITY - ")):
            c.value = FAC_HELP
            return
    ws["A1"] = str(ws["A1"].value).replace("(v1.0)", "(v1.1)")
    for c in ws["A"]:
        if str(c.value or "").startswith("5. To cancel a shift"):
            c.value = ("5. To cancel a shift, delete the row. Do not use strikethrough or colour - the app "
                       "cannot see formatting. Row colours come automatically from FACILITY.")
        if str(c.value or "").startswith("FUNCTION (required)"):
            r = c.row + 1
            ws.insert_rows(r)
            ws.cell(r, 1, FAC_HELP)
            ws.cell(r, 1)._style = copy(c._style)
            break

def main(path):
    wb = load_workbook(path)
    ws = wb["By Day"]
    h = _header(ws)
    fac_col = h.index(FAC_HEADER) + 1 if FAC_HEADER in h else _insert_facility_column(ws)
    # Optional column (blank = city calendar), so it looks like LOCATION, not the red required headers
    ws.cell(1, fac_col)._style = copy(ws.cell(1, _header(ws).index(INSERT_AT) + 1)._style)
    last_row = ws.max_row
    _write_lists(wb["Lists"])
    _add_branches(wb["Lists"])
    _set_validations(ws, fac_col, last_row)
    _set_row_colors(ws, fac_col, last_row)
    _update_instructions(wb["Instructions"])
    wb.save(path)
    print(f"Updated {path}: FACILITY in column {get_column_letter(fac_col)}, rows 2-{last_row}.")

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PATH)
