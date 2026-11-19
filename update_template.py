# -*- coding: utf-8 -*-
"""
Refresh a schedule file (or the template) from the admin workbook (see schedule_config.py):
  - FACILITY column on 'By Day' (inserted before LOCATION if missing) with a dropdown of facility codes
  - facility code table on the 'Lists' sheet (the dropdown's source)
  - conditional formatting that colors the whole row by facility (Facilities sheet, EXCEL COLOR)
  - ID column = formula <ID PREFIX>-<row number - 1>; the prefix is in Lists!P2 (one prefix per file)
  - Instructions sheet text
Safe to run again after changing a facility or color.

Usage:  python update_template.py "<admin workbook.xlsx>" ["<file.xlsx>"] [ID PREFIX]
  file: default the template in the templates folder.  ID PREFIX: default the file's current one (Lists!P2, or the
  prefix of its first ID). Saved by a script, the IDs have no stored value until Excel opens the file; the
  app works out the same IDs from the row numbers, so that is fine.
The file is overwritten; keep a copy first.
"""
import sys
from copy import copy
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import schedule_config
from schedule_config import FACILITIES

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

FAC_HELP = ("FACILITY - pick the facility code from the dropdown (list is on the 'Lists' sheet). On the Orlando "
            "calendars it sets the Outlook color, and it colours the row. It does not pick the calendar: every row "
            "goes to the calendar of the SharePoint folder the file is saved in.")

# Instructions lines replaced on every run: {start of the old line: new text}
INSTRUCTION_TEXT = {
    "3. ID ":            "3. ID fills in automatically: the file's ID PREFIX (Lists sheet, cell P2) + the row number, "
                         "e.g. LV-0001. Do not type over it.",
    "   Changing or deleting an ID": "   Add new shifts at the bottom. Inserting, deleting or sorting rows renumbers the rows "
                         "below; the app then updates those Outlook events. An empty ID cell is fine - the app works "
                         "out the same ID from the row number.",
    "   The prefix (":   "   Each file has its own ID PREFIX. Two files must never use the same prefix "
                         "(the app holds back rows with duplicate IDs).",
    "   that is how the app knows": "   The time zone comes from the calendar (admin workbook), not from this file.",
    "ID (required)":     "ID - fills in automatically (formula). Do not type over it.",
    "Name: ":            "Name: <YEAR> GES Electrical Show Schedule <CITY>.xlsx   "
                         "(e.g. 2027 GES Electrical Show Schedule ORL Loews.xlsx)",
    "One file per city": "Save the file in your calendar's folder under ElectricalSchedule (Chicago, Las Vegas, "
                         "Orlando - Lowes, Orlando Marriott). The folder decides the Outlook calendar.",
    "Save to the shared SharePoint": "Shifts dated before today are not sent to Outlook. Other sheets in the workbook "
                         "are fine; the app only reads 'By Day'.",
}

def _update_instructions(ws):
    """Document the FACILITY column and the ID formula on the Instructions sheet (text refreshed every run)."""
    if not any(str(c.value or "").startswith(("FACILITY (", "FACILITY - ")) for c in ws["A"]):
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
    ws["A1"] = str(ws["A1"].value).replace("(v1.1)", "(v1.2)")
    for c in ws["A"]:
        text = str(c.value or "")
        if text.startswith(("FACILITY (", "FACILITY - ")):
            c.value = FAC_HELP
        for start, new in INSTRUCTION_TEXT.items():
            if text.startswith(start):
                c.value = new
                break

PREFIX_CELL = "P2"                   # Lists!P1 = 'ID PREFIX' header, P2 = this file's prefix

def _set_id_formula(ws, lists, prefix):
    """ID column: =IF(DATE="","",Lists!$P$2&"-"&TEXT(ROW()-1,"0000")) on every row; Lists!P2 holds the prefix."""
    lists["P1"] = "ID PREFIX"
    lists["P1"]._style = copy(lists["C1"]._style)
    lists[PREFIX_CELL] = prefix
    lists.column_dimensions["P"].width = 12
    h = _header(ws)
    id_col, date_col = h.index("ID") + 1, get_column_letter(h.index("DATE") + 1)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, id_col).value = f'=IF({date_col}{r}="","",Lists!$P$2&"-"&TEXT(ROW()-1,"0000"))'

def _current_prefix(ws, lists):
    if lists[PREFIX_CELL].value:
        return str(lists[PREFIX_CELL].value).strip()
    id_col = _header(ws).index("ID") + 1
    first = next((str(ws.cell(r, id_col).value) for r in range(2, ws.max_row + 1)
                  if ws.cell(r, id_col).value and not str(ws.cell(r, id_col).value).startswith("=")), "")
    return first.rsplit("-", 1)[0] if "-" in first else ""

def main(path, prefix=""):
    wb = load_workbook(path)
    ws = wb["By Day"]
    prefix = prefix or _current_prefix(ws, wb["Lists"])
    if not prefix:
        sys.exit(f"{path}: no ID prefix found; give one, e.g. LV")
    h = _header(ws)
    fac_col = h.index(FAC_HEADER) + 1 if FAC_HEADER in h else _insert_facility_column(ws)
    # Optional column (blank = city calendar), so it looks like LOCATION, not the red required headers
    ws.cell(1, fac_col)._style = copy(ws.cell(1, _header(ws).index(INSERT_AT) + 1)._style)
    last_row = ws.max_row
    _write_lists(wb["Lists"])
    _set_validations(ws, fac_col, last_row)
    _set_row_colors(ws, fac_col, last_row)
    _set_id_formula(ws, wb["Lists"], prefix)
    _update_instructions(wb["Instructions"])
    wb.save(path)
    print(f"Updated {path}: FACILITY in column {get_column_letter(fac_col)}, ID prefix {prefix}, rows 2-{last_row}.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    schedule_config.apply(sys.argv[1])
    for w in schedule_config.CONFIG_WARNINGS:
        print("WARNING:", w)
    main(sys.argv[2] if len(sys.argv) > 2 else DEFAULT_PATH, sys.argv[3] if len(sys.argv) > 3 else "")
