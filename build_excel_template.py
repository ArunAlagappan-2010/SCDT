"""
Generates the initial School Discipline Tracker Excel database (data/SchoolDisciplineSystem.xlsx).

Run this ONCE to create the workbook:
    python build_excel_template.py

If the file already exists, this script will refuse to overwrite it (to avoid
wiping real data) unless you delete/rename the existing file first.
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.workbook.defined_name import DefinedName

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
DB_PATH = os.path.join(DATA_DIR, "SchoolDisciplineSystem.xlsx")

NAVY = "1F2A44"
BLUE = "2F5C8A"
LIGHT_BLUE = "DCE6F1"
RED = "C0392B"
LIGHT_RED = "FADBD8"
ORANGE = "D68910"
LIGHT_ORANGE = "FDEBD0"
GREEN = "1E8449"
LIGHT_GREEN = "D5F5E3"
GREY = "808B96"
WHITE = "FFFFFF"

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color=WHITE)
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
TITLE_FONT = Font(name="Calibri", size=20, bold=True, color=NAVY)
SUB_FONT = Font(name="Calibri", size=11, italic=True, color=GREY)
THIN = Side(style="thin", color="D0D3D4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CATEGORY_SHEETS = {
    "LateComers": {
        "label": "Late Comers",
        "color": BLUE,
        "light": LIGHT_BLUE,
        "extra_cols": ["Time", "Reason"],
    },
    "Defaulters": {
        "label": "Defaulters",
        "color": RED,
        "light": LIGHT_RED,
        "extra_cols": ["Reason"],
    },
    "Uniform": {
        "label": "Uniform",
        "color": ORANGE,
        "light": LIGHT_ORANGE,
        "extra_cols": ["Issue"],
    },
    "NailsHair": {
        "label": "Nails & Hair",
        "color": GREEN,
        "light": LIGHT_GREEN,
        "extra_cols": ["Issue"],
    },
}

BASE_COLS = ["EntryID", "Date"]
NAME_COLS = ["Name", "Class", "Section", "House"]
TAIL_COLS = ["Remarks", "LoggedBy", "Source"]


def style_header_row(ws, row, ncols, fill_color):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_master_list(wb):
    ws = wb.create_sheet("MasterList")
    headers = ["StudentID", "Name", "Class", "Section", "RollNo", "House"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers), NAVY)
    autosize(ws, [12, 28, 10, 10, 10, 14])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:F1"

    # A handful of example rows so the dropdowns/autocomplete have something to show.
    examples = [
        ("STU001", "Aarav Sharma", "9", "A", "1", "Red"),
        ("STU002", "Diya Patel", "9", "A", "2", "Blue"),
        ("STU003", "Kabir Singh", "9", "B", "1", "Green"),
        ("STU004", "Meera Nair", "10", "A", "1", "Yellow"),
        ("STU005", "Rohan Gupta", "10", "B", "3", "Red"),
    ]
    for row in examples:
        ws.append(row)
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER

    # Named range covering plenty of headroom for new students, used by dropdowns
    # in the log sheets and by the web app.
    max_students = 2000
    wb.defined_names["StudentNames"] = DefinedName(
        "StudentNames", attr_text=f"MasterList!$B$2:$B${max_students + 1}"
    )
    return ws


def build_category_sheet(wb, key, cfg):
    ws = wb.create_sheet(key)
    headers = BASE_COLS + NAME_COLS + cfg["extra_cols"] + TAIL_COLS
    ws.append(headers)
    style_header_row(ws, 1, len(headers), cfg["color"])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    widths = [10, 12, 26, 8, 8, 14] + [16] * len(cfg["extra_cols"]) + [26, 14, 10]
    autosize(ws, widths)

    name_col_idx = len(BASE_COLS) + 1  # Name is first of NAME_COLS
    name_col_letter = get_column_letter(name_col_idx)

    # Dropdown on Name, sourced from the master list -- this is the
    # "type a name, pick the right one from the updated list" behaviour
    # when someone edits the sheet directly in Excel.
    dv = DataValidation(type="list", formula1="=StudentNames", allow_blank=True, showDropDown=False)
    dv.error = "Please pick a name from the master student list."
    dv.errorTitle = "Unknown student"
    ws.add_data_validation(dv)
    dv.add(f"{name_col_letter}2:{name_col_letter}1000")

    # Auto-fill Class/Section/House from MasterList once a Name is chosen.
    for r in range(2, 1001):
        ws.cell(row=r, column=name_col_idx + 1).value = (
            f'=IFERROR(VLOOKUP({name_col_letter}{r},MasterList!$B:$D,2,FALSE),"")'
        )
        ws.cell(row=r, column=name_col_idx + 2).value = (
            f'=IFERROR(VLOOKUP({name_col_letter}{r},MasterList!$B:$D,3,FALSE),"")'
        )
        ws.cell(row=r, column=name_col_idx + 3).value = (
            f'=IFERROR(VLOOKUP({name_col_letter}{r},MasterList!$B:$F,5,FALSE),"")'
        )

    source_col_idx = len(headers)
    source_letter = get_column_letter(source_col_idx)
    dv_source = DataValidation(type="list", formula1='"Manual,OCR"', allow_blank=True)
    ws.add_data_validation(dv_source)
    dv_source.add(f"{source_letter}2:{source_letter}1000")

    for r in range(2, 1001):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER
    return ws


def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    autosize(ws, [3, 20, 16, 16, 16, 16, 3])

    ws["B2"] = "School Discipline & Grooming Dashboard"
    ws["B2"].font = TITLE_FONT
    ws["B3"] = "Live counts pull from the MasterList / LateComers / Defaulters / Uniform / NailsHair sheets"
    ws["B3"].font = SUB_FONT

    kpi_specs = [
        ("Late Comers", "LateComers", BLUE, LIGHT_BLUE, "B"),
        ("Defaulters", "Defaulters", RED, LIGHT_RED, "D"),
        ("Uniform Issues", "Uniform", ORANGE, LIGHT_ORANGE, "F"),
        ("Nails & Hair", "NailsHair", GREEN, LIGHT_GREEN, "H"),
    ]

    header_row = 5
    label_row = 6
    today_row = 7
    week_row = 8
    month_row = 9
    total_row = 10

    ws.cell(row=header_row, column=2, value="KPI (auto-counted)").font = Font(bold=True, color=GREY, size=9)

    ws.cell(row=today_row, column=1, value="Today").font = Font(bold=True, size=9, color=GREY)
    ws.cell(row=week_row, column=1, value="This Week").font = Font(bold=True, size=9, color=GREY)
    ws.cell(row=month_row, column=1, value="This Month").font = Font(bold=True, size=9, color=GREY)
    ws.cell(row=total_row, column=1, value="All Time").font = Font(bold=True, size=9, color=GREY)

    for title, sheet, color, light, col_letter in kpi_specs:
        c = ws[f"{col_letter}{label_row}"]
        c.value = title
        c.font = Font(bold=True, size=12, color=WHITE)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(f"{col_letter}{label_row}:{col_letter}{label_row}")
        ws.row_dimensions[label_row].height = 22

        formulas = {
            today_row: f'=COUNTIFS({sheet}!$B:$B,TODAY())',
            week_row: f'=COUNTIFS({sheet}!$B:$B,">="&TODAY()-WEEKDAY(TODAY(),2)+1,{sheet}!$B:$B,"<="&TODAY())',
            month_row: f'=COUNTIFS({sheet}!$B:$B,">="&EOMONTH(TODAY(),-1)+1,{sheet}!$B:$B,"<="&TODAY())',
            total_row: f'=COUNTA({sheet}!$B$2:$B$1000)',
        }
        for r, formula in formulas.items():
            cell = ws[f"{col_letter}{r}"]
            cell.value = formula
            cell.fill = PatternFill("solid", fgColor=light)
            cell.font = Font(bold=(r == total_row), size=14 if r == total_row else 11, color=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER
            ws.row_dimensions[r].height = 20

    # Top repeat cases (by total logged entries across all four sheets)
    top_row = 13
    ws.cell(row=top_row, column=2, value="Most Frequently Flagged Students").font = Font(bold=True, size=13, color=NAVY)
    hdr_row = top_row + 1
    headers = ["Name", "Class", "Late Comers", "Defaulters", "Uniform", "Nails & Hair", "Total"]
    for i, h in enumerate(headers):
        cell = ws.cell(row=hdr_row, column=2 + i, value=h)
        cell.font = HEADER_FONT
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER

    first_data_row = hdr_row + 1
    max_rows = 15
    for i in range(max_rows):
        r = first_data_row + i
        # Lists students from MasterList in order and counts their entries live;
        # sort by the Total column in Excel to surface repeat cases first.
        src_row = 2 + i
        ws.cell(row=r, column=2, value=f'=IFERROR(MasterList!B{src_row},"")')
        ws.cell(row=r, column=3, value=f'=IFERROR(MasterList!C{src_row},"")')
        ws.cell(row=r, column=4, value=f'=IF($B{r}="","",COUNTIF(LateComers!$C:$C,$B{r}))')
        ws.cell(row=r, column=5, value=f'=IF($B{r}="","",COUNTIF(Defaulters!$C:$C,$B{r}))')
        ws.cell(row=r, column=6, value=f'=IF($B{r}="","",COUNTIF(Uniform!$C:$C,$B{r}))')
        ws.cell(row=r, column=7, value=f'=IF($B{r}="","",COUNTIF(NailsHair!$C:$C,$B{r}))')
        ws.cell(row=r, column=8, value=f'=IF($B{r}="","",SUM(D{r}:G{r}))')
        for c in range(2, 9):
            ws.cell(row=r, column=c).border = BORDER
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")

    note_row = first_data_row + max_rows + 1
    ws.cell(row=note_row, column=2,
            value="Tip: sort/filter this table manually by the Total column to see repeat cases first.").font = SUB_FONT

    # A simple monthly comparison chart across the four categories (uses the KPI "This Month" row).
    chart = BarChart()
    chart.title = "This Month: Cases by Category"
    chart.y_axis.title = "Cases"
    chart.style = 10
    ws["B18"] = "Category"
    ws["C18"] = "This Month"
    cat_names = ["Late Comers", "Defaulters", "Uniform Issues", "Nails & Hair"]
    for i, (name, col_letter) in enumerate(zip(cat_names, ["B", "D", "F", "H"])):
        ws.cell(row=19 + i, column=2, value=name)
        ws.cell(row=19 + i, column=3, value=f"={col_letter}{month_row}")

    data_ref = Reference(ws, min_col=3, min_row=18, max_row=22)
    cats_ref = Reference(ws, min_col=2, min_row=19, max_row=22)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 14
    chart.height = 8
    ws.add_chart(chart, "J5")

    ws.sheet_properties.tabColor = NAVY
    return ws


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(DB_PATH):
        print(f"Refusing to overwrite existing database: {DB_PATH}")
        print("Delete or rename it first if you really want a fresh template.")
        return

    wb = Workbook()
    # Remove the default sheet; we'll add our own in a controlled order.
    default_sheet = wb.active
    wb.remove(default_sheet)

    build_master_list(wb)
    for key, cfg in CATEGORY_SHEETS.items():
        build_category_sheet(wb, key, cfg)
    build_dashboard(wb)  # inserted at index 0 so it opens first

    for name in ["MasterList"] + list(CATEGORY_SHEETS.keys()):
        wb[name].sheet_properties.tabColor = (
            NAVY if name == "MasterList" else CATEGORY_SHEETS[name]["color"]
        )

    wb.save(DB_PATH)
    print(f"Created {DB_PATH}")


if __name__ == "__main__":
    main()
