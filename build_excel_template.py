"""
Generates the School Discipline Tracker Excel database (data/SchoolDisciplineSystem.xlsx).

Run this ONCE to create the workbook:
    python build_excel_template.py

If the file already exists, this script will refuse to overwrite it (to avoid
wiping real data). To move an EXISTING database onto this layout without
losing anything, run `python restyle_workbook.py` instead.

What the app relies on (do not change without changing app.py too):
  * sheet names: MasterList, LateComers, Defaulters, Uniform, NailsHair
  * header in row 1 of every data sheet, data from row 2
  * column order of every data sheet (see BASE_COLS / NAME_COLS / TAIL_COLS)
  * the StudentNames defined name
Everything else here is presentation and can be restyled freely.
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule, Rule
from openpyxl.formatting.rule import DataBar, FormatObject
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
DB_PATH = os.path.join(DATA_DIR, "SchoolDisciplineSystem.xlsx")

# --- Palette: same tokens as static/style.css -------------------------------
INK = "232323"          # headers, titles
INK_SOFT = "686868"     # notes, secondary text
FAINT = "9A9A97"
LINE = "E6E6E3"
SURFACE = "F4F4F2"      # zebra rows, auto-filled cells
MINT = "A1FFCB"         # the one accent
MINT_SOFT = "E6FBEF"
MINT_INK = "0F3D26"
WHITE = "FFFFFF"
INPUT_FILL = "FFFBE6"   # cells the user is meant to type into on the Dashboard

# Kept for migrate_add_house.py, which imports these names.
NAVY = INK

FONT_NAME = "Arial"

HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
HEADER_FILL = PatternFill("solid", fgColor=INK)
BODY_FONT = Font(name=FONT_NAME, size=10, color=INK)
AUTO_FONT = Font(name=FONT_NAME, size=10, color=INK_SOFT)
ID_FONT = Font(name="Consolas", size=10, color=INK_SOFT)
TITLE_FONT = Font(name=FONT_NAME, size=22, bold=True, color=INK)
SECTION_FONT = Font(name=FONT_NAME, size=12, bold=True, color=INK)
SUB_FONT = Font(name=FONT_NAME, size=9, color=INK_SOFT)
LABEL_FONT = Font(name=FONT_NAME, size=10, bold=True, color=INK)

HAIR = Side(style="thin", color=LINE)
MINT_RULE = Side(style="medium", color=MINT)
BORDER = Border(bottom=HAIR)                      # rows: a single hairline under each
HEADER_BORDER = Border(bottom=MINT_RULE)
SURFACE_FILL = PatternFill("solid", fgColor=SURFACE)
MINT_FILL = PatternFill("solid", fgColor=MINT)
MINT_SOFT_FILL = PatternFill("solid", fgColor=MINT_SOFT)
# Conditional formats read their fill from bgColor, not fgColor.
CF_SURFACE = PatternFill(bgColor=SURFACE)
CF_MINT = PatternFill(bgColor=MINT)
CF_MINT_SOFT = PatternFill(bgColor=MINT_SOFT)

DATE_FMT = "dd mmm yyyy"
LOG_ROWS = 1000           # matches app.MAX_LOG_ROWS
MASTER_ROWS = 2000        # matches app.MAX_MASTER_ROWS

# One green ramp, dark to mint -- the same colours app.py uses per category.
CATEGORY_SHEETS = {
    "LateComers": {"label": "Late Comers", "color": "123524", "light": "E3EEE8", "extra_cols": ["Time", "Reason"]},
    "Defaulters": {"label": "Defaulters", "color": "2E7D57", "light": "E6F2EC", "extra_cols": ["Reason"]},
    "Uniform": {"label": "Uniform", "color": "5DBB8C", "light": "EAF7F0", "extra_cols": ["Issue"]},
    "NailsHair": {"label": "Nails & Hair", "color": "A1E8C0", "light": "EEFBF4", "extra_cols": ["Issue"]},
}

BASE_COLS = ["EntryID", "Date"]
NAME_COLS = ["Name", "Class", "Section", "House"]
TAIL_COLS = ["Remarks", "LoggedBy", "Source"]

# Friendlier header text shown in Excel. Row-1 values must stay as the
# internal names, so these go in as cell comments, not as the header itself.
HEADER_HELP = {
    "EntryID": "Set by the web app. Leave blank when typing a row by hand.",
    "Date": "Date of the case, e.g. 18 Nov 2026.",
    "Name": "Pick from the dropdown. It lists the MasterList sheet.",
    "Class": "Fills in from the MasterList once a name is picked.",
    "Section": "Fills in from the MasterList once a name is picked.",
    "House": "Fills in from the MasterList once a name is picked.",
    "Source": "Manual or OCR (scanned register).",
    "LoggedBy": "Staff member who logged it.",
}

MASTER_HEADERS = ["StudentID", "Name", "Class", "Section", "RollNo", "House"]
# Live per-student counts to the right of the editable master columns. The app
# only ever reads/writes columns A-F, so these can't disturb it.
MASTER_COUNT_HEADERS = ["Late", "Defaulted", "Uniform", "Nails/Hair", "Total"]
MASTER_RANK_COL = 12  # L: hidden tie-broken sort key used by the Dashboard top-10

EXAMPLE_STUDENTS = [
    ("STU001", "Aarav Sharma", "9", "A", "1", "Kailash"),
    ("STU002", "Diya Patel", "9", "A", "2", "Aravali"),
    ("STU003", "Kabir Singh", "9", "B", "1", "Nilgiri"),
    ("STU004", "Meera Nair", "10", "A", "1", "Vindhya"),
    ("STU005", "Rohan Gupta", "10", "B", "3", "Kailash"),
]


def category_headers(cfg):
    return BASE_COLS + NAME_COLS + cfg["extra_cols"] + TAIL_COLS


def style_header_row(ws, row, ncols, fill_color=INK, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.border = HEADER_BORDER
    ws.row_dimensions[row].height = 26


def autosize(ws, widths, start_col=1):
    for i, w in enumerate(widths, start=start_col):
        ws.column_dimensions[get_column_letter(i)].width = w


def setup_print(ws, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.oddFooter.center.text = "&A  -  page &P of &N"
    ws.oddFooter.center.size = 8


def add_zebra(ws, ref, key_col="A"):
    """Banded rows that only appear once a row actually holds data."""
    first_row = int("".join(ch for ch in ref.split(":")[0] if ch.isdigit()))
    ws.conditional_formatting.add(
        ref,
        FormulaRule(formula=[f'AND(${key_col}{first_row}<>"",MOD(ROW(),2)=0)'], fill=CF_SURFACE, stopIfTrue=False),
    )


# ---------------------------------------------------------------------------
# MasterList
# ---------------------------------------------------------------------------

def build_master_list(wb, examples=True):
    ws = wb.create_sheet("MasterList")
    ws.sheet_view.showGridLines = False
    ws.append(MASTER_HEADERS + MASTER_COUNT_HEADERS)
    style_header_row(ws, 1, len(MASTER_HEADERS))
    style_header_row(ws, 1, len(MASTER_COUNT_HEADERS), fill_color="3A3A3A", start_col=len(MASTER_HEADERS) + 1)
    for c in range(len(MASTER_HEADERS) + 1, len(MASTER_HEADERS) + len(MASTER_COUNT_HEADERS) + 1):
        ws.cell(row=1, column=c).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=1, column=7).comment = Comment("Columns G-K count each student's entries live. Don't type here.", "Tracker")
    ws.cell(row=1, column=MASTER_RANK_COL, value="SortKey")

    autosize(ws, [13, 28, 10, 11, 11, 14, 10, 14, 13, 14, 10, 4])
    ws.column_dimensions[get_column_letter(MASTER_RANK_COL)].hidden = True
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(MASTER_HEADERS) + len(MASTER_COUNT_HEADERS))}1"
    ws.sheet_properties.tabColor = INK

    if examples:
        for row in EXAMPLE_STUDENTS:
            ws.append(row)

    last = MASTER_ROWS + 1
    count_sheets = list(CATEGORY_SHEETS)
    for r in range(2, last + 1):
        for c in range(1, len(MASTER_HEADERS) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.font = ID_FONT if c == 1 else BODY_FONT
        # Live counts: entries for this name in each register, then the total.
        for i, sheet in enumerate(count_sheets):
            cell = ws.cell(row=r, column=7 + i)
            cell.value = f'=IF($B{r}="","",COUNTIF({sheet}!$C$2:$C${LOG_ROWS + 1},$B{r}))'
            cell.font = AUTO_FONT
            cell.alignment = Alignment(horizontal="center")
            cell.border = BORDER
        total = ws.cell(row=r, column=11)
        total.value = f'=IF($B{r}="","",SUM(G{r}:J{r}))'
        total.font = Font(name=FONT_NAME, size=10, bold=True, color=INK)
        total.alignment = Alignment(horizontal="center")
        total.border = BORDER
        # Sort key: Total plus a tiny row-based tie-breaker so LARGE() never
        # returns two equal values (which would make MATCH pick one student twice).
        ws.cell(row=r, column=MASTER_RANK_COL).value = f'=IF(OR($B{r}="",N(K{r})=0),"",K{r}+(10000-ROW())/100000)'

    # ISNUMBER guard: Excel ranks the "" of an empty row above every number.
    ws.conditional_formatting.add(
        f"K2:K{last}", FormulaRule(formula=["AND(ISNUMBER(K2),K2>=3)"], fill=CF_MINT, font=Font(bold=True, color=MINT_INK))
    )
    ws.conditional_formatting.add(f"G2:J{last}", CellIsRule(operator="equal", formula=["0"], font=Font(color="C9C9C5")))

    add_zebra(ws, f"A2:K{last}")  # added last = lowest priority, so mint wins

    setup_print(ws, landscape=False)
    ws.print_title_rows = "1:1"

    # Named range covering plenty of headroom for new students, used by dropdowns
    # in the log sheets and by the web app.
    wb.defined_names["StudentNames"] = DefinedName(
        "StudentNames", attr_text=f"MasterList!$B$2:$B${MASTER_ROWS + 1}"
    )
    return ws


# ---------------------------------------------------------------------------
# Register sheets
# ---------------------------------------------------------------------------

def build_category_sheet(wb, key, cfg):
    ws = wb.create_sheet(key)
    ws.sheet_view.showGridLines = False
    headers = category_headers(cfg)
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    for c, h in enumerate(headers, start=1):
        if h in HEADER_HELP:
            ws.cell(row=1, column=c).comment = Comment(HEADER_HELP[h], "Tracker")
    ws.freeze_panes = "D2"   # EntryID, Date and Name stay put while scrolling right
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    ws.sheet_properties.tabColor = cfg["color"]

    extra_widths = [11 if c == "Time" else 26 for c in cfg["extra_cols"]]
    autosize(ws, [12, 14, 26, 10, 11, 14] + extra_widths + [32, 15, 11])

    name_col_idx = len(BASE_COLS) + 1  # Name is first of NAME_COLS
    name_col_letter = get_column_letter(name_col_idx)
    last = LOG_ROWS
    source_col_idx = len(headers)
    source_letter = get_column_letter(source_col_idx)
    remarks_idx = headers.index("Remarks") + 1

    # Dropdown on Name, sourced from the master list -- this is the
    # "type a name, pick the right one from the updated list" behaviour
    # when someone edits the sheet directly in Excel.
    dv = DataValidation(type="list", formula1="=StudentNames", allow_blank=True, showDropDown=False)
    dv.error = "Please pick a name from the master student list."
    dv.errorTitle = "Unknown student"
    dv.promptTitle = "Student"
    dv.prompt = "Pick a name. Class, section and house fill in."
    dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(f"{name_col_letter}2:{name_col_letter}{last}")

    dv_source = DataValidation(type="list", formula1='"Manual,OCR"', allow_blank=True)
    ws.add_data_validation(dv_source)
    dv_source.add(f"{source_letter}2:{source_letter}{last}")

    date_dv = DataValidation(type="date", operator="lessThanOrEqual", formula1="TODAY()", allow_blank=True)
    date_dv.error = "Enter a real date, not in the future."
    date_dv.errorTitle = "Check the date"
    ws.add_data_validation(date_dv)
    date_dv.add(f"B2:B{last}")

    for r in range(2, last + 1):
        # Auto-fill Class/Section/House from MasterList once a Name is chosen.
        ws.cell(row=r, column=name_col_idx + 1).value = (
            f'=IFERROR(VLOOKUP({name_col_letter}{r},MasterList!$B:$D,2,FALSE)&"","")'
        )
        ws.cell(row=r, column=name_col_idx + 2).value = (
            f'=IFERROR(VLOOKUP({name_col_letter}{r},MasterList!$B:$D,3,FALSE)&"","")'
        )
        ws.cell(row=r, column=name_col_idx + 3).value = (
            f'=IFERROR(VLOOKUP({name_col_letter}{r},MasterList!$B:$F,5,FALSE)&"","")'
        )
        style_log_row(ws, r, len(headers), name_col_idx, remarks_idx)

    ws.conditional_formatting.add(
        f"{source_letter}2:{source_letter}{last}",
        CellIsRule(operator="equal", formula=['"OCR"'], fill=CF_MINT_SOFT, font=Font(color=MINT_INK)),
    )

    add_zebra(ws, f"A2:{get_column_letter(len(headers))}{last}")  # last = lowest priority

    setup_print(ws)
    ws.print_title_rows = "1:1"
    return ws


def style_log_row(ws, r, ncols, name_col_idx, remarks_idx):
    """Formatting for one data row. Kept separate so restyle_workbook.py can
    reuse it. Values written later by the app keep these styles."""
    for c in range(1, ncols + 1):
        cell = ws.cell(row=r, column=c)
        cell.border = BORDER
        cell.font = BODY_FONT
        cell.alignment = Alignment(vertical="top")
    ws.cell(row=r, column=1).font = ID_FONT
    ws.cell(row=r, column=2).number_format = DATE_FMT
    ws.cell(row=r, column=2).alignment = Alignment(horizontal="left", vertical="top")
    ws.cell(row=r, column=name_col_idx).font = Font(name=FONT_NAME, size=10, bold=True, color=INK)
    for c in (name_col_idx + 1, name_col_idx + 2, name_col_idx + 3):
        ws.cell(row=r, column=c).font = AUTO_FONT
    ws.cell(row=r, column=remarks_idx).alignment = Alignment(vertical="top", wrap_text=True)


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

def _period_formula(sheet, period):
    rng = f"{sheet}!$B$2:$B${LOG_ROWS + 1}"
    if period == "today":
        return f"=COUNTIFS({rng},TODAY())"
    if period == "week":
        return f'=COUNTIFS({rng},">="&(TODAY()-WEEKDAY(TODAY(),2)+1),{rng},"<="&TODAY())'
    if period == "month":
        return f'=COUNTIFS({rng},">="&(EOMONTH(TODAY(),-1)+1),{rng},"<="&TODAY())'
    if period == "year":
        return f'=COUNTIFS({rng},">="&DATE(YEAR(TODAY()),1,1),{rng},"<="&TODAY())'
    return f"=COUNT({rng})"


def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws.sheet_properties.tabColor = MINT
    autosize(ws, [3, 22, 14, 14, 14, 14, 14, 3, 3])

    keys = list(CATEGORY_SHEETS)
    labels = [CATEGORY_SHEETS[k]["label"] for k in keys]
    first_cat_col = 3            # C
    total_col = first_cat_col + len(keys)  # G
    cat_letters = [get_column_letter(first_cat_col + i) for i in range(len(keys))]
    total_letter = get_column_letter(total_col)

    # --- Title band ---------------------------------------------------------
    for c in range(1, 20):  # across the tables and the chart
        ws.cell(row=1, column=c).fill = MINT_FILL
    ws.row_dimensions[1].height = 6
    ws["B3"] = "MCTM Discipline Dashboard"
    ws["B3"].font = TITLE_FONT
    ws.row_dimensions[3].height = 34
    ws["B4"] = '="Counts update by themselves from the register sheets. As of "&TEXT(TODAY(),"dd mmm yyyy")&"."'
    ws["B4"].font = SUB_FONT

    def table_header(row, first, cells):
        for i, text in enumerate(cells):
            cell = ws.cell(row=row, column=first + i, value=text)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.border = HEADER_BORDER
            cell.alignment = Alignment(horizontal="left" if i == 0 else "center", vertical="center", indent=1 if i == 0 else 0)
        ws.row_dimensions[row].height = 24

    def section_title(row, text, note=None):
        ws.cell(row=row, column=2, value=text).font = SECTION_FONT
        if note:
            ws.cell(row=row + 1, column=2, value=note).font = SUB_FONT

    def body_cell(row, col, value, bold=False, fmt=None, align="center", fill=None):
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = Font(name=FONT_NAME, size=11 if bold else 10, bold=bold, color=INK)
        cell.alignment = Alignment(horizontal=align, vertical="center", indent=1 if align == "left" else 0)
        cell.border = BORDER
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        return cell

    # --- 1. Cases by period -------------------------------------------------
    section_title(6, "Cases by period", "Every register, counted from its Date column.")
    hdr = 8
    table_header(hdr, 2, ["Period"] + labels + ["All registers"])
    periods = [("Today", "today"), ("This week", "week"), ("This month", "month"), ("This year", "year"), ("All time", "all")]
    period_rows = {}
    for i, (label, period) in enumerate(periods):
        r = hdr + 1 + i
        period_rows[period] = r
        highlight = period == "month"
        fill = MINT_SOFT_FILL if highlight else None
        body_cell(r, 2, label, bold=highlight, align="left", fill=fill)
        for j, key in enumerate(keys):
            body_cell(r, first_cat_col + j, _period_formula(key, period), bold=highlight, fmt="0;-0;\"–\"", fill=fill)
        body_cell(r, total_col, f"=SUM({cat_letters[0]}{r}:{cat_letters[-1]}{r})", bold=True, fmt="0;-0;\"–\"", fill=fill)
        ws.row_dimensions[r].height = 22

    # --- 2. Month by month (current year) -----------------------------------
    m_top = hdr + len(periods) + 3          # 16
    section_title(m_top, "Month by month", "This calendar year. The chart on the right draws from this table.")
    m_hdr = m_top + 2                       # 18
    table_header(m_hdr, 2, ["Month"] + labels + ["All registers"])
    month_first = m_hdr + 1
    for m in range(1, 13):
        r = m_hdr + m
        cell = body_cell(r, 2, f"=DATE(YEAR(TODAY()),{m},1)", align="left", fmt="mmmm")
        for j, key in enumerate(keys):
            rng = f"{key}!$B$2:$B${LOG_ROWS + 1}"
            body_cell(r, first_cat_col + j, f'=COUNTIFS({rng},">="&$B{r},{rng},"<="&EOMONTH($B{r},0))', fmt="0;-0;\"–\"")
        body_cell(r, total_col, f"=SUM({cat_letters[0]}{r}:{cat_letters[-1]}{r})", bold=True, fmt="0;-0;\"–\"")
    month_last = m_hdr + 12
    # Current month row picks up the mint highlight.
    ws.conditional_formatting.add(
        f"B{month_first}:{total_letter}{month_last}",
        FormulaRule(formula=[f"MONTH($B{month_first})=MONTH(TODAY())"], fill=CF_MINT_SOFT, font=Font(bold=True)),
    )
    ws.conditional_formatting.add(
        f"{total_letter}{month_first}:{total_letter}{month_last}",
        _data_bar("5DBB8C"),
    )
    tr = month_last + 1
    body_cell(tr, 2, "Year total", bold=True, align="left", fill=SURFACE_FILL)
    for col in range(first_cat_col, total_col + 1):
        L = get_column_letter(col)
        body_cell(tr, col, f"=SUM({L}{month_first}:{L}{month_last})", bold=True, fmt="0;-0;\"–\"", fill=SURFACE_FILL)

    # --- 3. Most flagged students (top 10, sorted live) ---------------------
    t_top = tr + 3
    section_title(t_top, "Most flagged students", "Top 10 by total entries, all time. Students on 3 or more turn mint.")
    t_hdr = t_top + 2
    # Student | Class | one column per register | Total  (B..H)
    ws.column_dimensions["H"].width = 10
    table_header(t_hdr, 2, ["Student", "Class"] + labels + ["Total"])
    rank_rng = f"MasterList!$L$2:$L${MASTER_ROWS + 1}"
    for i in range(10):
        r = t_hdr + 1 + i
        n = i + 1
        match = f"MATCH(LARGE({rank_rng},{n}),{rank_rng},0)"
        body_cell(r, 2, f'=IFERROR(INDEX(MasterList!$B$2:$B${MASTER_ROWS + 1},{match}),"")', align="left")
        ws.cell(row=r, column=2).font = Font(name=FONT_NAME, size=10, bold=True, color=INK)
        # "9-B", or just "9" when the student has no section.
        sec = f"INDEX(MasterList!$D$2:$D${MASTER_ROWS + 1},{match})"
        body_cell(r, 3, f'=IFERROR(INDEX(MasterList!$C$2:$C${MASTER_ROWS + 1},{match})&IF({sec}="","","-"&{sec}),"")')
        for j, src in enumerate(["G", "H", "I", "J", "K"]):
            body_cell(r, 4 + j, f'=IFERROR(INDEX(MasterList!${src}$2:${src}${MASTER_ROWS + 1},{match}),"")', bold=(src == "K"))
    t_last = t_hdr + 10
    ws.conditional_formatting.add(
        f"H{t_hdr + 1}:H{t_last}",
        FormulaRule(formula=[f"AND(ISNUMBER(H{t_hdr + 1}),H{t_hdr + 1}>=3)"], fill=CF_MINT, font=Font(bold=True, color=MINT_INK)),
    )
    ws.cell(row=t_last + 1, column=2, value="Empty rows mean fewer than 10 students have any entries yet.").font = SUB_FONT

    # --- 4. By house ---------------------------------------------------------
    h_top = t_last + 4
    section_title(h_top, "By house", "Type your school's house names in the yellow cells; the counts follow.")
    h_hdr = h_top + 2
    table_header(h_hdr, 2, ["House"] + labels + ["All registers"])
    houses = sorted({s[5] for s in EXAMPLE_STUDENTS})
    for i in range(6):
        r = h_hdr + 1 + i
        house = houses[i] if i < len(houses) else None
        c = body_cell(r, 2, house, align="left", fill=PatternFill("solid", fgColor=INPUT_FILL))
        if i == 0:
            c.comment = Comment("Edit these house names to match your school. Blank rows are ignored.", "Tracker")
        for j, key in enumerate(keys):
            body_cell(r, first_cat_col + j, f'=IF($B{r}="","",COUNTIF({key}!$F$2:$F${LOG_ROWS + 1},$B{r}))', fmt="0;-0;\"–\"")
        body_cell(r, total_col, f'=IF($B{r}="","",SUM({cat_letters[0]}{r}:{cat_letters[-1]}{r}))', bold=True, fmt="0;-0;\"–\"")
    h_last = h_hdr + 6
    ws.conditional_formatting.add(
        f"{total_letter}{h_hdr + 1}:{total_letter}{h_last}",
        _data_bar("2E7D57"),
    )

    # --- Charts --------------------------------------------------------------
    trend = BarChart()
    trend.type = "col"
    trend.grouping = "stacked"
    trend.overlap = 100
    trend.gapWidth = 60
    trend.title = "Cases per month, this year"
    trend.y_axis.title = None
    trend.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=LINE))
    trend.y_axis.numFmt = "0"
    trend.x_axis.number_format = "mmm"
    data = Reference(ws, min_col=first_cat_col, max_col=first_cat_col + len(keys) - 1, min_row=m_hdr, max_row=month_last)
    cats = Reference(ws, min_col=2, min_row=month_first, max_row=month_last)
    trend.add_data(data, titles_from_data=True)
    trend.set_categories(cats)
    for series, key in zip(trend.series, keys):
        series.graphicalProperties.solidFill = CATEGORY_SHEETS[key]["color"]
        series.graphicalProperties.line.noFill = True
    _tidy_chart(trend)
    trend.width = 17
    trend.height = 8.5
    ws.add_chart(trend, "J6")


    setup_print(ws)
    ws.print_area = f"A1:R{h_last + 1}"
    return ws


def _data_bar(color):
    """In-cell bar that is truly empty at zero (Excel's default draws a 10% stub)."""
    bar = DataBar(cfvo=[FormatObject(type="num", val=0), FormatObject(type="max")], color=color, showValue=True, minLength=0, maxLength=100)
    return Rule(type="dataBar", dataBar=bar)


def _tidy_chart(chart):
    chart.roundedCorners = False
    chart.style = 2
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if chart.legend is not None:
        chart.legend.position = "b"
        chart.legend.overlay = False
    if chart.title is not None:
        chart.title.overlay = False


# ---------------------------------------------------------------------------
# Read-me sheet
# ---------------------------------------------------------------------------

def build_guide(wb):
    ws = wb.create_sheet("How to use")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "C9C9C5"
    autosize(ws, [3, 26, 80])
    ws["B2"] = "How this workbook works"
    ws["B2"].font = TITLE_FONT
    ws.row_dimensions[2].height = 32
    rows = [
        ("Dashboard", "Live totals. Nothing to type except the yellow house names."),
        ("MasterList", "One row per student. Columns A-F are yours to edit; G-K count each student's entries."),
        ("LateComers, Defaulters, Uniform, NailsHair", "One row per case. Pick the Name from its dropdown and Class, Section and House fill in."),
        ("Grey text", "Filled in by a formula. Leave it alone."),
        ("Mint", "Something to look at: the current month, a scanned entry, or a student on 3+ entries."),
        ("Yellow cells", "The only cells on the Dashboard meant for typing."),
        ("Do not", "Rename sheets, move columns or insert rows above row 2. The web app reads this file by position."),
        ("Saving", "Close Excel before staff log entries in the web app, or the app can't save."),
    ]
    for i, (k, v) in enumerate(rows):
        r = 4 + i
        a = ws.cell(row=r, column=2, value=k)
        a.font = LABEL_FONT
        a.alignment = Alignment(vertical="top", wrap_text=True)
        b = ws.cell(row=r, column=3, value=v)
        b.font = BODY_FONT
        b.alignment = Alignment(vertical="top", wrap_text=True)
        for c in (a, b):
            c.border = BORDER
        ws.row_dimensions[r].height = 30
    ws["B7"].font = Font(name=FONT_NAME, size=10, bold=True, color=INK_SOFT)
    ws["B8"].fill = MINT_FILL
    ws["B9"].fill = PatternFill("solid", fgColor=INPUT_FILL)
    return ws


def build_workbook(examples=True):
    wb = Workbook()
    wb.remove(wb.active)
    build_master_list(wb, examples=examples)
    for key, cfg in CATEGORY_SHEETS.items():
        build_category_sheet(wb, key, cfg)
    build_dashboard(wb)  # inserted at index 0 so it opens first
    build_guide(wb)
    wb.active = 0
    return wb


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(DB_PATH):
        print(f"Refusing to overwrite existing database: {DB_PATH}")
        print("To move an existing database onto the new layout, run: python restyle_workbook.py")
        return
    wb = build_workbook()
    wb.save(DB_PATH)
    print(f"Created {DB_PATH}")


if __name__ == "__main__":
    main()
