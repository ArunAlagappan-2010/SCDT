"""
One-time migration: adds the "House" column to an EXISTING database that was
created before House support was added, without losing any real data.

Rebuilds every sheet from scratch using the current build_excel_template.py
layout (so formulas/validations/styling are all correct and consistent),
then re-inserts whatever real rows were already in the old file.

Run:
    python migrate_add_house.py

Safe to run only once -- afterwards the file already has the House column and
build_excel_template's own sheets will match, so there's nothing left to migrate.
"""

import os
import shutil
from datetime import datetime

from openpyxl import load_workbook

import build_excel_template as bet

DB_PATH = bet.DB_PATH


def has_house_column(wb):
    ws = wb["MasterList"]
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    return "House" in header


def read_old_master_rows(wb):
    ws = wb["MasterList"]
    rows = []
    for r in range(2, ws.max_row + 1):
        sid = ws.cell(row=r, column=1).value
        if not sid:
            continue
        rows.append({
            "StudentID": sid,
            "Name": ws.cell(row=r, column=2).value,
            "Class": ws.cell(row=r, column=3).value,
            "Section": ws.cell(row=r, column=4).value,
            "RollNo": ws.cell(row=r, column=5).value,
        })
    return rows


def read_old_category_rows(wb, key):
    """Old layout: EntryID, Date, Name, Class, Section, <extra...>, Remarks, LoggedBy, Source"""
    ws = wb[key]
    cfg = bet.CATEGORY_SHEETS[key]
    old_headers = ["EntryID", "Date", "Name", "Class", "Section"] + cfg["extra_cols"] + ["Remarks", "LoggedBy", "Source"]

    rows = []
    for r in range(2, ws.max_row + 1):
        entry_id = ws.cell(row=r, column=1).value
        if not entry_id:
            continue
        row = {}
        for c, header in enumerate(old_headers, start=1):
            row[header] = ws.cell(row=r, column=c).value
        rows.append(row)
    return rows


def main():
    if not os.path.exists(DB_PATH):
        print(f"No database found at {DB_PATH} -- nothing to migrate.")
        return

    old_wb = load_workbook(DB_PATH)
    if has_house_column(old_wb):
        print("Database already has a House column -- nothing to migrate.")
        return

    backup_path = DB_PATH.replace(".xlsx", f"_backup_before_house_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy(DB_PATH, backup_path)
    print(f"Backed up existing database to {backup_path}")

    old_master_rows = read_old_master_rows(old_wb)
    old_category_rows = {key: read_old_category_rows(old_wb, key) for key in bet.CATEGORY_SHEETS}

    print(f"Preserving {len(old_master_rows)} master list student(s).")
    for key, rows in old_category_rows.items():
        print(f"Preserving {len(rows)} entr{'y' if len(rows) == 1 else 'ies'} in {key}.")

    new_wb = bet.Workbook()
    new_wb.remove(new_wb.active)

    master_ws = bet.build_master_list(new_wb)
    # Overwrite the fresh example rows with whatever was really in the old file
    # (falls back to the built-in example data + a blank House for genuinely
    # new students the old file didn't know about).
    example_house_by_id = {ex[0]: ex[5] for ex in [
        ("STU001", None, None, None, None, "Kailash"),
        ("STU002", None, None, None, None, "Aravali"),
        ("STU003", None, None, None, None, "Nilgiri"),
        ("STU004", None, None, None, None, "Vindhya"),
        ("STU005", None, None, None, None, "Kailash"),
    ]}
    for r, student in enumerate(old_master_rows, start=2):
        master_ws.cell(row=r, column=1, value=student["StudentID"])
        master_ws.cell(row=r, column=2, value=student["Name"])
        master_ws.cell(row=r, column=3, value=student["Class"])
        master_ws.cell(row=r, column=4, value=student["Section"])
        master_ws.cell(row=r, column=5, value=student["RollNo"])
        master_ws.cell(row=r, column=6, value=example_house_by_id.get(student["StudentID"], ""))
        for c in range(1, 7):
            master_ws.cell(row=r, column=c).border = bet.BORDER

    # Build a lookup so preserved category-sheet rows can get a real House
    # value from the (now House-aware) master list instead of a blank.
    house_by_name = {s["Name"]: (example_house_by_id.get(s["StudentID"], "") or "") for s in old_master_rows}

    for key, cfg in bet.CATEGORY_SHEETS.items():
        ws = bet.build_category_sheet(new_wb, key, cfg)
        cols = ["EntryID", "Date", "Name", "Class", "Section", "House"] + cfg["extra_cols"] + ["Remarks", "LoggedBy", "Source"]
        for r, old_row in enumerate(old_category_rows[key], start=2):
            values = dict(old_row)
            values["House"] = house_by_name.get(old_row.get("Name"), "")
            for c, header in enumerate(cols, start=1):
                ws.cell(row=r, column=c, value=values.get(header, ""))

    bet.build_dashboard(new_wb)
    for name in ["MasterList"] + list(bet.CATEGORY_SHEETS.keys()):
        new_wb[name].sheet_properties.tabColor = (
            bet.NAVY if name == "MasterList" else bet.CATEGORY_SHEETS[name]["color"]
        )

    new_wb.save(DB_PATH)
    print(f"Migration complete. {DB_PATH} now has the House column.")
    print("(Backup of the pre-migration file is kept alongside it, in case you need to compare.)")


if __name__ == "__main__":
    main()
