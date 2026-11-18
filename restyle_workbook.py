"""
Moves an EXISTING database onto the current build_excel_template.py layout
(new Dashboard, live counts on MasterList, styled registers, "How to use" tab)
without losing any students or entries.

    python restyle_workbook.py            # restyles data/SchoolDisciplineSystem.xlsx
    python restyle_workbook.py other.xlsx # or a specific file

A timestamped backup copy is written next to the file first. Columns are
matched by their row-1 header name, so any earlier layout of this tracker
works. Close the file in Excel (and stop app.py) before running.
"""

import os
import sys
import shutil
from datetime import datetime

from openpyxl import load_workbook

import build_excel_template as bet


def _headers(ws):
    return [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]


def _rows(ws, key_header):
    headers = _headers(ws)
    if key_header not in headers:
        return []
    key_idx = headers.index(key_header)
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if key_idx >= len(values) or values[key_idx] in (None, ""):
            continue
        rows.append({h: v for h, v in zip(headers, values) if h})
    return rows


def _plain(value):
    # A formula left over in an old row (e.g. an un-overwritten VLOOKUP) is
    # dropped; the new sheet already has its own formula in that cell.
    return None if isinstance(value, str) and value.startswith("=") else value


def restyle(path):
    if not os.path.exists(path):
        print(f"No database at {path}.")
        return 1
    old = load_workbook(path)

    students = _rows(old["MasterList"], "StudentID") if "MasterList" in old.sheetnames else []
    entries = {
        key: (_rows(old[key], "EntryID") if key in old.sheetnames else [])
        for key in bet.CATEGORY_SHEETS
    }
    dash_houses = []
    if "Dashboard" in old.sheetnames:
        ws = old["Dashboard"]
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=2).value == "House":
                for rr in range(r + 1, r + 7):
                    v = ws.cell(row=rr, column=2).value
                    if v and not str(v).startswith("="):
                        dash_houses.append(v)
                break

    backup = path.replace(".xlsx", f"_backup_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    shutil.copy(path, backup)
    print(f"Backup: {backup}")

    wb = bet.build_workbook(examples=False)

    ms = wb["MasterList"]
    for r, s in enumerate(students, start=2):
        for c, h in enumerate(bet.MASTER_HEADERS, start=1):
            ms.cell(row=r, column=c).value = _plain(s.get(h))
    print(f"Kept {len(students)} student(s).")

    for key, cfg in bet.CATEGORY_SHEETS.items():
        ws = wb[key]
        headers = bet.category_headers(cfg)
        for r, row in enumerate(entries[key], start=2):
            for c, h in enumerate(headers, start=1):
                if h in row:
                    value = _plain(row[h])
                    if value is None and h in ("Class", "Section", "House"):
                        continue  # keep the new auto-fill formula
                    ws.cell(row=r, column=c).value = value
        print(f"Kept {len(entries[key])} {cfg['label']} entr{'y' if len(entries[key]) == 1 else 'ies'}.")

    # Carry over house names typed on a previous version of the new Dashboard,
    # or else seed them from the houses actually used on the master list.
    houses = dash_houses or sorted({str(s.get("House")).strip() for s in students if s.get("House")})
    dash = wb["Dashboard"]
    for r in range(1, dash.max_row + 1):
        if dash.cell(row=r, column=2).value == "House":
            for i in range(6):
                dash.cell(row=r + 1 + i, column=2).value = houses[i] if i < len(houses) else None
            break

    try:
        wb.save(path)
    except PermissionError:
        print("Couldn't save: the file is open in Excel. Close it and run this again (the backup is already made).")
        return 1
    print(f"Restyled {path}")
    return 0


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else bet.DB_PATH
    sys.exit(restyle(target))
