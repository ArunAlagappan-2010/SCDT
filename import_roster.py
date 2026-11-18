"""
Loads the school's student roster (one sheet per grade, columns "Student Name"
and "House") into the MasterList of the tracker's database.

    python import_roster.py "C:\\path\\to\\SCHOOL RECORD.xlsx"            # import
    python import_roster.py "C:\\path\\to\\SCHOOL RECORD.xlsx" --dry-run  # just show what would happen
    python import_roster.py roster.xlsx --db other.xlsx                  # a different database

Close the database in Excel and stop the server first. A timestamped backup of
the database is saved next to it before anything is written.

What it does
  * Student numbers: STU001, STU002, ... grade by grade, in the order of the file.
    Students already in the MasterList keep their number (matched by name and class).
  * Class / section: "Grade 9" -> class 9; "Grade IB-1A" -> class IB-1, section A.
  * Houses are cleaned to one spelling each (Vindyas -> Vindhyas, Aravalis -> Aravallis,
    stray spaces removed). An unknown spelling is kept as typed and reported.
  * Two students with the same name get their class added, e.g. "Samyuktha (7)" and
    "Samyuktha (9)": the app links register entries to students by name.
  * The five built-in example students are removed, unless they already have entries.
  * Students in the MasterList who are not in the roster are kept and listed, never
    deleted (they may have entries).
  * The Dashboard's "By house" table is set to the roster's houses.
"""

import argparse
import os
import re
import shutil
import sys
from datetime import datetime

from openpyxl import load_workbook

import build_excel_template as bet

CANONICAL_HOUSES = ["Aravallis", "Kailash", "Nilgiris", "Vindhyas"]
HOUSE_ALIASES = {
    "aravallis": "Aravallis", "aravalis": "Aravallis", "aravalli": "Aravallis", "aravali": "Aravallis",
    "kailash": "Kailash", "kailas": "Kailash",
    "nilgiris": "Nilgiris", "nilgiri": "Nilgiris",
    "vindhyas": "Vindhyas", "vindyas": "Vindhyas", "vindhya": "Vindhyas", "vindya": "Vindhyas",
}
MASTER_COLS = 6  # StudentID, Name, Class, Section, RollNo, House


def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def clean_house(value, unknown):
    text = clean_text(value)
    if not text:
        return ""
    house = HOUSE_ALIASES.get(text.lower().rstrip("."))
    if house is None:
        unknown.add(text)
        return text
    return house


def class_from_sheet(title):
    """'Grade 9' -> ('9', ''), 'Grade IB-1A' -> ('IB-1', 'A'), 'Grade IB-2' -> ('IB-2', '')."""
    label = re.sub(r"(?i)^\s*(grade|class|std\.?)\s*", "", title).strip()
    m = re.fullmatch(r"(?i)(IB)\s*-?\s*(\d+)\s*([A-Z])?", label)
    if m:
        return f"IB-{m.group(2)}", (m.group(3) or "").upper()
    m = re.fullmatch(r"(\d+)\s*-?\s*([A-Z])?", label, flags=re.I)
    if m:
        return m.group(1), (m.group(2) or "").upper()
    return label, ""


def read_roster(path):
    """[{name, class, section, house, sheet}] in file order, plus problems found."""
    wb = load_workbook(path, data_only=True)
    students, unknown_houses, no_house, skipped_sheets = [], set(), [], []
    for ws in wb:
        headers = [clean_text(ws.cell(row=1, column=c).value).lower() for c in range(1, ws.max_column + 1)]
        name_col = next((i + 1 for i, h in enumerate(headers) if "name" in h), None)
        house_col = next((i + 1 for i, h in enumerate(headers) if "house" in h), None)
        if not name_col:
            skipped_sheets.append(ws.title)
            continue
        klass, section = class_from_sheet(ws.title)
        for r in range(2, ws.max_row + 1):
            name = clean_text(ws.cell(row=r, column=name_col).value)
            if not name:
                continue
            house = clean_house(ws.cell(row=r, column=house_col).value, unknown_houses) if house_col else ""
            if not house:
                no_house.append(f"{name} ({ws.title})")
            students.append({"name": name, "class": klass, "section": section, "house": house, "sheet": ws.title})
    return students, sorted(unknown_houses), no_house, skipped_sheets


def label_of(s):
    return f"{s['class']}{s['section']}" if s["class"].startswith("IB") else (
        f"{s['class']}-{s['section']}" if s["section"] else s["class"])


def make_names_unique(students):
    """Adds the class to every name that appears more than once."""
    counts = {}
    for s in students:
        counts[s["name"].lower()] = counts.get(s["name"].lower(), 0) + 1
    renamed = []
    for s in students:
        if counts[s["name"].lower()] > 1:
            new = f"{s['name']} ({label_of(s)})"
            renamed.append((s["name"], new))
            s["name"] = new
    return renamed


def register_names(wb):
    names = set()
    for key in bet.CATEGORY_SHEETS:
        if key not in wb.sheetnames:
            continue
        ws = wb[key]
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value and ws.cell(row=r, column=3).value:
                names.add(clean_text(ws.cell(row=r, column=3).value).lower())
    return names


def read_master(ws):
    rows = []
    for r in range(2, ws.max_row + 1):
        sid = ws.cell(row=r, column=1).value
        name = ws.cell(row=r, column=2).value
        if sid in (None, "") or name in (None, ""):
            continue
        rows.append({
            "id": str(sid), "name": clean_text(name),
            "class": clean_text(ws.cell(row=r, column=3).value),
            "section": clean_text(ws.cell(row=r, column=4).value),
            "roll": ws.cell(row=r, column=5).value,
            "house": clean_text(ws.cell(row=r, column=6).value),
        })
    return rows


def plan_import(roster, existing, logged_names):
    """Works out the final MasterList. Returns (final_rows, report)."""
    examples = {(e[0], e[1]) for e in bet.EXAMPLE_STUDENTS}
    report = {"kept_ids": 0, "new": 0, "examples_removed": [], "not_in_roster": []}

    def key(name, klass):
        return (name.lower(), klass.lower())

    def is_unused_example(e):
        return (e["id"], e["name"]) in examples and e["name"].lower() not in logged_names

    # An untouched example student is never "the same person" as a real student
    # who happens to share the name, so they can't hand over their number.
    matchable = [e for e in existing if not is_unused_example(e)]
    by_name_class = {key(e["name"], e["class"]): e for e in matchable}
    by_name = {}
    for e in matchable:
        by_name.setdefault(e["name"].lower(), []).append(e)

    used = set()
    final = []
    for s in roster:
        match = by_name_class.get(key(s["name"], s["class"]))
        if match is None and len(by_name.get(s["name"].lower(), [])) == 1:
            match = by_name[s["name"].lower()][0]
        if match is not None and match["id"] not in used:
            used.add(match["id"])
            final.append({**s, "id": match["id"], "roll": match.get("roll")})
            report["kept_ids"] += 1
        else:
            final.append({**s, "id": None, "roll": None})
            report["new"] += 1

    leftovers = []
    for e in existing:
        if e["id"] in used:
            continue
        if is_unused_example(e):
            report["examples_removed"].append(e["name"])
            continue
        leftovers.append(e)
        report["not_in_roster"].append(f"{e['id']} {e['name']}")

    # Numbers: keep existing ones; give new students the next free STU numbers,
    # in roster order. (With only example students before, this starts at STU001.)
    taken = {r["id"] for r in final if r["id"]} | {e["id"] for e in leftovers}
    nums = [int(i[3:]) for i in taken if re.fullmatch(r"STU\d+", i)]
    counter = max(nums, default=0)
    for r in final:
        if not r["id"]:
            counter += 1
            r["id"] = f"STU{counter:03d}"
    return final + leftovers, report


def write_master(ws, rows):
    last = bet.MASTER_ROWS + 1
    for r in range(2, last + 1):
        for c in range(1, MASTER_COLS + 1):
            ws.cell(row=r, column=c).value = None
    for i, s in enumerate(rows, start=2):
        ws.cell(row=i, column=1).value = s["id"]
        ws.cell(row=i, column=2).value = s["name"]
        ws.cell(row=i, column=3).value = s["class"] or None
        ws.cell(row=i, column=4).value = s["section"] or None
        ws.cell(row=i, column=5).value = s.get("roll") or None
        ws.cell(row=i, column=6).value = s["house"] or None


def set_dashboard_houses(wb, houses):
    if "Dashboard" not in wb.sheetnames:
        return False
    ws = wb["Dashboard"]
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=2).value == "House":
            for i in range(6):
                ws.cell(row=r + 1 + i, column=2).value = houses[i] if i < len(houses) else None
            return True
    return False


def main(argv=None):
    ap = argparse.ArgumentParser(description="Load the school roster into the tracker's MasterList.")
    ap.add_argument("roster", help="the roster workbook, e.g. SCHOOL RECORD.xlsx")
    ap.add_argument("--db", default=bet.DB_PATH, help="the tracker database (default: data/SchoolDisciplineSystem.xlsx)")
    ap.add_argument("--dry-run", action="store_true", help="show what would change without writing anything")
    args = ap.parse_args(argv)

    if not os.path.exists(args.roster):
        print(f"Roster not found: {args.roster}")
        return 1
    if not os.path.exists(args.db):
        print(f"Database not found: {args.db}\nCreate it first: python build_excel_template.py")
        return 1

    roster, unknown_houses, no_house, skipped = read_roster(args.roster)
    if not roster:
        print("No students found in the roster (each sheet needs a 'Student Name' column).")
        return 1
    renamed = make_names_unique(roster)

    wb = load_workbook(args.db)
    if "MasterList" not in wb.sheetnames:
        print("That database has no MasterList sheet.")
        return 1
    existing = read_master(wb["MasterList"])
    final, report = plan_import(roster, existing, register_names(wb))
    if len(final) > bet.MASTER_ROWS:
        print(f"Too many students ({len(final)}); the MasterList holds {bet.MASTER_ROWS}.")
        return 1

    by_class = {}
    for s in roster:
        by_class[label_of(s)] = by_class.get(label_of(s), 0) + 1
    houses_found = sorted({s["house"] for s in roster if s["house"]})

    print(f"Roster: {len(roster)} students from {len(by_class)} classes")
    print("  " + ", ".join(f"{k}: {v}" for k, v in by_class.items()))
    print(f"Houses: {', '.join(houses_found)}")
    if unknown_houses:
        print(f"  Kept as typed (not a known house spelling): {', '.join(unknown_houses)}")
    if no_house:
        print(f"  No house given: {', '.join(no_house)}")
    for old, new in renamed:
        print(f"  Same name twice, stored as: {new}")
    if skipped:
        print(f"  Skipped sheets with no 'Name' column: {', '.join(skipped)}")
    print(f"MasterList: {report['kept_ids']} kept their student number, {report['new']} new"
          f" -> {final[0]['id']} .. {max((r['id'] for r in final), key=lambda i: (len(i), i))}")
    if report["examples_removed"]:
        print(f"  Example students removed: {', '.join(report['examples_removed'])}")
    if report["not_in_roster"]:
        print(f"  Kept, but not in the roster: {', '.join(report['not_in_roster'])}")

    if args.dry_run:
        print("\nDry run: nothing was written.")
        return 0

    backup = args.db.replace(".xlsx", f"_backup_before_roster_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    shutil.copy(args.db, backup)
    write_master(wb["MasterList"], final)
    set_dashboard_houses(wb, [h for h in CANONICAL_HOUSES if h in houses_found] +
                         [h for h in houses_found if h not in CANONICAL_HOUSES])
    try:
        wb.save(args.db)
    except PermissionError:
        print("\nCouldn't save: the database is open in Excel. Close it and run this again.")
        return 1
    print(f"\nSaved {args.db}\nBackup of the previous version: {backup}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
