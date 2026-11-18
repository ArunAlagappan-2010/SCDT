"""
School Discipline & Grooming Tracker
-------------------------------------
A small Flask app that uses the Excel workbook at data/SchoolDisciplineSystem.xlsx
as its database. Meant to run on ONE PC (the "server") on the school LAN; every
other classroom PC just opens http://<server-ip>:5000 in a browser.

Run:
    python app.py

See README.md for full setup (including the Tesseract OCR engine install).
"""

import io
import os
import re
import json
import time
import base64
import difflib
import shutil
import threading
from collections import Counter
from datetime import datetime, date, timedelta, time as clock_time

import requests
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, jsonify, flash, send_file
from openpyxl import load_workbook
from PIL import Image, ImageOps

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "SchoolDisciplineSystem.xlsx")
UPLOAD_DIR = os.path.join(DATA_DIR, "scans")
os.makedirs(UPLOAD_DIR, exist_ok=True)

load_dotenv(os.path.join(BASE_DIR, ".env"))

app = Flask(__name__)
app.secret_key = "school-discipline-tracker-local-only"

wb_lock = threading.RLock()

# --- OCR backend #1: Tesseract (offline, free, default) ---------------------
# pytesseract is only a Python wrapper: the Tesseract program itself must be
# installed too (Setup MCTM.cmd does it). Look in the usual install folders,
# then on PATH; OCR counts as available only when the program is really there.
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TESSERACT_CANDIDATES = [
    TESSERACT_CMD,
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
]


def find_tesseract():
    for path in TESSERACT_CANDIDATES:
        if path and os.path.exists(path):
            return path
    return shutil.which("tesseract")


try:
    import pytesseract
    TESSERACT_PATH = find_tesseract()
    if TESSERACT_PATH:
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
    OCR_AVAILABLE = bool(TESSERACT_PATH)
except ImportError:
    TESSERACT_PATH = None
    OCR_AVAILABLE = False

# --- OCR backend #2: Gemini (needs internet + a free API key, better on
# handwriting). Used automatically when GEMINI_API_KEY is set in .env --------
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()
# Low-end models that work on a FREE key, tried in order after GEMINI_MODEL.
# Each model has its own rate-limit bucket, so when one runs out of quota the
# next one usually still has some. (gemini-4-argon is NOT here on purpose: its
# free-tier limit is 0, so every call is rejected with a 429 that looks like a
# broken key.)
GEMINI_FALLBACK_MODELS = [
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-flash-lite-latest",
    "gemma-4-26b-a4b-it",
]


def _gemini_model_chain():
    """GEMINI_MODELS (comma-separated) replaces the whole chain if set;
    otherwise it's GEMINI_MODEL followed by the built-in fallbacks."""
    override = os.environ.get("GEMINI_MODELS", "").strip()
    names = override.split(",") if override else [GEMINI_MODEL] + GEMINI_FALLBACK_MODELS
    chain = []
    for name in names:
        name = name.strip()
        if name and name not in chain:
            chain.append(name)
    return chain


GEMINI_MODELS = _gemini_model_chain()
# "low" keeps token use (and so free-quota use) down; the lite models read a
# register fine without extended thinking.
GEMINI_THINKING = os.environ.get("GEMINI_THINKING", "low").strip()
# Seconds to wait on ONE model, and on the whole chain, before moving on.
GEMINI_TIMEOUT = int(os.environ.get("GEMINI_TIMEOUT", "45"))
GEMINI_TOTAL_TIMEOUT = int(os.environ.get("GEMINI_TOTAL_TIMEOUT", "90"))
# Photos are shrunk to this long-edge size before upload: fewer image tokens
# per scan, so the free quota stretches further.
GEMINI_MAX_IMAGE_PX = 1600
# How long a model is skipped after failing, by failure kind (seconds).
GEMINI_COOLDOWN_SECONDS = {"quota": 300, "gone": 86400, "busy": 120}
GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/interactions"
USING_GEMINI = bool(GEMINI_API_KEY)
_gemini_cooldown_until = {}

# Must match the column order build_excel_template.py wrote into each sheet.
# Colors are all drawn from one grey/green family (minimalist palette) --
# categories are told apart by lightness/shade, not by clashing hues.
CATEGORIES = {
    # One green ramp, dark to mint, matching static/style.css and the Excel
    # workbook: categories differ by shade, never by a clashing hue.
    "LateComers": {
        "label": "Late Comers",
        "color": "#123524",
        "extra_fields": [("Time", "text"), ("Reason", "text")],
    },
    "Defaulters": {
        "label": "Defaulters",
        "color": "#2E7D57",
        "extra_fields": [("Reason", "text")],
    },
    "Uniform": {
        "label": "Uniform",
        "color": "#5DBB8C",
        "extra_fields": [("Issue", "text")],
    },
    "NailsHair": {
        "label": "Nails & Hair",
        "color": "#A1E8C0",
        "extra_fields": [("Issue", "text")],
    },
}

CATEGORY_LABELS = [cfg["label"] for cfg in CATEGORIES.values()]


def _gemini_ocr_prompt():
    labels = ", ".join(f'"{cfg["label"]}"' for cfg in CATEGORIES.values())
    return (
        "This is a photo of a handwritten or printed school register page laid "
        "out as a TABLE, with columns including House, Name, and Class (there "
        "may be other columns too -- ignore those). "
        "First, look at the page's title/heading (if there is one) and guess which "
        f"kind of register this is, from exactly these options: {labels}. If you "
        "can't tell, use an empty string. "
        "Then, for every row that is an actual student entry, extract its House, "
        "Name, and Class. "
        "\n\n"
        "HOW TO READ THE TABLE -- this is the part that matters most:\n"
        "1. Work out where each column starts and ends horizontally, then read "
        "the page ROW BY ROW. A value belongs to whichever column it physically "
        "sits under on the page. Do not assign fields by reading order.\n"
        "2. NEVER shift a value left or right into a neighbouring column. If a "
        "cell is empty or unreadable, return an empty string for THAT field and "
        "keep the other fields in their own columns. A blank House must not "
        "cause the Name to be reported as the House.\n"
        "3. House is typically one short word or an abbreviation (often a "
        "colour, a saint's or founder's name, or a single letter). It is often "
        "repeated down a block of consecutive rows, or written once with ditto "
        "marks below -- if a row clearly continues the House above it, repeat "
        "that House rather than leaving it blank.\n"
        "4. Name is a person's name, usually two or more words. If a name wraps "
        "onto a second line within the same row, join it into one name. Never "
        "split one student's name across two rows.\n"
        "5. Class looks like a form or grade such as \"9A\", \"XI B\", \"7-C\", "
        "\"10 Sec B\". If you see something class-shaped sitting in the Name "
        "column, or a person's name sitting in the Class column, you have "
        "misaligned the row -- re-read that row's column positions.\n"
        "6. Do NOT include the header/title row -- that is the row that just "
        "repeats the column labels themselves (words like \"House\", \"Name\", "
        "\"Class\", \"Section\", \"Roll\", \"S.No\", \"Date\" etc used as "
        "headings, not as a real student's data) -- skip it entirely, it is not "
        "a student. Also skip totals, signature lines and any row with no name.\n"
        "7. Return the rows in the same top-to-bottom order they appear on the "
        "page, one JSON object per student row -- never merge two students into "
        "one row or split one student into two.\n\n"
        "DATES:\n"
        "8. If the page shows a date for the whole page (in the title, heading or "
        "a 'Date:' line), return it as page_date. If the table has a Date column, "
        "or a date is written beside a row, return it as that row's date. If dates "
        "are written once and the rows under them belong to it (a new date heading "
        "partway down the page), give each row the date it sits under. Write dates "
        "as YYYY-MM-DD when you can read day, month and year; the register is from "
        "India, so 05/11 means 5 November. If you can only read part of a date, "
        "write it exactly as it appears. Use an empty string when there is no date.\n\n"
        "If a value is missing or unreadable for a real row, use an empty string "
        "for that field, but still include the row if at least the name is readable. "
        "Respond with ONLY a raw JSON object, no markdown code fences, no commentary, "
        "in exactly this shape: "
        '{"category_guess": "...", "page_date": "...", "rows": [{"house": "...", "name": "...", '
        '"class": "...", "date": "..."}, ...]}'
    )


GEMINI_OCR_PROMPT = _gemini_ocr_prompt()

MAX_LOG_ROWS = 1000
MAX_MASTER_ROWS = 2000


def sheet_columns(key):
    cfg = CATEGORIES[key]
    return ["EntryID", "Date", "Name", "Class", "Section", "House"] + [f[0] for f in cfg["extra_fields"]] + [
        "Remarks",
        "LoggedBy",
        "Source",
    ]


# ---------------------------------------------------------------------------
# Workbook helpers
# ---------------------------------------------------------------------------

class WorkbookLocked(Exception):
    """The Excel file couldn't be written -- almost always because someone
    has it open in Excel, which locks it against other writers."""


def open_wb():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(
            f"Database not found at {DB_PATH}. Run build_excel_template.py first."
        )
    return load_workbook(DB_PATH)


def save_wb(wb):
    try:
        wb.save(DB_PATH)
    except PermissionError as exc:
        raise WorkbookLocked() from exc


def parse_entry_date(raw):
    """Date typed into an entry form (YYYY-MM-DD from <input type=date>).
    Blank means today. Returns (date, error_message)."""
    raw = (raw or "").strip()
    if not raw:
        return date.today(), None
    try:
        d = datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError:
        return None, "That date isn't valid."
    if d > date.today():
        return None, "The date can't be in the future."
    return d, None


# Reading the workbook takes ~0.3 s, and every page and every keystroke in a
# name search needs it. So it is read once and reused until the file changes
# on disk (any save -- the app's own or Excel's -- changes its timestamp).
_snapshot_cache = {"key": None, "data": None}
_snapshot_lock = threading.Lock()


def _db_stamp():
    st = os.stat(DB_PATH)
    return (os.path.abspath(DB_PATH), st.st_mtime_ns, st.st_size)


def _read_snapshot():
    """Students and every register's rows, from one read-only pass."""
    wb = load_workbook(DB_PATH, read_only=True)
    try:
        students = []
        for row in wb["MasterList"].iter_rows(min_row=2, max_col=6, values_only=True):
            row = tuple(row) + (None,) * (6 - len(row))
            sid, name = row[0], row[1]
            if not sid or not name:
                continue
            students.append({
                "id": sid, "name": name,
                "class": row[2] or "", "section": row[3] or "", "roll": row[4] or "", "house": row[5] or "",
            })
        registers = {}
        for key in CATEGORIES:
            cols = sheet_columns(key)
            rows = []
            if key in wb.sheetnames:
                for values in wb[key].iter_rows(min_row=2, max_col=len(cols), values_only=True):
                    if not values or not values[0]:
                        continue
                    values = tuple(values) + (None,) * (len(cols) - len(values))
                    # openpyxl round-trips empty-string writes as None; normalize back to "".
                    rows.append({c: ("" if v is None else v) for c, v in zip(cols, values)})
            rows.sort(key=lambda row: (row.get("Date") or date.min, row.get("EntryID") or ""), reverse=True)
            registers[key] = rows
        # One spelling per house and class everywhere in the app
        # ("ARAVALLIS" -> Aravallis, "IB2" / "1B2" -> IB-2, "Grade 6" -> 6).
        canon, canon_cls = {}, {}
        for rows in registers.values():
            for row in rows:
                raw = str(row.get("House") or "")
                if raw not in canon:
                    canon[raw] = match_house(raw, students)
                row["House"] = canon[raw]
                raw_cls = (str(row.get("Class") or ""), str(row.get("Section") or ""))
                if raw_cls not in canon_cls:
                    canon_cls[raw_cls] = match_class(raw_cls[0], raw_cls[1], students)
                row["Class"], row["Section"] = canon_cls[raw_cls]
        return {"students": students, "registers": registers}
    finally:
        wb.close()


def _snapshot():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"Database not found at {DB_PATH}. Run build_excel_template.py first.")
    with _snapshot_lock:
        stamp = _db_stamp()
        if _snapshot_cache["key"] != stamp:
            with wb_lock:
                data = _read_snapshot()
                stamp = _db_stamp()  # the file can't have changed while wb_lock was held by us
            _snapshot_cache["key"], _snapshot_cache["data"] = stamp, data
        return _snapshot_cache["data"]


def get_master_students():
    """Returns list of dicts for every student in MasterList."""
    return [dict(s) for s in _snapshot()["students"]]


def find_student_by_name(name, students=None):
    students = students if students is not None else get_master_students()
    name_lower = (name or "").strip().lower()
    for s in students:
        if s["name"].strip().lower() == name_lower:
            return s
    return None


def find_student_by_id(student_id, students=None):
    students = students if students is not None else get_master_students()
    sid = (student_id or "").strip().lower()
    if not sid:
        return None
    for s in students:
        if str(s["id"]).strip().lower() == sid:
            return s
    return None


def resolve_student(name_or_id, student_id="", students=None):
    """Finds the master-list record for an entry: by an explicit student ID
    first, then by whatever was typed in the name box -- which may itself be
    an ID like STU003 -- then by exact name."""
    students = students if students is not None else get_master_students()
    return (
        find_student_by_id(student_id, students)
        or find_student_by_id(name_or_id, students)
        or find_student_by_name(name_or_id, students)
    )


def class_label(student):
    """"10-B" style label from a master-list record."""
    if not student:
        return ""
    klass = str(student.get("class") or "")
    section = str(student.get("section") or "")
    return f"{klass}-{section}" if klass and section else klass or section


def next_master_row(ws):
    for r in range(2, MAX_MASTER_ROWS + 2):
        if ws.cell(row=r, column=1).value in (None, ""):
            return r
    return ws.max_row + 1


def next_master_id(ws):
    max_num = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and str(v).startswith("STU"):
            try:
                max_num = max(max_num, int(str(v)[3:]))
            except ValueError:
                pass
    return f"STU{max_num + 1:03d}"


def add_student(name, klass, section, roll, house=""):
    with wb_lock:
        wb = open_wb()
        ws = wb["MasterList"]
        r = next_master_row(ws)
        sid = next_master_id(ws)
        ws.cell(row=r, column=1, value=sid)
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=klass)
        ws.cell(row=r, column=4, value=section)
        ws.cell(row=r, column=5, value=roll)
        ws.cell(row=r, column=6, value=house)
        save_wb(wb)
        return sid


def update_student(student_id, name, klass, section, roll, house=""):
    with wb_lock:
        wb = open_wb()
        ws = wb["MasterList"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == student_id:
                ws.cell(row=r, column=2, value=name)
                ws.cell(row=r, column=3, value=klass)
                ws.cell(row=r, column=4, value=section)
                ws.cell(row=r, column=5, value=roll)
                ws.cell(row=r, column=6, value=house)
                save_wb(wb)
                return True
        return False


def next_log_row(ws):
    for r in range(2, MAX_LOG_ROWS + 2):
        if ws.cell(row=r, column=1).value in (None, ""):
            return r
    return ws.max_row + 1


def next_log_entry_id(ws, key):
    max_num = 0
    prefix = key[:2].upper()
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and str(v).startswith(prefix + "-"):
            try:
                max_num = max(max_num, int(str(v).split("-")[1]))
            except (ValueError, IndexError):
                pass
    return f"{prefix}-{max_num + 1:05d}"


def _entry_row_values(key, values_by_header, logged_by, students):
    """Builds the sheet values for one entry. When the student is found in the
    master list (by ID or name) the master record is the source of truth for
    Name/Class/Section/House; otherwise whatever was typed is kept."""
    match = resolve_student(values_by_header.get("Name", ""), values_by_header.get("StudentID", ""), students)
    row_values = {
        "Date": values_by_header.get("Date") or date.today(),
        "Name": match["name"] if match else values_by_header.get("Name", ""),
        "Class": match["class"] if match else values_by_header.get("Class", ""),
        "Section": match["section"] if match else values_by_header.get("Section", ""),
        "House": (match["house"] if match and match["house"] else match_house(values_by_header.get("House", ""), students)),
        "Remarks": values_by_header.get("Remarks", ""),
        "LoggedBy": logged_by or "",
    }
    for field, _ in CATEGORIES[key]["extra_fields"]:
        row_values[field] = values_by_header.get(field, "")
    return row_values


def append_log_entry(key, values_by_header, logged_by, source):
    """values_by_header: dict covering Date, Name (or StudentID), <extra fields>, Remarks."""
    return append_log_entries([(key, values_by_header)], logged_by, source)[0]


def append_log_entries(items, logged_by, source):
    """Writes many entries -- [(register key, values_by_header), ...] -- with
    ONE open and ONE save of the workbook. Saving the workbook takes seconds,
    so a 500-row scan saved row by row took the best part of an hour; this way
    it takes as long as a single entry. All or nothing: if the file can't be
    saved (open in Excel), none of the rows are written. Returns the new IDs."""
    students = get_master_students()
    prepared = [(key, _entry_row_values(key, values, logged_by, students)) for key, values in items]
    with wb_lock:
        wb = open_wb()
        ids = []
        for key, row_values in prepared:
            ws = wb[key]
            r = next_log_row(ws)
            entry_id = next_log_entry_id(ws, key)
            row_values["EntryID"] = entry_id
            row_values["Source"] = source
            for c, header in enumerate(sheet_columns(key), start=1):
                ws.cell(row=r, column=c, value=row_values.get(header, ""))
            ids.append(entry_id)
        if ids:
            save_wb(wb)
        return ids


def _find_log_row(ws, entry_id):
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == entry_id:
            return r
    return None


def get_log_entry(key, entry_id):
    for row in read_log_rows(key, limit=100000):
        if row.get("EntryID") == entry_id:
            return row
    return None


def update_log_entry(key, entry_id, values_by_header, logged_by):
    """Overwrites an existing entry's editable fields (EntryID and Source stay).
    Returns False if the entry no longer exists."""
    cols = sheet_columns(key)
    row_values = _entry_row_values(key, values_by_header, logged_by, get_master_students())

    with wb_lock:
        wb = open_wb()
        ws = wb[key]
        r = _find_log_row(ws, entry_id)
        if r is None:
            return False
        for c, header in enumerate(cols, start=1):
            if header in ("EntryID", "Source"):
                continue
            ws.cell(row=r, column=c).value = row_values.get(header, "")
        save_wb(wb)
        return True


def delete_log_entry(key, entry_id):
    """Blanks the entry's row rather than deleting it: openpyxl's
    delete_rows() doesn't rewrite the per-row VLOOKUP formulas in the rows
    below, which would leave every later row looking up the wrong name. The
    blank row is simply reused by the next new entry. Returns False if the
    entry no longer exists."""
    cols = sheet_columns(key)
    with wb_lock:
        wb = open_wb()
        ws = wb[key]
        r = _find_log_row(ws, entry_id)
        if r is None:
            return False
        for c in range(1, len(cols) + 1):
            # Assign .value directly: ws.cell(..., value=None) is a no-op.
            ws.cell(row=r, column=c).value = None
        # Restore the auto-fill formulas the template puts in empty rows, so
        # picking a name in this row directly in Excel still fills it in.
        name_ref = f"C{r}"
        ws.cell(row=r, column=4).value = f'=IFERROR(VLOOKUP({name_ref},MasterList!$B:$D,2,FALSE)&"","")'
        ws.cell(row=r, column=5).value = f'=IFERROR(VLOOKUP({name_ref},MasterList!$B:$D,3,FALSE)&"","")'
        ws.cell(row=r, column=6).value = f'=IFERROR(VLOOKUP({name_ref},MasterList!$B:$F,5,FALSE)&"","")'
        save_wb(wb)
        return True


def read_log_rows(key, limit=200):
    """A register's saved entries, newest first."""
    return [dict(r) for r in _snapshot()["registers"][key][:limit]]


def count_in_range(rows, start, end):
    n = 0
    for row in rows:
        d = row.get("Date")
        if isinstance(d, datetime):
            d = d.date()
        if isinstance(d, date) and start <= d <= end:
            n += 1
    return n


MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def top_n(counts, n=8):
    items = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:n]
    top_max = max((v for _, v in items), default=0) or 1
    return [{"label": k, "count": v} for k, v in items], top_max


def category_month_breakdown(rows, year):
    """Jan-Dec counts for one category, for the "whole year" view."""
    counts = [0] * 12
    for row in rows:
        d = row.get("Date")
        if isinstance(d, datetime):
            d = d.date()
        if isinstance(d, date) and d.year == year:
            counts[d.month - 1] += 1
    return [{"label": MONTH_NAMES[i], "count": counts[i]} for i in range(12)]


def compute_weekly_repeat_flags(all_rows_by_cat, threshold=3):
    """Flags students who show up `threshold` or more times (across ANY
    category combined) within the same ISO calendar week."""
    tally = {}
    for key, rows in all_rows_by_cat.items():
        for row in rows:
            name = (row.get("Name") or "").strip()
            if not name:
                continue
            d = row.get("Date")
            if isinstance(d, datetime):
                d = d.date()
            if not isinstance(d, date):
                continue
            iso_year, iso_week, _ = d.isocalendar()
            k = (name, iso_year, iso_week)
            entry = tally.setdefault(k, {
                "name": name,
                "class": row.get("Class") or "",
                "iso_year": iso_year,
                "iso_week": iso_week,
                "count": 0,
                "categories": {c: 0 for c in CATEGORIES},
            })
            entry["count"] += 1
            entry["categories"][key] += 1

    flagged = [e for e in tally.values() if e["count"] >= threshold]
    for e in flagged:
        e["week_label"] = "Week of " + date.fromisocalendar(e["iso_year"], e["iso_week"], 1).strftime("%d %b")
    flagged.sort(key=lambda e: (e["iso_year"], e["iso_week"], e["count"]), reverse=True)
    return flagged


def weekly_flag_chart_series(flagged, num_weeks=10):
    """Bar-chart series: how many distinct students were flagged (3+ times
    that week) per week, for the most recent `num_weeks` weeks."""
    today = date.today()
    weeks = []
    seen = set()
    cursor = today
    while len(weeks) < num_weeks:
        iso_year, iso_week, _ = cursor.isocalendar()
        if (iso_year, iso_week) not in seen:
            seen.add((iso_year, iso_week))
            weeks.append((iso_year, iso_week))
        cursor -= timedelta(days=7)
    weeks.reverse()

    counts = {w: 0 for w in weeks}
    for e in flagged:
        wk = (e["iso_year"], e["iso_week"])
        if wk in counts:
            counts[wk] += 1

    return [
        {"label": date.fromisocalendar(iso_year, iso_week, 1).strftime("%d %b"), "count": counts[(iso_year, iso_week)]}
        for iso_year, iso_week in weeks
    ]


WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
MAX_PATTERNS = 10


def _row_date(row):
    d = row.get("Date")
    if isinstance(d, datetime):
        return d.date()
    return d if isinstance(d, date) else None


def _parse_clock_time(value):
    """Late-comer arrival times are free text ("8:45", "8:45 AM", "9am")."""
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, clock_time):
        return value
    text = str(value or "").strip().upper().replace(".", ":")
    for fmt in ("%I:%M %p", "%I:%M%p", "%H:%M", "%I %p", "%I%p"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            continue
    return None


def _row_class_label(row):
    klass = str(row.get("Class") or "").strip()
    section = str(row.get("Section") or "").strip()
    return f"{klass}-{section}" if klass and section else klass or section


def _join_names(names, limit=3):
    shown = list(names)[:limit]
    extra = len(names) - len(shown)
    return ", ".join(shown) + (f" and {extra} more" if extra > 0 else "")


def build_patterns(all_rows_by_cat, weekly_flags, today):
    """Plain-language observations about the data, most urgent first. Each is
    {"text": ..., "tone": "alert" | "up" | "down" | "info"} -- "up" means more
    incidents (worse), "down" means fewer (better)."""
    alerts, trends, observations = [], [], []

    # Students already at 3+ entries this week, and students who keep hitting it.
    this_week = today.isocalendar()[:2]
    flagged_now = sorted({f["name"] for f in weekly_flags if (f["iso_year"], f["iso_week"]) == this_week})
    if flagged_now:
        alerts.append({
            "text": f"{len(flagged_now)} student{'s are' if len(flagged_now) != 1 else ' is'} already at 3+ entries this week: {_join_names(flagged_now)}.",
            "tone": "alert",
        })
    weeks_per_name = Counter(f["name"] for f in weekly_flags)
    chronic = [name for name, n in weeks_per_name.most_common() if n >= 2]
    for name in chronic[:3]:
        alerts.append({
            "text": f"{name} has hit 3+ entries in a single week {weeks_per_name[name]} separate times.",
            "tone": "alert",
        })

    # This month so far vs the same stretch of last month (day 1 to today's day).
    month_start = today.replace(day=1)
    last_month_end = month_start - timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)
    last_month_cutoff = min(last_month_start + timedelta(days=today.day - 1), last_month_end)
    month_rows = []
    for key, rows in all_rows_by_cat.items():
        label = CATEGORIES[key]["label"]
        this_m = last_m = 0
        for row in rows:
            d = _row_date(row)
            if d is None:
                continue
            if month_start <= d <= today:
                this_m += 1
                month_rows.append((key, row))
            elif last_month_start <= d <= last_month_cutoff:
                last_m += 1
        if not this_m and not last_m:
            continue
        if last_m == 0:
            trends.append({"text": f"{label}: {this_m} so far this month, none by this point last month.", "tone": "up"})
        elif this_m == last_m:
            trends.append({"text": f"{label}: {this_m} so far this month, the same as by this point last month.", "tone": "info"})
        else:
            direction = "up" if this_m > last_m else "down"
            pct = round(abs(this_m - last_m) / last_m * 100)
            trends.append({
                "text": f"{label}: {this_m} so far this month, {direction} {pct}% on the {last_m} logged by this point last month.",
                "tone": direction,
            })

    # Who and where, this month.
    month_total = len(month_rows)
    if month_total:
        registers_by_name = {}
        for key, row in month_rows:
            name = (row.get("Name") or "").strip()
            if name:
                registers_by_name.setdefault(name, set()).add(key)
        multi = sorted(n for n, keys in registers_by_name.items() if len(keys) >= 2)
        if multi:
            described = [
                f"{n} ({', '.join(CATEGORIES[k]['label'] for k in CATEGORIES if k in registers_by_name[n])})"
                for n in multi
            ]
            observations.append({
                "text": f"{len(multi)} student{'s appear' if len(multi) != 1 else ' appears'} in more than one register this month: {_join_names(described)}.",
                "tone": "info",
            })

        per_student = Counter((row.get("Name") or "").strip() for _, row in month_rows if (row.get("Name") or "").strip())
        if month_total >= 8 and len(per_student) > 3:
            top3 = sum(n for _, n in per_student.most_common(3))
            share = round(top3 / month_total * 100)
            if share >= 30:
                observations.append({
                    "text": f"Just 3 students account for {share}% of this month's {month_total} entries.",
                    "tone": "info",
                })

        if month_total >= 5:
            for what, counts in (
                ("Class", Counter(_row_class_label(row) for _, row in month_rows if _row_class_label(row))),
                ("House", Counter((row.get("House") or "").strip() for _, row in month_rows if (row.get("House") or "").strip())),
            ):
                if len(counts) >= 2:
                    top, n = counts.most_common(1)[0]
                    observations.append({
                        "text": f"{what} {top} has the most entries this month: {n} of {month_total} ({round(n / month_total * 100)}%).",
                        "tone": "info",
                    })

    # Habits within each register, across all recorded history.
    for key, rows in all_rows_by_cat.items():
        label = CATEGORIES[key]["label"]
        dated = [d for d in (_row_date(r) for r in rows) if d]
        if len(dated) >= 6:
            weekday, n = Counter(d.weekday() for d in dated).most_common(1)[0]
            if n / len(dated) >= 0.3:
                observations.append({
                    "text": f"{label} entries peak on {WEEKDAY_NAMES[weekday]}s: {n} of {len(dated)} ({round(n / len(dated) * 100)}%).",
                    "tone": "info",
                })

        if any(field == "Time" for field, _ in CATEGORIES[key]["extra_fields"]):
            times = [t for t in (_parse_clock_time(r.get("Time")) for r in rows) if t]
            if len(times) >= 4:
                slot, n = Counter((t.hour, t.minute // 15 * 15) for t in times).most_common(1)[0]
                start = datetime.combine(today, clock_time(*slot))
                end = start + timedelta(minutes=15)
                observations.append({
                    "text": f"Most late arrivals are between {start.strftime('%I:%M').lstrip('0')} and {end.strftime('%I:%M %p').lstrip('0')} ({n} of {len(times)} with a time recorded).",
                    "tone": "info",
                })

        detail_field = CATEGORIES[key]["extra_fields"][-1][0]
        if detail_field in ("Reason", "Issue"):
            originals = {}
            counts = Counter()
            for r in rows:
                text = str(r.get(detail_field) or "").strip()
                if text and text.lower() not in ("null", "none", "-"):
                    norm = text.lower()
                    originals.setdefault(norm, text)
                    counts[norm] += 1
            if counts:
                norm, n = counts.most_common(1)[0]
                if n >= 2:
                    observations.append({
                        "text": f"Most common {label} {detail_field.lower()}: “{originals[norm]}” ({n} times).",
                        "tone": "info",
                    })

    return (alerts + trends + observations)[:MAX_PATTERNS]


def _bars(counts_in_order, label_key="label"):
    """[(label, count), ...] -> rows with a 0-100 width for CSS bars."""
    top = max((n for _, n in counts_in_order), default=0) or 1
    return [{"label": label, "count": n, "pct": round(n / top * 100)} for label, n in counts_in_order]


def more_dashboard_stats(all_rows_by_cat, today):
    """Extra numbers for the dashboard: at-a-glance totals, the last 30 days,
    weekday and class patterns, late-arrival times and the top students."""
    rows = [(key, row, _row_date(row)) for key, rs in all_rows_by_cat.items() for row in rs]
    dated = [(k, r, d) for k, r, d in rows if d]
    week_start = today - timedelta(days=today.weekday())
    month_start = today.replace(day=1)
    last_month_end = month_start - timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)
    last_month_cutoff = min(last_month_start + timedelta(days=today.day - 1), last_month_end)

    def name_of(r):
        return (r.get("Name") or "").strip()

    month_rows = [(k, r) for k, r, d in dated if month_start <= d <= today]
    students = get_master_students()
    involved = {name_of(r).lower() for _, r in month_rows if name_of(r)}
    this_m = len(month_rows)
    last_m = sum(1 for _, _, d in dated if last_month_start <= d <= last_month_cutoff)
    glance = {
        "today": sum(1 for _, _, d in dated if d == today),
        "week": sum(1 for _, _, d in dated if week_start <= d <= today),
        "month": this_m,
        "month_change": (None if not last_m else round((this_m - last_m) / last_m * 100)),
        "students_month": len(involved),
        "students_total": len(students),
        "share_month": round(len(involved) / len(students) * 100) if students else 0,
        "all_time": len(rows),
    }

    # Last 30 days, one bar per day.
    days = [today - timedelta(days=i) for i in range(29, -1, -1)]
    per_day = Counter(d for _, _, d in dated if days[0] <= d <= today)
    daily = _bars([(d, per_day.get(d, 0)) for d in days])
    for point in daily:
        point["weekend"] = point["label"].weekday() >= 5
        point["title"] = f"{point['label'].strftime('%a %d %b')}: {point['count']}"
        point["label"] = point["label"].strftime("%d")

    # Weekdays (school days first; Saturday/Sunday only if they have entries).
    wd = Counter(d.weekday() for _, _, d in dated)
    shown = [i for i in range(7) if i < 5 or wd.get(i)]
    by_weekday = _bars([(WEEKDAY_NAMES[i][:3], wd.get(i, 0)) for i in shown])

    # Classes, all time.
    by_class_counts = Counter(_row_class_label(r) for _, r, _ in rows if _row_class_label(r))
    order = lambda c: (0, int(re.match(r"\d+", c).group()), c) if re.match(r"\d+", c) else (1, 0, c)  # noqa: E731
    by_class = _bars(sorted(by_class_counts.items(), key=lambda kv: order(kv[0])))

    # Late arrivals: 15-minute slots.
    times = [t for t in (_parse_clock_time(r.get("Time")) for r in all_rows_by_cat.get("LateComers", [])) if t]
    slots = Counter((t.hour, t.minute // 15 * 15) for t in times)
    late_slots = _bars([(datetime.combine(today, clock_time(h, m)).strftime("%I:%M").lstrip("0"), slots[(h, m)])
                        for h, m in sorted(slots)])
    median_time = _usual_time(times)

    # Top students this year.
    year_start = today.replace(month=1, day=1)
    per_student = {}
    for k, r, d in dated:
        if d < year_start or not name_of(r):
            continue
        e = per_student.setdefault(name_of(r).lower(), {"name": name_of(r), "class": _row_class_label(r),
                                                        "house": r.get("House") or "", "count": 0,
                                                        "by": Counter(), "last": d})
        e["count"] += 1
        e["by"][k] += 1
        e["last"] = max(e["last"], d)
    top_students = sorted(per_student.values(), key=lambda e: (-e["count"], e["name"]))[:10]

    sources = Counter((r.get("Source") or "Manual") for _, r, _ in rows)
    return {
        "glance": glance,
        "daily": daily,
        "by_weekday": by_weekday,
        "by_class": by_class,
        "late_slots": late_slots,
        "late_median": median_time,
        "late_timed": len(times),
        "top_students": top_students,
        "scanned_share": round(sources.get("OCR", 0) / len(rows) * 100) if rows else 0,
    }


def dashboard_data():
    today = date.today()
    month_start = today.replace(day=1)
    year_start = today.replace(month=1, day=1)

    summary = {}
    all_rows_by_cat = {}
    for key, cfg in CATEGORIES.items():
        rows = read_log_rows(key, limit=100000)
        all_rows_by_cat[key] = rows
        summary[key] = {
            "label": cfg["label"],
            "color": cfg["color"],
            "month": count_in_range(rows, month_start, today),
            "year": count_in_range(rows, year_start, today),
            "months": category_month_breakdown(rows, today.year),
        }

    max_month = max((s["month"] for s in summary.values()), default=0) or 1
    max_year_month = max(
        (point["count"] for s in summary.values() for point in s["months"]), default=0
    ) or 1

    house_counts = {}
    for rows in all_rows_by_cat.values():
        for row in rows:
            house = (row.get("House") or "").strip()
            if house:
                house_counts[house] = house_counts.get(house, 0) + 1
    house_breakdown, house_max = top_n(house_counts)

    weekly_flags = compute_weekly_repeat_flags(all_rows_by_cat, threshold=3)
    weekly_flag_chart = weekly_flag_chart_series(weekly_flags, num_weeks=10)
    weekly_flag_chart_max = max((w["count"] for w in weekly_flag_chart), default=0) or 1

    stats_more = more_dashboard_stats(all_rows_by_cat, today)

    stats = {
        **stats_more,
        "max_month": max_month,
        "max_year_month": max_year_month,
        "house_breakdown": house_breakdown,
        "house_max": house_max,
        "weekly_flags": weekly_flags,
        "weekly_flag_chart": weekly_flag_chart,
        "weekly_flag_chart_max": weekly_flag_chart_max,
        "current_year": today.year,
        "patterns": build_patterns(all_rows_by_cat, weekly_flags, today),
    }

    return summary, stats


# ---------------------------------------------------------------------------
# OCR helpers
# ---------------------------------------------------------------------------

def run_ocr(image_path):
    """Returns (text, error, warning, backend). `error` means nothing usable
    came back. `warning` means it succeeded but via a fallback worth telling
    the user about. `backend` is "gemini" or "tesseract" -- whichever one
    actually produced `text`, since Gemini output is structured JSON and
    Tesseract output is plain text; callers need to know which to parse."""
    if USING_GEMINI:
        text, error = run_ocr_gemini(image_path)
        if error is None:
            return text, None, None, "gemini"
        if OCR_AVAILABLE:
            fallback_text, fallback_error = run_ocr_tesseract(image_path)
            if fallback_error is None:
                warning = (
                    f"Gemini was unavailable right now ({error}) — used the offline "
                    "Tesseract engine instead for this scan. Handwriting accuracy will "
                    "be lower than usual; double-check names carefully below."
                )
                return fallback_text, None, warning, "tesseract"
            return "", f"Gemini failed ({error}), and the Tesseract fallback also failed ({fallback_error}).", None, None
        return text, error, None, None
    text, error = run_ocr_tesseract(image_path)
    return text, error, None, "tesseract"


def ocr_ready():
    return USING_GEMINI or OCR_AVAILABLE


def ocr_backend_label():
    """Gemini is the main reader whenever a key is set; Tesseract only steps in
    for a photo Gemini couldn't read (quota, outage, no internet)."""
    if USING_GEMINI:
        return "Gemini (cloud) · Tesseract backup" if OCR_AVAILABLE else "Gemini (cloud)"
    return "Tesseract (offline)"


def run_ocr_tesseract(image_path):
    if not OCR_AVAILABLE:
        return "", "the Tesseract reading program isn't installed on the server PC (run Setup MCTM.cmd)"
    try:
        img = Image.open(image_path)
        img = ImageOps.exif_transpose(img)
        img = img.convert("L")
        img = ImageOps.autocontrast(img)
        text = pytesseract.image_to_string(img)
        return text, None
    except Exception as exc:  # Tesseract binary missing, etc.
        return "", str(exc)


def parse_gemini_response(data):
    """Extracts transcribed text from a Gemini /v1beta/interactions response.
    Response shape: {"steps": [{"type": "model_output", "content": [{"type": "text", "text": "..."}]}]}
    """
    text_parts = []
    for step in data.get("steps", []):
        if step.get("type") != "model_output":
            continue
        for block in step.get("content", []):
            if block.get("type") == "text":
                text_parts.append(block.get("text", ""))
    return "\n".join(text_parts).strip()


def _encode_image_for_gemini(image_path):
    img = Image.open(image_path)
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img.thumbnail((GEMINI_MAX_IMAGE_PX, GEMINI_MAX_IMAGE_PX))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _gemini_error_message(resp):
    try:
        message = resp.json()["error"]["message"]
    except (ValueError, KeyError, TypeError):
        message = resp.text
    return f"HTTP {resp.status_code}: {str(message)[:160]}"


def _gemini_request(model, b64_image, timeout):
    """One OCR call to one model. Returns (text, error, kind); `kind` says
    what to do about a failure:
      "auth"    the key itself was rejected -- no other model will work either
      "network" Gemini can't be reached at all -- same
      "quota"   this model's rate limit is used up -- try the next model
      "gone"    this model was retired or isn't offered to this key
      "busy"    overloaded, timed out, or answered with nothing"""
    payload = {
        "model": model,
        "input": [
            {"type": "text", "text": GEMINI_OCR_PROMPT},
            {"type": "image", "data": b64_image, "mime_type": "image/jpeg"},
        ],
        "generation_config": {"thinking_level": GEMINI_THINKING},
    }
    headers = {"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"}
    try:
        resp = requests.post(GEMINI_ENDPOINT, headers=headers, json=payload, timeout=timeout)
        if resp.status_code == 400 and "thinking" in resp.text.lower():
            # Some models (the Gemma family) reject thinking_level outright.
            del payload["generation_config"]
            resp = requests.post(GEMINI_ENDPOINT, headers=headers, json=payload, timeout=timeout)
    except requests.exceptions.Timeout:
        return "", f"no answer within {int(timeout)}s", "busy"
    except requests.exceptions.RequestException as exc:
        return "", f"could not reach Gemini (check the internet connection): {str(exc)[:160]}", "network"

    if resp.status_code == 200:
        try:
            text = parse_gemini_response(resp.json())
        except ValueError:
            text = ""
        if text:
            return text, None, None
        return "", "answered with no text", "busy"

    error = _gemini_error_message(resp)
    if resp.status_code in (401, 403) or (resp.status_code == 400 and "api key" in resp.text.lower()):
        return "", error, "auth"
    if resp.status_code == 429:
        return "", error, "quota"
    if resp.status_code == 404:
        return "", error, "gone"
    return "", error, "busy"


def run_ocr_gemini(image_path):
    """Tries each model in GEMINI_MODELS until one reads the photo. Models
    that just failed are skipped for a while (see GEMINI_COOLDOWN_SECONDS) so
    a dead one doesn't slow down every scan. Returns (text, error)."""
    try:
        b64_image = _encode_image_for_gemini(image_path)
    except Exception as exc:
        return "", f"could not read the photo: {exc}"

    now = time.monotonic()
    candidates = [m for m in GEMINI_MODELS if _gemini_cooldown_until.get(m, 0) <= now]
    if not candidates:
        candidates = list(GEMINI_MODELS)  # everything is cooling down -- try them all anyway

    deadline = now + GEMINI_TOTAL_TIMEOUT
    failures = []
    for model in candidates:
        remaining = deadline - time.monotonic()
        if remaining < 2:
            failures.append("ran out of time before trying the remaining models")
            break
        text, error, kind = _gemini_request(model, b64_image, min(GEMINI_TIMEOUT, remaining))
        if error is None:
            _gemini_cooldown_until.pop(model, None)
            if failures:
                print(f"[OCR] {model} answered after: {'; '.join(failures)}")
            return text, None
        failures.append(f"{model}: {error}")
        if kind == "auth":
            return "", f"Gemini rejected the API key ({error}). Check GEMINI_API_KEY in .env"
        if kind == "network":
            return "", error
        _gemini_cooldown_until[model] = time.monotonic() + GEMINI_COOLDOWN_SECONDS.get(kind, 120)

    print(f"[OCR] every Gemini model failed: {'; '.join(failures)}")
    return "", "no Gemini model could read it — " + "; ".join(failures)


# Column-heading words that should never be treated as a student's data --
# guards against the OCR engine (or the heuristic column splitter below)
# picking up the register's own header row as if it were a real entry.
HEADER_WORDS = {
    "house", "name", "class", "section", "roll", "roll no", "roll no.",
    "reason", "remarks", "date", "time", "issue", "sl", "sl.no", "s.no",
    "sno", "no", "no.", "signature", "student name", "student",
}


def is_header_like(tokens):
    cleaned = [t.strip(" .:").lower() for t in tokens if t.strip()]
    if not cleaned:
        return False
    hits = sum(1 for t in cleaned if t in HEADER_WORDS)
    return hits >= max(1, (len(cleaned) + 1) // 2)


def match_category_label(text):
    """Fuzzy-matches free text (a page title, or Gemini's category_guess)
    against the known category labels/keys. Returns a CATEGORIES key, or None."""
    text = (text or "").strip()
    if not text:
        return None
    label_to_key = {cfg["label"]: key for key, cfg in CATEGORIES.items()}
    for label, key in label_to_key.items():
        if text.lower() == label.lower() or text.lower() == key.lower():
            return key
    # Case-insensitive fuzzy match: difflib.ratio() is case-sensitive, and a
    # register title is often ALL CAPS while our labels are Title Case.
    lower_to_key = {label.lower(): key for label, key in label_to_key.items()}
    lower_to_key.update({key.lower(): key for key in CATEGORIES})
    matches = difflib.get_close_matches(text.lower(), list(lower_to_key.keys()), n=1, cutoff=0.6)
    if not matches:
        return None
    return lower_to_key[matches[0]]


def parse_gemini_extraction(text):
    """Parses Gemini's {"category_guess", "rows":[...]} JSON output (also
    tolerates a bare rows array, for robustness). Returns (category_guess_text,
    rows) -- rows is None if the text isn't valid JSON in a recognizable shape
    (signals the caller to fall back to plain-text line parsing)."""
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```[a-zA-Z]*\n?", "", cleaned)
        cleaned = re.sub(r"```\s*$", "", cleaned).strip()
    try:
        data = json.loads(cleaned)
    except (ValueError, TypeError):
        return "", None

    if isinstance(data, dict):
        category_guess = str(data.get("category_guess") or "").strip()
        rows = data.get("rows")
    elif isinstance(data, list):
        category_guess = ""
        rows = data
    else:
        return "", None
    if not isinstance(rows, list):
        return category_guess, None

    result = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        result.append({
            "house": str(row.get("house") or "").strip(),
            "name": str(row.get("name") or "").strip(),
            "klass": str(row.get("class") or "").strip(),
            "date": str(row.get("date") or "").strip(),
        })
    return category_guess, result


def gemini_page_date(text):
    """The page_date Gemini read, as a date, or None."""
    cleaned = re.sub(r"^```[a-zA-Z]*\n?|```\s*$", "", (text or "").strip()).strip()
    try:
        data = json.loads(cleaned)
    except (ValueError, TypeError):
        return None
    return parse_register_date(data.get("page_date")) if isinstance(data, dict) else None


# ---------------------------------------------------------------------------
# Dates written on register pages
# ---------------------------------------------------------------------------

_MONTHS = {m.lower(): i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}
_MONTH_RE = r"(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?"
# Tried in order; each yields (day, month, year-or-None).
_DATE_PATTERNS = [
    # 2026-11-18
    (re.compile(r"\b(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\b"), lambda m: (m[3], m[2], m[1])),
    # 18/11/2026, 18-11-26, 18.11.2026 (day first: Indian registers)
    (re.compile(r"\b(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{2,4})\b"), lambda m: (m[1], m[2], m[3])),
    # 18 Nov 2026, 18th November, 18-Nov-26
    (re.compile(r"\b(\d{1,2})(?:st|nd|rd|th)?[\s\-./,]*" + _MONTH_RE + r"(?:[\s\-./,]*(\d{2,4}))?\b", re.I),
     lambda m: (m[1], m[2], m[3])),
    # Nov 18, 2026
    (re.compile(r"\b" + _MONTH_RE + r"\s*(\d{1,2})(?:st|nd|rd|th)?(?:[\s,]+(\d{4}))?\b", re.I),
     lambda m: (m[2], m[1], m[3])),
    # 18/11 (no year). Slash only: "8.45" is a late-arrival time, not 8 April.
    (re.compile(r"\b(\d{1,2})\s*/\s*(\d{1,2})\b(?!\s*/\s*\d)"), lambda m: (m[1], m[2], None)),
]


def _build_date(day, month, year, today):
    try:
        day = int(day)
        month = _MONTHS[month.lower()[:3]] if not str(month).isdigit() else int(month)
    except (ValueError, KeyError):
        return None
    if year is None:
        # No year written: this year, unless that's still ahead of us
        # ("18/12" read in January means last December).
        for y in (today.year, today.year - 1):
            try:
                d = date(y, month, day)
            except ValueError:
                continue
            if d <= today:
                return d
        return None
    year = int(year)
    if year < 100:
        year += 2000
    try:
        d = date(year, month, day)
    except ValueError:
        return None
    if d > today or d.year < today.year - 2:
        return None  # a future date, or one far back, is a misread
    return d


def find_register_dates(text, today=None):
    """Every date found in a piece of OCR text, as (date, start, end) spans,
    in reading order. Day-first; future dates are ignored as misreads."""
    today = today or date.today()
    found, taken = [], []
    for pattern, parts in _DATE_PATTERNS:
        for m in pattern.finditer(text or ""):
            if any(m.start() < e and s < m.end() for s, e in taken):
                continue
            d = _build_date(*parts(m), today)
            if d:
                found.append((d, m.start(), m.end()))
                taken.append((m.start(), m.end()))
    return sorted(found, key=lambda f: f[1])


def parse_register_date(text, today=None):
    """The first readable date in `text` (e.g. '18/11/2026', '18th Nov',
    '2026-11-18'), or None."""
    found = find_register_dates(str(text or ""), today)
    return found[0][0] if found else None


def parse_structured_ocr_json(text):
    """Convenience wrapper around parse_gemini_extraction: just the rows."""
    _, rows = parse_gemini_extraction(text)
    return rows


def _make_candidate(raw_text, house, name, klass, students, row_date=None):
    """One review-table row. If the name matches a student in the master
    list, House/Class/ID come from that record -- what OCR read for those
    columns is only kept for students it couldn't identify. `row_date` is a
    date read for this row (or None: the page's date applies)."""
    student = match_student(name, klass, students)
    return {
        "raw_text": raw_text,
        # The master list's house wins; if it has none, use what the page says.
        "guess_house": (student["house"] if student and student["house"] else match_house(house, students)),
        "guess_name": student["name"] if student else tidy_caps(name),
        "guess_class": class_label(student) if student else klass,
        "student_id": student["id"] if student else "",
        "confident": bool(student),
        "row_date": row_date,
    }


MATCH_CUTOFF = 0.75        # letter-by-letter similarity a scanned name needs
SOUND_ALIKE_CUTOFF = 0.90  # sound-alike similarity needed instead ("Achit" ~ Achyut)
SOUND_ALIKE_CAP = 0.94     # an exact spelling (1.0) always beats a sound-alike
MATCH_MARGIN = 0.05        # the best student must lead the next by this much, or it's a tie


def _alnum(text):
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def tidy_caps(text):
    """'GAUTHAM K' -> 'Gautham K' (registers are often written in block
    capitals); text with any lower case is left as written."""
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    return text.title() if text and text.upper() == text and any(ch.isalpha() for ch in text) else text


# Spelling variants of the same sound in Indian names, applied in order.
_SOUND_RULES = [
    (r"[^a-z ]", ""),
    (r"([a-z])\1+", r"\1"),          # doubled letters: Aarav / Arav, Nikkesh / Nikesh
    (r"au|aw|ow|ou", "o"),           # Gautham / Gowtham
    (r"ee|ie", "i"),                 # Kabeer / Kabir
    (r"oo", "u"),                    # Pooja / Puja
    (r"([bcdgkpst])h", r"\1"),       # th/t dh/d sh/s ch/c kh/k gh/g bh/b ph/p
    (r"w", "v"),                     # Vasudev / Wasudev
    (r"y", "i"),                     # Achyut / Achiut, Riya / Ria
    (r"z", "s"), (r"q", "k"), (r"x", "ks"),
    (r"([a-z])\1+", r"\1"),
]


def sound_key(text):
    """How a name sounds: 'Gowtham' and 'Gautham' both become 'gotam'."""
    s = str(text or "").lower()
    for pattern, repl in _SOUND_RULES:
        s = re.sub(pattern, repl, s)
    return re.sub(r"\s+", " ", s).strip()


def _consonants(key):
    """First letter of each word, then only its consonants: 'acit' -> 'act'."""
    return " ".join(w[0] + re.sub(r"[aeiou]", "", w[1:]) for w in key.split())


def sound_score(a, b):
    """0..SOUND_ALIKE_CAP: how alike two names sound. Averages the sound keys'
    similarity with their consonant outlines', so 'Achit' ~ 'Achyut' scores
    0.94 but 'Arun' ~ 'Tarun' only 0.78. Names too short to judge by sound
    count only when they sound exactly the same."""
    ka, kb = sound_key(a), sound_key(b)
    if not ka or not kb:
        return 0.0
    ca, cb = _consonants(ka), _consonants(kb)
    if min(len(ca.replace(" ", "")), len(cb.replace(" ", ""))) < 3:
        return SOUND_ALIKE_CAP if ka == kb else 0.0
    score = (difflib.SequenceMatcher(None, ka, kb).ratio() + difflib.SequenceMatcher(None, ca, cb).ratio()) / 2
    return min(score, SOUND_ALIKE_CAP)


def name_score(scanned, known):
    """(letter similarity, sound similarity) of a scanned name to a master-list
    name, both ignoring capitals."""
    a = re.sub(r"\s+", " ", str(scanned or "").strip().lower())
    b = re.sub(r"\s+", " ", str(known or "").strip().lower())
    return difflib.SequenceMatcher(None, a, b).ratio(), sound_score(a, b)


def match_class(klass, section, students):
    """(class, section) as the master list writes them, for what a page or a
    scanner wrote: 'IB2' / '1B2' / 'IB-2-A' -> ('IB-2', ''), 'Grade 6' -> ('6', ''),
    '9B' -> ('9', 'B') when 9 has a B section, else ('9', ''). Unknown classes
    are kept as written."""
    known = {}
    for s in students:
        c, sec = str(s.get("class") or "").strip(), str(s.get("section") or "").strip().upper()
        if c:
            known.setdefault(c.upper(), (c, set()))[1].add(sec)
    text = re.sub(r"\s+", " ", f"{klass or ''} {section or ''}").strip().upper()
    if not text:
        return "", ""
    text = re.sub(r"^(GRADE|CLASS|STD|STANDARD)[.]?[ ]*", "", text)  # "Grade 6" -> "6"
    text = re.sub(r"^[1lI][B8](?=[-\s]?\d|$)", "IB", text)  # "1B2" is how a scanner reads "IB2"
    m = re.fullmatch(r"(IB)\s*-?\s*(\d+)\s*-?\s*([A-Z])?", text) or re.fullmatch(r"()(\d+)\s*-?\s*([A-Z])?", text)
    if not m:
        return str(klass or "").strip(), str(section or "").strip()
    cls = f"IB-{m.group(2)}" if m.group(1) else m.group(2)
    sec = m.group(3) or ""
    if cls.upper() not in known:
        return cls, sec
    real, sections = known[cls.upper()]
    return real, (sec if sec in sections else "")


def match_house(text, students):
    """The master-list house a scanned house means ('VINDYAS' -> Vindhyas,
    'Aravali' -> Aravallis), else the text as written (capitals tidied)."""
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if not _alnum(text):
        return ""
    houses = sorted({s["house"] for s in students if s.get("house")})
    best, best_score = None, 0.0
    for h in houses:
        letters, sound = name_score(text, h)
        score = max(letters if letters >= 0.8 else 0, sound if sound >= SOUND_ALIKE_CUTOFF else 0)
        if score > best_score:
            best, best_score = h, score
    return best or tidy_caps(text)


def match_student(name, klass, students):
    """The master-list student a scanned name means, or None when unsure.
    Capitals never matter. In order of strength: the exact spelling; a small
    misread ("Arun Alagapan" -> Arun Alagappan); a different spelling of the
    same sound, at 90%+ ("Achit" -> Achyut, "Gautham" -> Gowtham). Loose
    resemblance never counts. When several students fit equally (two
    "Samyuktha"s), the class written on the page decides; if it can't,
    nobody is picked, so a person checks the row instead of it being filed
    under the wrong student."""
    n = re.sub(r"\s+", " ", str(name or "").strip().lower())
    if len(_alnum(n)) < 2:
        return None

    def by_class(candidates):
        if len(candidates) > 1 and _alnum(klass):
            k = _alnum(klass)
            same = [s for s in candidates if k in (_alnum(s["class"]) + _alnum(s["section"]), _alnum(s["class"]))]
            if same:
                return same
        return candidates

    # 1. The exact spelling (ignoring capitals and spaces) settles it, even
    #    when a classmate's name is one letter away (Siddarth / Siddharth).
    key = _alnum(n)
    exact = [s for s in students
             if key in (_alnum(s["name"]), _alnum(re.sub(r"\s*\([^)]*\)\s*$", "", s["name"])))]
    if exact:
        exact = by_class(exact)
        return exact[0] if len(exact) == 1 else None

    # 2. Misreads and other spellings of the same sound.
    # One wrong letter in a short first name is a different child ("Arun" is
    # not "Tarun"), so short names must match almost exactly.
    cutoff = 0.92 if len(_alnum(n)) <= 6 else MATCH_CUTOFF
    scored = []
    for s in students:
        full = s["name"].strip().lower()
        base = re.sub(r"\s*\([^)]*\)\s*$", "", full)  # "samyuktha (9)" -> "samyuktha"
        r = 0.0
        for known in {full, base}:
            letters, sound = name_score(n, known)
            if letters >= cutoff:
                r = max(r, letters)
            if sound >= SOUND_ALIKE_CUTOFF:
                r = max(r, sound)
        if r:
            scored.append((r, s))
    if not scored:
        return None
    best = max(r for r, _ in scored)
    # A house name in the name column ("Kailash") is a misread column, not a
    # student who merely looks like it (Akileash). Exact names never get here.
    for h in {s["house"] for s in students if s.get("house")}:
        letters, sound = name_score(n, h)
        if letters >= 0.8 or sound >= SOUND_ALIKE_CUTOFF:
            return None
    top = by_class([s for r, s in scored if r >= best - MATCH_MARGIN])
    return top[0] if len(top) == 1 else None


def build_candidates_from_structured(rows, students):
    candidates = []
    for row in rows:
        name = row["name"]
        if not name or name.strip().lower() in HEADER_WORDS:
            continue  # no name to work with, or the model slipped in a header row anyway
        raw_text = " / ".join(v for v in (row["house"], name, row["klass"], row.get("date", "")) if v)
        candidates.append(_make_candidate(raw_text, row["house"], name, row["klass"], students,
                                          parse_register_date(row.get("date"))))
    return candidates


def split_columns(line):
    """Best-effort column split for plain OCR text (Tesseract has no concept
    of table structure) -- splits on runs of 2+ spaces or tabs, which is how
    column gaps often survive into OCR'd text."""
    parts = re.split(r"\s{2,}|\t+", line.strip())
    return [p.strip() for p in parts if p.strip()]


def parse_ocr_lines(raw_text, students):
    """Heuristic House/Name/Class extraction from plain OCR text (the
    Tesseract path -- Gemini's structured JSON is handled separately, far
    more reliably, by build_candidates_from_structured)."""
    lines = [ln.strip() for ln in raw_text.splitlines()]
    lines = [re.sub(r"^[\W_]+", "", ln) for ln in lines]  # strip leading bullets/numbers/punct
    lines = [ln for ln in lines if len(ln) >= 2]

    candidates = []
    heading_date = None  # a date written on its own line applies to the rows below it
    for ln in lines:
        dates = find_register_dates(ln)
        body = ln
        for _, start, end in reversed(dates):
            body = body[:start] + "  " + body[end:]
        body = re.sub(r"(?i)\b(date|dated|on)\b\s*[:\-]?", " ", body).strip(" -:,/\t")
        if dates and len(body) < 2:
            heading_date = dates[0][0]  # a "Date: 18/11/2026" line, not a student
            continue
        if dates and match_category_label(body):
            heading_date = dates[0][0]  # "UNIFORM REGISTER 18/11/2026"
            continue

        parts = split_columns(body)
        if is_header_like(parts if len(parts) > 1 else body.split()):
            continue  # this line is the column-title row, not a student

        found = None if len(parts) >= 3 else split_around_known_name(body, students)
        if found:
            house, name, klass = found  # Tesseract squeezed the column gaps to single spaces
        elif len(parts) >= 3:
            house, name, klass = parts[0], parts[1], parts[2]
        elif len(parts) == 2:
            house, name, klass = "", parts[0], parts[1]
        else:
            house, name, klass = "", body, ""

        row_date = dates[0][0] if dates else heading_date
        candidates.append(_make_candidate(ln, house, name, klass, students, row_date))
    return candidates


def page_date_from_text(raw_text):
    """The date written near the top of a plain-text (Tesseract) page: the
    first date in its first eight lines, or None."""
    head = "\n".join([ln for ln in (raw_text or "").splitlines() if ln.strip()][:8])
    return parse_register_date(head)


def split_around_known_name(line, students):
    """For a line whose columns ran together ("Kailash ArunAlagappan IB-1A"):
    finds the run of words that is a master-list student's name, and returns
    (house, name, class) as the words before it, the name, and the words
    after it. Exact matches (ignoring spaces, so "ArunAlagappan" counts) win;
    otherwise a close match via match_student. None if no name is found."""
    tokens = [t for t in line.split() if _alnum(t)]  # drop OCR specks like "|" or "�"
    if not tokens:
        return None
    exact = {}  # spelling without spaces -> every student it could mean
    for s in students:
        full = _alnum(s["name"])
        base = _alnum(re.sub(r"\s*\([^)]*\)\s*$", "", s["name"]))
        exact.setdefault(full, []).append(s)
        if base != full:
            exact.setdefault(base, []).append(s)
    # A house ("Kailash") or class ("IB-1A") is never a name, however much it
    # looks like one (Kailash ~ Akileash).
    houses = {_alnum(s["house"]) for s in students if s.get("house")}

    def not_a_name(span):
        key = _alnum(span)
        if any(ch.isdigit() for ch in key) or key in houses:
            return True
        return any(difflib.SequenceMatcher(None, key, h).ratio() >= 0.8 for h in houses)

    best = None  # (score, length, start, end, name)
    for start in range(len(tokens)):
        for end in range(start + 1, min(start + 4, len(tokens)) + 1):
            span = " ".join(tokens[start:end])
            key = _alnum(span)
            if len(key) < 3:
                continue
            if key in exact:
                matches = {s["id"]: s for s in exact[key]}
                # One student: use their real name ("Samyuktha 9" -> "Samyuktha (9)").
                # Several (a bare "Samyuktha"): keep the text; the class decides later.
                name = next(iter(matches.values()))["name"] if len(matches) == 1 else span
                score = 2
            elif not any(not_a_name(t) for t in tokens[start:end]) and match_student(span, "", students):
                name, score = span, 1
            else:
                continue
            cand = (score, end - start, start, end, name)
            if best is None or cand[:2] > best[:2]:
                best = cand
    if best is None:
        return None
    _, _, start, end, name = best
    return " ".join(tokens[:start]), name, " ".join(tokens[end:])


def _title_text(line):
    """A page-title line without its date or filler words: 'LATE COMERS.
    Date: 17/11/2026' -> 'LATE COMERS'."""
    for _, start, end in reversed(find_register_dates(line)):
        line = line[:start] + " " + line[end:]
    line = re.sub(r"(?i)\b(date|dated|register|book|list|record|records)\b", " ", line)
    return re.sub(r"[^\w&]+", " ", line).strip()


def guess_category_from_text_lines(raw_text):
    """Best-effort register-type detection for the Tesseract (plain text)
    path: checks the first few lines for something that looks like a title
    matching one of the known categories (e.g. "UNIFORM REGISTER")."""
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()][:5]
    for ln in lines:
        key = match_category_label(ln) or match_category_label(_title_text(ln))
        if key:
            return key
    return None


def extract_candidates(raw_text, students, backend):
    """Turns raw OCR output into (candidates, category_guess_key), using
    structured JSON parsing for Gemini and heuristic line parsing for
    Tesseract. category_guess_key is a CATEGORIES key, or None if the page's
    register type couldn't be confidently identified from its title."""
    if not raw_text:
        return [], None
    if backend == "gemini":
        category_guess_text, structured = parse_gemini_extraction(raw_text)
        if structured is not None:
            candidates = build_candidates_from_structured(structured, students)
            return candidates, match_category_label(category_guess_text)
        # Gemini didn't follow the JSON format for some reason -- fall back
        # to treating its output as plain text rather than losing the scan.
    return parse_ocr_lines(raw_text, students), guess_category_from_text_lines(raw_text)


def extract_page_date(raw_text, backend):
    """The date written for the whole page, or None if none could be read."""
    if not raw_text:
        return None
    if backend == "gemini" and parse_structured_ocr_json(raw_text) is not None:
        return gemini_page_date(raw_text)
    # Tesseract text, or Gemini text that wasn't JSON: look near the top.
    return page_date_from_text(raw_text)


# ---------------------------------------------------------------------------
# Scanning several photos at once
# ---------------------------------------------------------------------------

MAX_SCAN_PHOTOS = 10
SCAN_WORKERS = 3  # photos read at the same time (keeps a free Gemini key inside its rate limit)
ALLOWED_PHOTO_EXT = {".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif", ".bmp", ".tif", ".tiff", ".gif"}


def safe_upload_name(original, index):
    """A filename that can't escape the scans folder, keeps the extension,
    and won't collide with another photo uploaded in the same second."""
    base, ext = os.path.splitext(os.path.basename(original or ""))
    ext = ext.lower() if ext.lower() in ALLOWED_PHOTO_EXT else ".jpg"
    base = re.sub(r"[^A-Za-z0-9_-]+", "_", base).strip("_")[:40] or "photo"
    return f"scan_{datetime.now():%Y%m%d_%H%M%S}_{index + 1:02d}_{base}{ext}"


def read_one_scan(path, students, hint_key, today):
    """OCR + parse one saved photo into a review group."""
    raw_text, error, warning, backend = run_ocr(path)
    candidates, category_guess = extract_candidates(raw_text, students, backend)
    page_date = extract_page_date(raw_text, backend)
    if page_date is None:
        row_dates = Counter(c["row_date"] for c in candidates if c["row_date"])
        page_date = row_dates.most_common(1)[0][0] if row_dates else None
    return {
        "candidates": candidates,
        "detected_key": category_guess or hint_key,
        "category_was_detected": bool(category_guess),
        "page_date": page_date or today,
        "date_was_detected": page_date is not None,
        "raw_text": raw_text,
        "error": error,
        "warning": warning,
    }


def read_scans(saved, hint_key, on_progress=None):
    """Reads every saved photo (a few at a time) and returns one group per
    photo, in upload order. `on_progress()` is called after each photo."""
    from concurrent.futures import ThreadPoolExecutor

    students = get_master_students()
    today = date.today()

    def one(item):
        try:
            result = read_one_scan(item[1], students, hint_key, today)
        except Exception as exc:  # one bad photo must not sink a 100-photo batch
            result = {"candidates": [], "detected_key": hint_key, "category_was_detected": False,
                      "page_date": today, "date_was_detected": False, "raw_text": "",
                      "error": f"couldn't read this photo ({exc})", "warning": None}
        if on_progress:
            on_progress()
        return result

    with ThreadPoolExecutor(max_workers=min(SCAN_WORKERS, len(saved)) or 1) as pool:
        results = list(pool.map(one, saved))
    groups = []
    for (original, path), result in zip(saved, results):
        result["filename"] = os.path.basename(path)
        result["original_name"] = original
        groups.append(result)
    return groups


# ---------------------------------------------------------------------------
# Scan folder: drop up to 100 photos in one folder, scan them all at once
# ---------------------------------------------------------------------------

MAX_FOLDER_PHOTOS = 100
SETTINGS_PATH = os.path.join(DATA_DIR, "settings.json")
DEFAULT_SCAN_FOLDER = os.path.join(DATA_DIR, "scan-folder")
SCANNED_SUBFOLDER = "Scanned"  # photos move here once their rows are saved
SCAN_JOB_TTL = 12 * 3600       # seconds a finished batch is kept for review

SCAN_JOBS = {}
scan_jobs_lock = threading.Lock()
SCAN_JOBS_DIR = os.path.join(DATA_DIR, "scan-jobs")  # finished batches, so a restart doesn't lose a review


def _job_file(job_id):
    return os.path.join(SCAN_JOBS_DIR, f"{job_id}.json") if re.fullmatch(r"[A-Za-z0-9_-]{6,40}", job_id or "") else None


def _to_json(value):
    if isinstance(value, (date, datetime)):
        return {"__date__": value.isoformat()[:10]}
    if isinstance(value, dict):
        return {k: _to_json(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_to_json(v) for v in value]
    return value


def _from_json(value):
    if isinstance(value, dict):
        if set(value) == {"__date__"}:
            return date.fromisoformat(value["__date__"])
        return {k: _from_json(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_from_json(v) for v in value]
    return value


def _save_job(job):
    path = _job_file(job["id"])
    if not path:
        return
    os.makedirs(SCAN_JOBS_DIR, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(_to_json(job), f)
    os.replace(tmp, path)


def _load_job(job_id):
    path = _job_file(job_id)
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            return _from_json(json.load(f))
    except (OSError, ValueError):
        return None


def _delete_job_file(job_id):
    path = _job_file(job_id)
    if path:
        try:
            os.remove(path)
        except OSError:
            pass


def load_settings():
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(changes):
    data = load_settings()
    data.update(changes)
    os.makedirs(os.path.dirname(SETTINGS_PATH), exist_ok=True)
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def scan_folder_path():
    """The folder to scan: the one chosen on the Scan page, else data/scan-folder
    (created on first use)."""
    folder = load_settings().get("scan_folder") or DEFAULT_SCAN_FOLDER
    if os.path.abspath(folder) == os.path.abspath(DEFAULT_SCAN_FOLDER):
        os.makedirs(folder, exist_ok=True)
    return folder


def list_folder_photos(folder):
    """Photos directly in `folder` (not in subfolders such as Scanned), oldest first."""
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    photos = [os.path.join(folder, n) for n in names
              if os.path.splitext(n)[1].lower() in ALLOWED_PHOTO_EXT and os.path.isfile(os.path.join(folder, n))]
    return sorted(photos, key=lambda p: (os.path.getmtime(p), os.path.basename(p).lower()))


def _forget_old_jobs():
    now = time.time()
    with scan_jobs_lock:
        for job_id in [j for j, job in SCAN_JOBS.items()
                       if job["status"] != "running" and now - job["created"] > SCAN_JOB_TTL]:
            del SCAN_JOBS[job_id]
    try:
        for name in os.listdir(SCAN_JOBS_DIR):
            path = os.path.join(SCAN_JOBS_DIR, name)
            if now - os.path.getmtime(path) > SCAN_JOB_TTL:
                os.remove(path)
    except OSError:
        pass


def running_folder_job():
    with scan_jobs_lock:
        return next((job for job in SCAN_JOBS.values() if job["status"] == "running"), None)


def start_folder_job(folder, hint_key):
    """Copies up to MAX_FOLDER_PHOTOS photos from `folder` into the scans
    store and reads them in a background thread. Returns the job."""
    _forget_old_jobs()
    sources = list_folder_photos(folder)[:MAX_FOLDER_PHOTOS]
    job = {
        "id": base64.urlsafe_b64encode(os.urandom(9)).decode(),
        "status": "running", "total": len(sources), "done": 0, "groups": None, "error": None,
        "hint": hint_key, "folder": folder, "sources": sources, "created": time.time(),
    }
    with scan_jobs_lock:
        SCAN_JOBS[job["id"]] = job

    def work():
        try:
            saved = []
            for i, src in enumerate(sources):
                dest = os.path.join(UPLOAD_DIR, safe_upload_name(os.path.basename(src), i))
                shutil.copy2(src, dest)
                saved.append((os.path.basename(src), dest))

            def tick():
                with scan_jobs_lock:
                    job["done"] += 1

            groups = read_scans(saved, hint_key, on_progress=tick)
            with scan_jobs_lock:
                job["groups"] = groups
                job["status"] = "done"
            _save_job(job)
        except Exception as exc:
            with scan_jobs_lock:
                job["error"] = str(exc)
                job["status"] = "error"

    threading.Thread(target=work, daemon=True, name=f"scan-folder-{job['id']}").start()
    return job


def get_job(job_id):
    """A scan batch: from memory, or from disk after a server restart."""
    with scan_jobs_lock:
        job = SCAN_JOBS.get(job_id)
    if job is None:
        job = _load_job(job_id)
        if job is not None:
            with scan_jobs_lock:
                SCAN_JOBS[job_id] = job
    return job


def move_to_scanned(paths, folder):
    """Moves photos whose rows were saved into <folder>/Scanned, so the next
    folder scan doesn't read them again. Returns how many were moved."""
    dest_dir = os.path.join(folder, SCANNED_SUBFOLDER)
    moved = 0
    for src in paths:
        if not os.path.isfile(src):
            continue
        os.makedirs(dest_dir, exist_ok=True)
        base, ext = os.path.splitext(os.path.basename(src))
        dest = os.path.join(dest_dir, base + ext)
        n = 2
        while os.path.exists(dest):
            dest = os.path.join(dest_dir, f"{base} ({n}){ext}")
            n += 1
        try:
            shutil.move(src, dest)
            moved += 1
        except OSError:
            pass  # open in another program: it simply gets offered again next time
    return moved


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def dashboard():
    summary, stats = dashboard_data()
    return render_template(
        "dashboard.html",
        categories=CATEGORIES,
        summary=summary,
        stats=stats,
        today=date.today(),
    )


@app.errorhandler(WorkbookLocked)
def workbook_locked(_exc):
    flash(
        "Couldn't save: the Excel file is open somewhere (usually in Excel on the server PC). "
        "Close it there and try again.",
        "error",
    )
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/api/students")
def api_students():
    q = (request.args.get("q") or "").strip().lower()
    students = get_master_students()
    if not q:
        results = students[:8]
    else:
        starts = [s for s in students if s["name"].lower().startswith(q) or str(s["id"]).lower().startswith(q)]
        contains = [s for s in students if q in s["name"].lower() and s not in starts]
        results = (starts + contains)[:8]
    return jsonify(results)


@app.route("/api/students/lookup")
def api_student_lookup():
    """Exact match by student ID or full name -- used to fill in the rest of
    a form when someone types a name/ID without picking from the list."""
    return jsonify(resolve_student(request.args.get("q", "")))


MAX_ENTRIES_SHOWN = 500


def _entry_form_values(cfg):
    """Reads an entry form. Returns (values, logged_by, error_message)."""
    entry_date, error = parse_entry_date(request.form.get("date"))
    if error:
        return None, None, error
    name = request.form.get("name", "").strip()
    student_id = request.form.get("student_id", "").strip()
    if not name and not student_id:
        return None, None, "Please pick a student (type their name or student ID)."
    if not name and not find_student_by_id(student_id):
        return None, None, f"There's no student with ID {student_id} in the master list."
    values = {
        "Date": entry_date,
        "Name": name,
        "StudentID": student_id,
        "Class": request.form.get("klass", "").strip(),
        "Section": request.form.get("section", "").strip(),
        "House": request.form.get("house", "").strip(),
        "Remarks": request.form.get("remarks", "").strip(),
    }
    for field, _ in cfg["extra_fields"]:
        values[field] = request.form.get(field, "").strip()
    return values, request.form.get("logged_by", "").strip(), None


def _form_from_entry(row, cfg):
    """Pre-fills the edit form from a saved entry."""
    d = _row_date(row)
    student = find_student_by_name(row.get("Name", ""))
    form = {
        "date": d.isoformat() if d else "",
        "name": row.get("Name", ""),
        "student_id": student["id"] if student else "",
        "klass": row.get("Class", ""),
        "section": row.get("Section", ""),
        "house": row.get("House", ""),
        "logged_by": row.get("LoggedBy", ""),
        "remarks": row.get("Remarks", ""),
    }
    for field, _ in cfg["extra_fields"]:
        form[field] = row.get(field, "")
    return form


def _render_entry_page(key, form, status=200):
    cfg = CATEGORIES[key]
    rows = read_log_rows(key, limit=MAX_ENTRIES_SHOWN)
    return render_template(
        "entry.html", key=key, cfg=cfg, categories=CATEGORIES, rows=rows,
        form=form, today=date.today(), max_shown=MAX_ENTRIES_SHOWN,
    ), status


@app.route("/entry/<key>", methods=["GET", "POST"])
def entry(key):
    if key not in CATEGORIES:
        return "Unknown category", 404
    cfg = CATEGORIES[key]

    if request.method == "POST":
        values, logged_by, error = _entry_form_values(cfg)
        if error:
            flash(error, "error")
            return _render_entry_page(key, request.form, status=400)
        entry_id = append_log_entry(key, values, logged_by, "Manual")
        saved = get_log_entry(key, entry_id)
        flash(f"Saved {entry_id} for {saved['Name'] if saved else values['Name']} ({values['Date'].strftime('%d %b %Y')}).", "success")
        return redirect(url_for("entry", key=key))

    return _render_entry_page(key, {"date": date.today().isoformat()})


@app.route("/entry/<key>/<entry_id>/edit", methods=["GET", "POST"])
def edit_entry(key, entry_id):
    if key not in CATEGORIES:
        return "Unknown category", 404
    cfg = CATEGORIES[key]
    row = get_log_entry(key, entry_id)
    if row is None:
        flash(f"Entry {entry_id} doesn't exist (it may have been deleted).", "error")
        return redirect(url_for("entry", key=key))

    if request.method == "POST":
        values, logged_by, error = _entry_form_values(cfg)
        if error:
            flash(error, "error")
            return render_template(
                "entry_edit.html", key=key, cfg=cfg, categories=CATEGORIES,
                entry_id=entry_id, row=row, form=request.form, today=date.today(),
            ), 400
        if update_log_entry(key, entry_id, values, logged_by):
            flash(f"Updated {entry_id}.", "success")
        else:
            flash(f"Entry {entry_id} doesn't exist anymore.", "error")
        return redirect(url_for("entry", key=key))

    return render_template(
        "entry_edit.html", key=key, cfg=cfg, categories=CATEGORIES,
        entry_id=entry_id, row=row, form=_form_from_entry(row, cfg), today=date.today(),
    )


@app.route("/entry/<key>/<entry_id>/delete", methods=["POST"])
def delete_entry(key, entry_id):
    if key not in CATEGORIES:
        return "Unknown category", 404
    if delete_log_entry(key, entry_id):
        flash(f"Deleted {entry_id}.", "success")
    else:
        flash(f"Entry {entry_id} was already gone.", "error")
    return redirect(url_for("entry", key=key))


def _request_is_from_server_pc():
    """True when the browser is running on the PC that hosts the app, either
    via localhost or via the server's own LAN address."""
    remote = request.remote_addr or ""
    host = (request.host or "").rsplit(":", 1)[0].strip("[]")
    return remote in ("127.0.0.1", "::1") or remote == host


def _open_file_on_server(path):
    os.startfile(path)


@app.route("/open-excel")
def open_excel():
    """On the server PC, opens the live workbook in Excel. From any other PC
    (or with ?download=1) it downloads a copy instead -- changes made to that
    copy are NOT saved back to the app."""
    if request.args.get("download") != "1" and _request_is_from_server_pc() and hasattr(os, "startfile"):
        try:
            _open_file_on_server(DB_PATH)
        except OSError as exc:
            flash(f"Couldn't open Excel: {exc}", "error")
        else:
            flash(
                "Opened the Excel file on this PC. Close Excel before logging or editing "
                "entries here; the app can't save while the file is open.",
                "success",
            )
        return redirect(request.referrer or url_for("dashboard"))
    return send_file(DB_PATH, as_attachment=True, download_name="SchoolDisciplineSystem.xlsx")


@app.route("/scan", methods=["GET", "POST"])
def scan():
    # `hint`: which category the user was looking at when they clicked Scan
    # (e.g. from a dashboard card). Used only as a fallback default if the
    # photo's own title can't be confidently detected -- detection always
    # wins when it succeeds, since it reflects the actual photo.
    if request.method == "POST":
        hint_key = request.form.get("hint")
        hint_key = hint_key if hint_key in CATEGORIES else None
        files = [f for f in request.files.getlist("photo") if f and f.filename]
        if not files:
            flash("Please choose at least one photo of a register page.", "error")
            return redirect(url_for("scan", hint=hint_key or ""))
        if len(files) > MAX_SCAN_PHOTOS:
            flash(f"You chose {len(files)} photos; only the first {MAX_SCAN_PHOTOS} were read. "
                  "Scan the rest in another batch.", "error")
            files = files[:MAX_SCAN_PHOTOS]

        saved = []
        for i, f in enumerate(files):
            path = os.path.join(UPLOAD_DIR, safe_upload_name(f.filename, i))
            f.save(path)
            saved.append((f.filename, path))

        groups = read_scans(saved, hint_key)
        return render_template(
            "scan_review.html",
            categories=CATEGORIES,
            groups=groups,
            hint_key=hint_key,
            total_rows=sum(len(g["candidates"]) for g in groups),
            ocr_available=ocr_ready(),
            ocr_backend=ocr_backend_label(),
            today=date.today(),
        )

    hint_key = request.args.get("hint")
    hint_key = hint_key if hint_key in CATEGORIES else ""
    folder = scan_folder_path()
    return render_template(
        "scan.html", categories=CATEGORIES, hint_key=hint_key, max_photos=MAX_SCAN_PHOTOS,
        ocr_available=ocr_ready(), ocr_backend=ocr_backend_label(),
        folder=folder, folder_exists=os.path.isdir(folder), folder_count=len(list_folder_photos(folder)),
        folder_is_default=os.path.abspath(folder) == os.path.abspath(DEFAULT_SCAN_FOLDER),
        max_folder=MAX_FOLDER_PHOTOS, on_server_pc=_request_is_from_server_pc(), using_gemini=USING_GEMINI,
        running_job=running_folder_job(),
    )


@app.route("/scan/photo/<filename>")
def scan_photo(filename):
    """A saved register photo, shown beside its rows on the review page."""
    if not re.fullmatch(r"scan_[A-Za-z0-9_-]+\.[A-Za-z0-9]{2,5}", filename):
        return "Not found", 404
    from flask import send_from_directory
    return send_from_directory(UPLOAD_DIR, filename)


@app.route("/scan/folder/start", methods=["POST"])
def scan_folder_start():
    hint_key = request.form.get("hint")
    hint_key = hint_key if hint_key in CATEGORIES else None
    running = running_folder_job()
    if running:
        return redirect(url_for("scan_job", job_id=running["id"]))
    folder = scan_folder_path()
    photos = list_folder_photos(folder)
    if not photos:
        flash(f"There are no photos in {folder}. Put the register photos there, then scan.", "error")
        return redirect(url_for("scan", hint=hint_key or ""))
    if len(photos) > MAX_FOLDER_PHOTOS:
        flash(f"The folder has {len(photos)} photos; the oldest {MAX_FOLDER_PHOTOS} are being read now. "
              "Scan the folder again afterwards for the rest.", "error")
    job = start_folder_job(folder, hint_key)
    return redirect(url_for("scan_job", job_id=job["id"]))


@app.route("/scan/job/<job_id>")
def scan_job(job_id):
    job = get_job(job_id)
    if job is None:
        flash("That batch is no longer available (the server was restarted, or it was already saved). "
              "Scan the folder again.", "error")
        return redirect(url_for("scan"))
    if job["status"] == "running":
        return render_template("scan_job.html", categories=CATEGORIES, job=job)
    if job["status"] == "error":
        flash(f"The folder scan stopped: {job['error']}", "error")
        return redirect(url_for("scan"))
    groups = job["groups"]
    return render_template(
        "scan_review.html", categories=CATEGORIES, groups=groups, hint_key=job["hint"], job=job,
        total_rows=sum(len(g["candidates"]) for g in groups),
        ocr_available=ocr_ready(), ocr_backend=ocr_backend_label(), today=date.today(),
    )


@app.route("/scan/job/<job_id>/status")
def scan_job_status(job_id):
    job = get_job(job_id)
    if job is None:
        return jsonify({"status": "gone"}), 404
    return jsonify({"status": job["status"], "done": job["done"], "total": job["total"]})


@app.route("/scan/folder/settings", methods=["POST"])
def scan_folder_settings():
    hint = request.form.get("hint") or ""
    raw = (request.form.get("folder") or "").strip().strip('"')
    if request.form.get("reset") or not raw:
        save_settings({"scan_folder": ""})
        flash(f"Scanning the default folder again: {scan_folder_path()}", "success")
        return redirect(url_for("scan", hint=hint))
    folder = os.path.abspath(os.path.expanduser(raw))
    if not os.path.isdir(folder):
        flash(f"There's no folder at {folder} on the server PC. Create it first, or check the path.", "error")
        return redirect(url_for("scan", hint=hint))
    save_settings({"scan_folder": folder})
    flash(f"Photos will be scanned from {folder}.", "success")
    return redirect(url_for("scan", hint=hint))


@app.route("/scan/gemini-key", methods=["POST"])
def scan_gemini_key():
    """Opens SETUP_GEMINI_KEY.bat on the server PC. The key is typed into that
    window, never into the web page, so it never crosses the network."""
    hint = request.form.get("hint") or ""
    script = os.path.join(BASE_DIR, "SETUP_GEMINI_KEY.bat")
    if _request_is_from_server_pc() and hasattr(os, "startfile") and os.path.exists(script):
        try:
            _open_file_on_server(script)
            flash("Paste the key into the window that opened. The tracker restarts by itself when it's saved.", "success")
        except OSError as exc:
            flash(f"Couldn't open the key setup: {exc}", "error")
    else:
        flash("Add the Gemini key on the server PC: run SETUP_GEMINI_KEY.bat there.", "error")
    return redirect(url_for("scan", hint=hint))


@app.route("/scan/folder/open", methods=["POST"])
def scan_folder_open():
    """Opens the scan folder in File Explorer -- only on the server PC itself."""
    hint = request.form.get("hint") or ""
    folder = scan_folder_path()
    if _request_is_from_server_pc() and hasattr(os, "startfile"):
        try:
            _open_file_on_server(folder)
        except OSError as exc:
            flash(f"Couldn't open the folder: {exc}", "error")
    else:
        flash(f"The scan folder is on the server PC: {folder}", "error")
    return redirect(url_for("scan", hint=hint))


@app.route("/scan/commit", methods=["POST"])
def scan_commit():
    """Saves the ticked rows of one or more scanned photos. Each photo
    ("group" g) has its own register (group_category_<g>) and page date
    (group_date_<g>); a row may carry its own date (row_date), which wins.
    A form without groups (the old single-photo shape) uses "category" and
    "date" for everything. Every row is checked before anything is written:
    one bad register or date means nothing is saved."""
    form = request.form
    names = form.getlist("row_name")
    student_ids = form.getlist("row_student_id")
    houses = form.getlist("row_house")
    classes = form.getlist("row_class")
    remarks_list = form.getlist("row_remarks")
    row_dates = form.getlist("row_date")
    row_groups = form.getlist("row_group")
    logged_by = form.get("logged_by", "").strip()
    job = get_job(form.get("job_id", ""))  # set when the rows came from a folder scan
    try:
        included = {int(i) for i in form.getlist("row_include")}  # indices of the ticked rows
    except ValueError:
        included = set()

    def at(values, i):
        return values[i].strip() if i < len(values) and values[i] else ""

    planned, problems = [], []
    for i, name in enumerate(names):
        if i not in included:
            continue
        name = name.strip()
        student_id = at(student_ids, i)
        if not name and not student_id:
            continue
        g = at(row_groups, i) or "0"
        key = form.get(f"group_category_{g}") or form.get("category")
        if key not in CATEGORIES:
            problems.append(f"photo {int(g) + 1 if g.isdigit() else g}: choose which register it is")
            continue
        raw_date = at(row_dates, i) or form.get(f"group_date_{g}") or form.get("date")
        entry_date, date_error = parse_entry_date(raw_date)
        if date_error:
            problems.append(f"{name or student_id}: {date_error.rstrip('.').lower()}")
            continue
        planned.append((key, {
            "Date": entry_date, "Name": name, "StudentID": student_id,
            "House": at(houses, i), "Class": at(classes, i), "Remarks": at(remarks_list, i),
        }, g))

    if problems:
        unique = list(dict.fromkeys(problems))
        if job:
            # A folder batch is kept on the server: go back to it, not to a rescan.
            flash("Nothing was saved. Fix these, then save again: " + "; ".join(unique[:6])
                  + (f" (and {len(unique) - 6} more)" if len(unique) > 6 else "") + ".", "error")
            return redirect(url_for("scan_job", job_id=job["id"]))
        flash("Nothing was saved. Fix these and scan again: " + "; ".join(unique[:6])
              + (f" (and {len(unique) - 6} more)" if len(unique) > 6 else "") + ".", "error")
        first_key = form.get("group_category_0") or form.get("category")
        return redirect(url_for("scan", hint=first_key if first_key in CATEGORIES else ""))

    if not planned:
        flash("No rows were ticked, so nothing was saved.", "error")
        return redirect(url_for("scan_job", job_id=job["id"]) if job else url_for("scan"))

    per_register = Counter()
    dates = set()
    # One write for the whole batch (raises WorkbookLocked -> nothing saved).
    append_log_entries([(key, values) for key, values, _ in planned], logged_by, "OCR")
    for key, values, _ in planned:
        per_register[key] += 1
        dates.add(values["Date"])

    where = ", ".join(f"{n} in {CATEGORIES[k]['label']}" for k, n in per_register.items())
    when = (min(dates).strftime("%d %b %Y") if len(dates) == 1
            else f"{min(dates).strftime('%d %b')} to {max(dates).strftime('%d %b %Y')}")
    total = len(planned)
    message = f"Saved {total} entr{'y' if total == 1 else 'ies'} from the scan: {where}, dated {when}."

    if job:
        # Photos that produced saved rows move to <folder>/Scanned; unreadable
        # or untouched ones stay in the folder for the next scan.
        used = {int(g) for _, _, g in planned if str(g).isdigit()}
        done_paths = [p for i, p in enumerate(job["sources"]) if i in used]
        moved = move_to_scanned(done_paths, job["folder"])
        left = len(list_folder_photos(job["folder"]))
        message += f" {moved} photo{'' if moved == 1 else 's'} moved to the {SCANNED_SUBFOLDER} folder"
        message += f"; {left} still waiting in the scan folder." if left else "."
        with scan_jobs_lock:
            SCAN_JOBS.pop(job["id"], None)
        _delete_job_file(job["id"])

    flash(message, "success")
    top = per_register.most_common(1)[0][0]
    return redirect(url_for("entry", key=top))


@app.route("/students", methods=["GET", "POST"])
def students():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add":
            name = request.form.get("name", "").strip()
            if name:
                add_student(
                    name,
                    request.form.get("klass", "").strip(),
                    request.form.get("section", "").strip(),
                    request.form.get("roll", "").strip(),
                    request.form.get("house", "").strip(),
                )
                flash(f"Added {name} to the master list.", "success")
        elif action == "edit_all":
            ids = request.form.getlist("student_id")
            names = request.form.getlist("name")
            klasses = request.form.getlist("klass")
            sections = request.form.getlist("section")
            rolls = request.form.getlist("roll")
            houses = request.form.getlist("house")
            for i, sid in enumerate(ids):
                update_student(
                    sid,
                    names[i].strip() if i < len(names) else "",
                    klasses[i].strip() if i < len(klasses) else "",
                    sections[i].strip() if i < len(sections) else "",
                    rolls[i].strip() if i < len(rolls) else "",
                    houses[i].strip() if i < len(houses) else "",
                )
            flash(f"Updated {len(ids)} student record(s).", "success")
        return redirect(url_for("students"))

    return render_template("students.html", students=get_master_students(), categories=CATEGORIES)


# ---------------------------------------------------------------------------
# Find (who was late / who keeps turning up) and student profiles
# ---------------------------------------------------------------------------

FIND_PERIODS = {
    "all": "All time",
    "today": "Today",
    "week": "This week",
    "month": "This month",
    "year": "This year",
}
FIND_MIN_TIMES = [1, 2, 3, 5, 10]


def period_bounds(period, today=None):
    """(start, end) dates for a FIND_PERIODS key; (None, None) for all time."""
    today = today or date.today()
    if period == "today":
        return today, today
    if period == "week":
        return today - timedelta(days=today.weekday()), today
    if period == "month":
        return today.replace(day=1), today
    if period == "year":
        return today.replace(month=1, day=1), today
    return None, None


def all_entries(keys=None):
    """Every saved entry across the chosen registers, each tagged with its
    register key/label. Newest first."""
    out = []
    for key in keys or CATEGORIES:
        for row in read_log_rows(key, limit=100000):
            r = dict(row)
            r["_key"] = key
            r["_label"] = CATEGORIES[key]["label"]
            r["_date"] = _row_date(row)
            out.append(r)
    out.sort(key=lambda r: (r["_date"] or date.min, r.get("EntryID") or ""), reverse=True)
    return out


def _norm(value):
    return str(value or "").strip().lower()


def read_find_filters(args):
    register = args.get("register", "LateComers")
    if register != "all" and register not in CATEGORIES:
        register = "LateComers"
    period = args.get("period", "all")
    if period not in FIND_PERIODS:
        period = "all"
    try:
        min_times = max(1, int(args.get("min", 1)))
    except (TypeError, ValueError):
        min_times = 1
    return {
        "q": (args.get("q") or "").strip(),
        "klass": (args.get("klass") or "").strip(),
        "section": (args.get("section") or "").strip(),
        "house": (args.get("house") or "").strip(),
        "register": register,
        "period": period,
        "min": min_times,
    }


def find_students(filters, students=None):
    """One result per student who has matching entries, most entries first.
    The name box matches any part of the name or the student ID."""
    students = students if students is not None else get_master_students()
    id_by_name = {_norm(s["name"]): s["id"] for s in students}
    keys = list(CATEGORIES) if filters["register"] == "all" else [filters["register"]]
    start, end = period_bounds(filters["period"])
    q = _norm(filters["q"])

    groups = {}
    for row in all_entries(keys):
        name = (row.get("Name") or "").strip()
        if not name:
            continue
        sid = id_by_name.get(_norm(name), "")
        if q and q not in _norm(name) and q != _norm(sid):
            continue
        if filters["klass"] and _norm(row.get("Class")) != _norm(filters["klass"]):
            continue
        if filters["section"] and _norm(row.get("Section")) != _norm(filters["section"]):
            continue
        if filters["house"] and _norm(row.get("House")) != _norm(filters["house"]):
            continue
        d = row["_date"]
        if start and (d is None or not (start <= d <= end)):
            continue

        g = groups.setdefault(_norm(name), {
            "name": name,
            "id": sid,
            "class": _row_class_label(row),
            "house": row.get("House") or "",
            "count": 0,
            "by_register": {k: 0 for k in CATEGORIES},
            "last_date": None,
            "times": [],
            "last_detail": "",
        })
        g["count"] += 1
        g["by_register"][row["_key"]] += 1
        if g["last_date"] is None:  # rows arrive newest first
            g["last_date"] = d
            detail_field = CATEGORIES[row["_key"]]["extra_fields"][-1][0]
            g["last_detail"] = str(row.get(detail_field) or "")
        t = _parse_clock_time(row.get("Time")) if row["_key"] == "LateComers" else None
        if t:
            g["times"].append(t)

    results = []
    for g in groups.values():
        if g["count"] < filters["min"]:
            continue
        g["usual_time"] = _usual_time(g.pop("times"))
        results.append(g)
    results.sort(key=lambda g: (-g["count"], -(g["last_date"] or date.min).toordinal(), g["name"].lower()))
    return results


def _usual_time(times):
    """Median arrival time, e.g. '8:42 AM', or '' if none recorded."""
    if not times:
        return ""
    mins = sorted(t.hour * 60 + t.minute for t in times)
    m = mins[len(mins) // 2]
    return datetime.combine(date.today(), clock_time(m // 60, m % 60)).strftime("%I:%M %p").lstrip("0")


def filter_options(students):
    """Distinct classes, sections and houses for the filter dropdowns."""
    def distinct(field, key=None):
        return sorted({str(s.get(field) or "").strip() for s in students if str(s.get(field) or "").strip()}, key=key)
    # Classes: 6, 7, ... 10 before IB-1, IB-2 (numeric grades first, in number order).
    by_grade = lambda v: (0, int(v), "") if v.isdigit() else (1, 0, v)  # noqa: E731
    return {"classes": distinct("class", by_grade), "sections": distinct("section"), "houses": distinct("house")}


def student_profile_data(name):
    """Everything about one student. Returns None if they have neither a
    master-list record nor any entries."""
    students = get_master_students()
    record = resolve_student(name, students=students)
    display_name = record["name"] if record else (name or "").strip()
    entries = [r for r in all_entries() if _norm(r.get("Name")) == _norm(display_name)]
    if not record and not entries:
        return None

    today = date.today()
    month_start = today.replace(day=1)
    totals = {}
    for key, cfg in CATEGORIES.items():
        rows = [r for r in entries if r["_key"] == key]
        totals[key] = {
            "label": cfg["label"],
            "color": cfg["color"],
            "all": len(rows),
            "month": sum(1 for r in rows if r["_date"] and month_start <= r["_date"] <= today),
        }

    notes = []
    late = [r for r in entries if r["_key"] == "LateComers"]
    times = [t for t in (_parse_clock_time(r.get("Time")) for r in late) if t]
    if times:
        recorded = "1 time recorded" if len(times) == 1 else f"{len(times)} times recorded"
        notes.append(f"Usually arrives around {_usual_time(times)} when late ({recorded}).")
    dated = [r["_date"] for r in late if r["_date"]]
    if len(dated) >= 3:
        wd, n = Counter(d.weekday() for d in dated).most_common(1)[0]
        if n >= 2:
            notes.append(f"Late most often on {WEEKDAY_NAMES[wd]}s ({n} of {len(dated)}).")
    reasons = Counter(str(r.get("Reason") or "").strip() for r in late if str(r.get("Reason") or "").strip())
    if reasons:
        reason, n = reasons.most_common(1)[0]
        if n >= 2:
            notes.append(f"Most common late reason: “{reason}” ({n} times).")
    flags = [f for f in compute_weekly_repeat_flags({k: [r for r in entries if r["_key"] == k] for k in CATEGORIES})
             if _norm(f["name"]) == _norm(display_name)]
    if flags:
        notes.append(f"Logged 3+ times in one week on {len(flags)} occasion{'s' if len(flags) != 1 else ''} (latest: {flags[0]['week_label']}).")

    return {
        "name": display_name,
        "record": record,
        "class": class_label(record) if record else (_row_class_label(entries[0]) if entries else ""),
        "house": (record or {}).get("house") or (entries[0].get("House") if entries else ""),
        "entries": entries,
        "totals": totals,
        "total": len(entries),
        "notes": notes,
        "first_date": min((r["_date"] for r in entries if r["_date"]), default=None),
        "last_date": max((r["_date"] for r in entries if r["_date"]), default=None),
    }


def entry_detail(row):
    """The register-specific text of an entry, e.g. '8:50 AM · Missed bus'."""
    parts = [str(row.get(f) or "").strip() for f, _ in CATEGORIES[row["_key"]]["extra_fields"]]
    return " · ".join(p for p in parts if p)


def xlsx_response(title, headers, rows, filename, widths=None):
    """A small styled .xlsx built in memory, for the Download buttons."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="232323")
    ws["A2"] = f"Exported from MCTM Tracker on {date.today().strftime('%d %b %Y')}"
    ws["A2"].font = Font(name="Arial", size=9, color="686868")
    hdr_row = 4
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=hdr_row, column=c, value=h)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="232323")
        cell.border = Border(bottom=Side(style="medium", color="A1FFCB"))
        cell.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[hdr_row].height = 22
    line = Border(bottom=Side(style="thin", color="E6E6E3"))
    for r, values in enumerate(rows, start=hdr_row + 1):
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="Arial", size=10, color="232323")
            cell.border = line
            if isinstance(v, (date, datetime)):
                cell.number_format = "dd mmm yyyy"
                cell.alignment = Alignment(horizontal="left")
    for i, w in enumerate(widths or [18] * len(headers), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
    if rows:
        ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(headers))}{hdr_row + len(rows)}"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _find_title(filters):
    reg = "All registers" if filters["register"] == "all" else CATEGORIES[filters["register"]]["label"]
    bits = [reg, FIND_PERIODS[filters["period"]]]
    for label, key in (("Class", "klass"), ("Section", "section"), ("House", "house")):
        if filters[key]:
            bits.append(f"{label} {filters[key]}")
    if filters["q"]:
        bits.append(f"“{filters['q']}”")
    if filters["min"] > 1:
        bits.append(f"{filters['min']}+ times")
    return " · ".join(bits)


@app.route("/find")
def find():
    students = get_master_students()
    filters = read_find_filters(request.args)
    searched = any(k in request.args for k in ("q", "klass", "section", "house", "register", "period", "min"))
    results = find_students(filters, students)
    return render_template(
        "find.html", categories=CATEGORIES, filters=filters, results=results, searched=searched,
        options=filter_options(students), periods=FIND_PERIODS, min_choices=FIND_MIN_TIMES,
        title=_find_title(filters), query_string=request.query_string.decode(),
    )


@app.route("/find/export.xlsx")
def find_export():
    filters = read_find_filters(request.args)
    results = find_students(filters)
    headers = ["Student", "Student ID", "Class", "House", "Times"]
    if filters["register"] == "all":
        headers += [CATEGORIES[k]["label"] for k in CATEGORIES]
    headers += ["Last entry", "Usual arrival", "Latest reason / issue"]
    rows = []
    for g in results:
        values = [g["name"], g["id"], g["class"], g["house"], g["count"]]
        if filters["register"] == "all":
            values += [g["by_register"][k] for k in CATEGORIES]
        values += [g["last_date"], g["usual_time"], g["last_detail"]]
        rows.append(values)
    widths = [24, 12, 9, 13, 8] + ([12] * len(CATEGORIES) if filters["register"] == "all" else []) + [13, 13, 30]
    return xlsx_response(f"Find: {_find_title(filters)}", headers, rows,
                         f"MCTM find {date.today().isoformat()}.xlsx", widths)


@app.route("/student")
def student_profile():
    data = student_profile_data(request.args.get("name", ""))
    if data is None:
        flash(f"No student called “{request.args.get('name', '')}” in the master list or the registers.", "error")
        return redirect(url_for("find"))
    return render_template("student.html", categories=CATEGORIES, s=data, entry_detail=entry_detail, today=date.today())


@app.route("/student/export.xlsx")
def student_export():
    data = student_profile_data(request.args.get("name", ""))
    if data is None:
        return redirect(url_for("find"))
    rows = [[r["_date"], r["_label"], entry_detail(r), r.get("Remarks") or "", r.get("LoggedBy") or "",
             "Scan" if r.get("Source") == "OCR" else (r.get("Source") or ""), r.get("EntryID") or ""]
            for r in data["entries"]]
    title = f"{data['name']}" + (f" ({data['record']['id']})" if data["record"] else "") + f", {data['class']}"
    safe = re.sub(r"[^A-Za-z0-9 _-]", "", data["name"]).strip() or "student"
    return xlsx_response(title, ["Date", "Register", "Time / reason / issue", "Remarks", "Logged by", "Source", "Entry"],
                         rows, f"MCTM {safe}.xlsx", [13, 14, 32, 30, 15, 9, 11])


@app.template_global()
def student_url(name):
    return url_for("student_profile", name=name) if name else "#"


if __name__ == "__main__":
    if not os.path.exists(DB_PATH):
        print(f"Database not found at {DB_PATH}.")
        print("Run: python build_excel_template.py")
    else:
        print(f"Using database: {DB_PATH}")
        print(f"OCR backend: {ocr_backend_label()} (ready: {ocr_ready()})")
        if USING_GEMINI:
            fallback = "Tesseract" if OCR_AVAILABLE else "nothing (Tesseract not installed)"
            print(f"OCR model order: {' -> '.join(GEMINI_MODELS)} -> {fallback}")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
