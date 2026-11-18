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
# Optional: point pytesseract at the Tesseract engine if it's not on PATH.
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
try:
    import pytesseract
    if os.path.exists(TESSERACT_CMD):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    OCR_AVAILABLE = True
except ImportError:
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
        "If a value is missing or unreadable for a real row, use an empty string "
        "for that field, but still include the row if at least the name is readable. "
        "Respond with ONLY a raw JSON object, no markdown code fences, no commentary, "
        "in exactly this shape: "
        '{"category_guess": "...", "rows": [{"house": "...", "name": "...", "class": "..."}, ...]}'
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


def get_master_students():
    """Returns list of dicts for every student in MasterList."""
    with wb_lock:
        wb = open_wb()
        ws = wb["MasterList"]
        students = []
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
            sid = row[0].value
            name = row[1].value
            if not sid or not name:
                continue
            students.append({
                "id": sid,
                "name": name,
                "class": row[2].value or "",
                "section": row[3].value or "",
                "roll": row[4].value or "",
                "house": (row[5].value or "") if len(row) > 5 else "",
            })
        return students


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
        "House": match["house"] if match else values_by_header.get("House", ""),
        "Remarks": values_by_header.get("Remarks", ""),
        "LoggedBy": logged_by or "",
    }
    for field, _ in CATEGORIES[key]["extra_fields"]:
        row_values[field] = values_by_header.get(field, "")
    return row_values


def append_log_entry(key, values_by_header, logged_by, source):
    """values_by_header: dict covering Date, Name (or StudentID), <extra fields>, Remarks."""
    cols = sheet_columns(key)
    row_values = _entry_row_values(key, values_by_header, logged_by, get_master_students())

    with wb_lock:
        wb = open_wb()
        ws = wb[key]
        r = next_log_row(ws)
        entry_id = next_log_entry_id(ws, key)
        row_values["EntryID"] = entry_id
        row_values["Source"] = source

        for c, header in enumerate(cols, start=1):
            ws.cell(row=r, column=c, value=row_values.get(header, ""))

        save_wb(wb)
        return entry_id


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
    cols = sheet_columns(key)
    with wb_lock:
        wb = open_wb()
        ws = wb[key]
        rows = []
        for r in range(2, ws.max_row + 1):
            entry_id = ws.cell(row=r, column=1).value
            if not entry_id:
                continue
            row = {}
            for c in range(1, len(cols) + 1):
                v = ws.cell(row=r, column=c).value
                # openpyxl round-trips empty-string writes as None; normalize back to "".
                row[cols[c - 1]] = "" if v is None else v
            rows.append(row)
        rows.sort(key=lambda row: (row.get("Date") or date.min, row.get("EntryID") or ""), reverse=True)
        return rows[:limit]


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

    stats = {
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
    return "Gemini (cloud)" if USING_GEMINI else "Tesseract (offline)"


def run_ocr_tesseract(image_path):
    if not OCR_AVAILABLE:
        return "", "pytesseract is not installed on the server."
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


def match_name(name, names):
    if not name:
        return None
    matches = difflib.get_close_matches(name, names, n=1, cutoff=0.45)
    return matches[0] if matches else None


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
        })
    return category_guess, result


def parse_structured_ocr_json(text):
    """Convenience wrapper around parse_gemini_extraction: just the rows."""
    _, rows = parse_gemini_extraction(text)
    return rows


def _make_candidate(raw_text, house, name, klass, students):
    """One review-table row. If the name matches a student in the master
    list, House/Class/ID come from that record -- what OCR read for those
    columns is only kept for students it couldn't identify."""
    best = match_name(name, [s["name"] for s in students])
    student = find_student_by_name(best, students) if best else None
    return {
        "raw_text": raw_text,
        "guess_house": student["house"] if student else house,
        "guess_name": best or name,
        "guess_class": class_label(student) if student else klass,
        "student_id": student["id"] if student else "",
        "confident": bool(best),
    }


def build_candidates_from_structured(rows, students):
    candidates = []
    for row in rows:
        name = row["name"]
        if not name or name.strip().lower() in HEADER_WORDS:
            continue  # no name to work with, or the model slipped in a header row anyway
        raw_text = " / ".join(v for v in (row["house"], name, row["klass"]) if v)
        candidates.append(_make_candidate(raw_text, row["house"], name, row["klass"], students))
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
    for ln in lines:
        parts = split_columns(ln)
        if is_header_like(parts if len(parts) > 1 else ln.split()):
            continue  # this line is the column-title row, not a student

        if len(parts) >= 3:
            house, name, klass = parts[0], parts[1], parts[2]
        elif len(parts) == 2:
            house, name, klass = "", parts[0], parts[1]
        else:
            house, name, klass = "", ln, ""

        candidates.append(_make_candidate(ln, house, name, klass, students))
    return candidates


def guess_category_from_text_lines(raw_text):
    """Best-effort register-type detection for the Tesseract (plain text)
    path: checks the first few lines for something that looks like a title
    matching one of the known categories (e.g. "UNIFORM REGISTER")."""
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()][:5]
    for ln in lines:
        key = match_category_label(ln)
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
        file = request.files.get("photo")
        if not file or file.filename == "":
            flash("Please choose a photo of the register page.", "error")
            return redirect(url_for("scan"))

        hint_key = request.form.get("hint")
        hint_key = hint_key if hint_key in CATEGORIES else None

        filename = f"scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}"
        save_path = os.path.join(UPLOAD_DIR, filename)
        file.save(save_path)

        raw_text, error, warning, backend = run_ocr(save_path)
        students = get_master_students()
        candidates, category_guess = extract_candidates(raw_text, students, backend)
        detected_key = category_guess or hint_key

        return render_template(
            "scan_review.html",
            categories=CATEGORIES,
            detected_key=detected_key,
            category_was_detected=bool(category_guess),
            candidates=candidates,
            raw_text=raw_text,
            error=error,
            warning=warning,
            ocr_available=ocr_ready(),
            ocr_backend=ocr_backend_label(),
            today=date.today(),
        )

    hint_key = request.args.get("hint")
    hint_key = hint_key if hint_key in CATEGORIES else ""
    return render_template(
        "scan.html", categories=CATEGORIES, hint_key=hint_key,
        ocr_available=ocr_ready(), ocr_backend=ocr_backend_label(),
    )


@app.route("/scan/commit", methods=["POST"])
def scan_commit():
    key = request.form.get("category")
    if key not in CATEGORIES:
        flash("Please choose which register this scan belongs to before saving.", "error")
        return redirect(url_for("scan"))

    entry_date, date_error = parse_entry_date(request.form.get("date"))
    if date_error:
        flash(f"{date_error} Nothing was saved -- please scan the page again.", "error")
        return redirect(url_for("scan", hint=key))

    names = request.form.getlist("row_name")
    student_ids = request.form.getlist("row_student_id")
    houses = request.form.getlist("row_house")
    classes = request.form.getlist("row_class")
    includes = request.form.getlist("row_include")  # values are the row indices that were checked
    remarks_list = request.form.getlist("row_remarks")
    logged_by = request.form.get("logged_by", "").strip()

    saved = 0
    included = set(int(i) for i in includes)
    for i, name in enumerate(names):
        if i not in included:
            continue
        name = name.strip()
        student_id = student_ids[i].strip() if i < len(student_ids) else ""
        if not name and not student_id:
            continue
        values = {
            "Date": entry_date,
            "Name": name,
            "StudentID": student_id,
            "House": houses[i].strip() if i < len(houses) else "",
            "Class": classes[i].strip() if i < len(classes) else "",
            "Remarks": remarks_list[i] if i < len(remarks_list) else "",
        }
        append_log_entry(key, values, logged_by, "OCR")
        saved += 1

    flash(
        f"Saved {saved} entr{'y' if saved == 1 else 'ies'} from the scan into "
        f"{CATEGORIES[key]['label']}, dated {entry_date.strftime('%d %b %Y')}.",
        "success",
    )
    return redirect(url_for("entry", key=key))


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
