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

import os
import re
import json
import math
import base64
import difflib
import threading
from datetime import datetime, date, timedelta

import requests
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, jsonify, flash
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
GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/interactions"
USING_GEMINI = bool(GEMINI_API_KEY)

# Must match the column order build_excel_template.py wrote into each sheet.
# Colors are all drawn from one grey/green family (minimalist palette) --
# categories are told apart by lightness/shade, not by clashing hues.
CATEGORIES = {
    "LateComers": {
        "label": "Late Comers",
        "color": "#2F6FD6",
        "light": "#DCE7FA",
        "extra_fields": [("Time", "text"), ("Reason", "text")],
    },
    "Defaulters": {
        "label": "Defaulters",
        "color": "#E4432D",
        "light": "#FBDEDA",
        "extra_fields": [("Reason", "text")],
    },
    "Uniform": {
        "label": "Uniform",
        "color": "#F0A202",
        "light": "#FCEBCB",
        "extra_fields": [("Issue", "text")],
    },
    "NailsHair": {
        "label": "Nails & Hair",
        "color": "#9B51E0",
        "light": "#EEE0F9",
        "extra_fields": [("Issue", "text")],
    },
}

CATEGORY_LABELS = [cfg["label"] for cfg in CATEGORIES.values()]


def _gemini_ocr_prompt():
    labels = ", ".join(f'"{cfg["label"]}"' for cfg in CATEGORIES.values())
    return (
        "This is a photo of a handwritten or printed school register page with "
        "columns including House, Name, and Class (there may be other columns too "
        "-- ignore those). "
        "First, look at the page's title/heading (if there is one) and guess which "
        f"kind of register this is, from exactly these options: {labels}. If you "
        "can't tell, use an empty string. "
        "Then, for every row that is an actual student entry, extract its House, "
        "Name, and Class. "
        "Do NOT include the header/title row -- that is the row that just repeats "
        "the column labels themselves (words like \"House\", \"Name\", \"Class\", "
        "\"Section\", \"Roll\", \"S.No\", \"Date\" etc used as headings, not as a "
        "real student's data) -- skip it entirely, it is not a student. "
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

def open_wb():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(
            f"Database not found at {DB_PATH}. Run build_excel_template.py first."
        )
    return load_workbook(DB_PATH)


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
        wb.save(DB_PATH)
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
                wb.save(DB_PATH)
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


def append_log_entry(key, values_by_header, logged_by, source):
    """values_by_header: dict covering Date, Name, <extra fields>, Remarks."""
    cols = sheet_columns(key)
    students = get_master_students()
    match = find_student_by_name(values_by_header.get("Name", ""), students)

    with wb_lock:
        wb = open_wb()
        ws = wb[key]
        r = next_log_row(ws)
        entry_id = next_log_entry_id(ws, key)

        row_values = {
            "EntryID": entry_id,
            "Date": values_by_header.get("Date") or date.today(),
            "Name": values_by_header.get("Name", ""),
            "Class": match["class"] if match else values_by_header.get("Class", ""),
            "Section": match["section"] if match else values_by_header.get("Section", ""),
            "House": match["house"] if match else values_by_header.get("House", ""),
            "Remarks": values_by_header.get("Remarks", ""),
            "LoggedBy": logged_by or "",
            "Source": source,
        }
        for field, _ in CATEGORIES[key]["extra_fields"]:
            row_values[field] = values_by_header.get(field, "")

        for c, header in enumerate(cols, start=1):
            ws.cell(row=r, column=c, value=row_values.get(header, ""))

        wb.save(DB_PATH)
        return entry_id


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


def dashboard_data():
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    month_start = today.replace(day=1)

    summary = {}
    all_rows_by_cat = {}
    for key, cfg in CATEGORIES.items():
        rows = read_log_rows(key, limit=100000)
        all_rows_by_cat[key] = rows
        summary[key] = {
            "label": cfg["label"],
            "color": cfg["color"],
            "light": cfg["light"],
            "today": count_in_range(rows, today, today),
            "week": count_in_range(rows, week_start, today),
            "month": count_in_range(rows, month_start, today),
            "total": len(rows),
        }

    # Repeat offenders: tally by name across all four categories.
    tally = {}
    for key, rows in all_rows_by_cat.items():
        for row in rows:
            name = (row.get("Name") or "").strip()
            if not name:
                continue
            entry = tally.setdefault(name, {"name": name, "class": row.get("Class") or "", "counts": {k: 0 for k in CATEGORIES}})
            entry["counts"][key] += 1
    repeat_list = sorted(tally.values(), key=lambda e: sum(e["counts"].values()), reverse=True)
    for e in repeat_list:
        e["total"] = sum(e["counts"].values())
    repeat_list = [e for e in repeat_list if e["total"] > 0][:10]

    max_month = max((s["month"] for s in summary.values()), default=0) or 1

    # Overall (all categories combined) this-month vs last-month, for the
    # big headline number + trend + delta card.
    last_month_end = month_start - timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)
    overall_month_total = sum(s["month"] for s in summary.values())
    overall_last_month_total = sum(
        count_in_range(rows, last_month_start, last_month_end) for rows in all_rows_by_cat.values()
    )
    overall_month_delta = overall_month_total - overall_last_month_total

    recent = []
    for key, rows in all_rows_by_cat.items():
        for row in rows[:15]:
            recent.append({
                "category": CATEGORIES[key]["label"],
                "color": CATEGORIES[key]["color"],
                "date": row.get("Date"),
                "name": row.get("Name"),
                "detail": next((row.get(f[0]) for f in CATEGORIES[key]["extra_fields"] if row.get(f[0])), "") or (row.get("Remarks") or ""),
                "source": row.get("Source"),
            })
    recent.sort(key=lambda r: r["date"] or date.min, reverse=True)
    recent = recent[:12]

    stats = build_extra_stats(all_rows_by_cat)
    stats["overall_month_total"] = overall_month_total
    stats["overall_month_delta"] = overall_month_delta
    stats["top_student"] = repeat_list[0] if repeat_list else None

    return summary, repeat_list, max_month, recent, stats


def build_extra_stats(all_rows_by_cat):
    """Deeper statistics for the dashboard's expandable analytics panel --
    computed once here so the main KPI view stays cheap and this only runs
    when the page is actually requested (the panel itself is revealed on
    click client-side, but the data still needs to be ready when it opens)."""
    today = date.today()

    # 14-day trend per category, for a small line/bar chart.
    days = [today - timedelta(days=i) for i in range(13, -1, -1)]
    trends = {}
    for key, rows in all_rows_by_cat.items():
        counts_by_day = {d: 0 for d in days}
        for row in rows:
            d = row.get("Date")
            if isinstance(d, datetime):
                d = d.date()
            if d in counts_by_day:
                counts_by_day[d] += 1
        trends[key] = [{"label": d.strftime("%d %b"), "count": counts_by_day[d]} for d in days]
    trend_max = max((point["count"] for series in trends.values() for point in series), default=0) or 1

    # Class / House / Source breakdowns across all categories combined.
    class_counts, house_counts, source_counts = {}, {}, {"Manual": 0, "OCR": 0}
    for rows in all_rows_by_cat.values():
        for row in rows:
            klass = (row.get("Class") or "").strip()
            if klass:
                class_counts[klass] = class_counts.get(klass, 0) + 1
            house = (row.get("House") or "").strip()
            if house:
                house_counts[house] = house_counts.get(house, 0) + 1
            src = row.get("Source")
            if src in source_counts:
                source_counts[src] += 1

    def top_n(counts, n=8):
        items = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:n]
        top_max = max((v for _, v in items), default=0) or 1
        return [{"label": k, "count": v} for k, v in items], top_max

    class_breakdown, class_max = top_n(class_counts)
    house_breakdown, house_max = top_n(house_counts)
    total_sourced = sum(source_counts.values()) or 1
    source_ocr_pct = round(source_counts["OCR"] / total_sourced * 100)

    # Combined (all categories) trend, for the headline sparkline card.
    overall_trend = [
        {"label": days[i].strftime("%d %b"), "count": sum(trends[key][i]["count"] for key in trends)}
        for i in range(len(days))
    ]
    overall_trend_max = max((p["count"] for p in overall_trend), default=0) or 1

    return {
        "trends": trends,
        "trend_max": trend_max,
        "overall_trend": overall_trend,
        "overall_trend_max": overall_trend_max,
        "class_breakdown": class_breakdown,
        "class_max": class_max,
        "house_breakdown": house_breakdown,
        "house_max": house_max,
        "source_counts": source_counts,
        "source_manual_pct": round(source_counts["Manual"] / total_sourced * 100),
        "source_ocr_pct": source_ocr_pct,
        "ocr_gauge": gauge_geometry(source_ocr_pct),
    }


def gauge_geometry(value, cx=100, cy=100, r=70):
    """Needle endpoint for a 0-100 semicircle gauge (0 = left/180deg, 100 =
    right/0deg, sweeping over the top), plus the value clamped to [0,100]."""
    value = max(0, min(100, value))
    theta = math.pi * (1 - value / 100)
    return {
        "value": value,
        "needle_x": round(cx + r * math.cos(theta), 1),
        "needle_y": round(cy - r * math.sin(theta), 1),
    }


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


def run_ocr_gemini(image_path):
    try:
        img = Image.open(image_path)
        img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")

        import io
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=90)
        b64_image = base64.b64encode(buf.getvalue()).decode("ascii")

        payload = {
            "model": GEMINI_MODEL,
            "input": [
                {"type": "text", "text": GEMINI_OCR_PROMPT},
                {"type": "image", "data": b64_image, "mime_type": "image/jpeg"},
            ],
            # OCR is a simple extraction task -- skip extended "thinking" for
            # lower latency and cost. gemini-3.5-flash-lite already has thinking
            # off by default, but this is a harmless safety net if the model
            # or its defaults ever change.
            "generation_config": {"thinking_level": "low"},
        }
        resp = requests.post(
            GEMINI_ENDPOINT,
            headers={"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"},
            json=payload,
            # flash-lite typically responds in 5-10s; 45s leaves generous
            # headroom for a slow moment while still failing (and falling
            # back to Tesseract) well before staff start wondering if it's
            # frozen.
            timeout=45,
        )
        if resp.status_code != 200:
            return "", f"Gemini API error {resp.status_code}: {resp.text[:300]}"

        text = parse_gemini_response(resp.json())
        if not text:
            return "", "Gemini returned no text (unexpected response shape)."
        return text, None
    except requests.exceptions.RequestException as exc:
        return "", f"Could not reach Gemini API (check internet connection): {exc}"
    except Exception as exc:
        return "", f"Gemini OCR failed: {exc}"


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


def build_candidates_from_structured(rows, students):
    names = [s["name"] for s in students]
    candidates = []
    for row in rows:
        name = row["name"]
        if not name or name.strip().lower() in HEADER_WORDS:
            continue  # no name to work with, or the model slipped in a header row anyway
        best = match_name(name, names)
        candidates.append({
            "raw_text": " / ".join(v for v in (row["house"], name, row["klass"]) if v),
            "guess_house": row["house"],
            "guess_name": best or name,
            "guess_class": row["klass"],
            "confident": bool(best),
        })
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
    names = [s["name"] for s in students]
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

        best = match_name(name, names)
        candidates.append({
            "raw_text": ln,
            "guess_house": house,
            "guess_name": best or name,
            "guess_class": klass,
            "confident": bool(best),
        })
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
    summary, repeat_list, max_month, recent, stats = dashboard_data()
    return render_template(
        "dashboard.html",
        categories=CATEGORIES,
        summary=summary,
        repeat_list=repeat_list,
        max_month=max_month,
        recent=recent,
        stats=stats,
        today=date.today(),
    )


@app.route("/api/students")
def api_students():
    q = (request.args.get("q") or "").strip().lower()
    students = get_master_students()
    if not q:
        results = students[:8]
    else:
        starts = [s for s in students if s["name"].lower().startswith(q)]
        contains = [s for s in students if q in s["name"].lower() and s not in starts]
        results = (starts + contains)[:8]
    return jsonify(results)


@app.route("/entry/<key>", methods=["GET", "POST"])
def entry(key):
    if key not in CATEGORIES:
        return "Unknown category", 404
    cfg = CATEGORIES[key]

    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            flash("Please pick a student name.", "error")
            return redirect(url_for("entry", key=key))

        values = {"Name": name, "Remarks": request.form.get("remarks", "").strip()}
        for field, _ in cfg["extra_fields"]:
            values[field] = request.form.get(field, "").strip()

        entry_id = append_log_entry(key, values, request.form.get("logged_by", "").strip(), "Manual")
        flash(f"Saved ({entry_id}) for {name}.", "success")
        return redirect(url_for("entry", key=key))

    recent_rows = read_log_rows(key, limit=15)
    return render_template("entry.html", key=key, cfg=cfg, categories=CATEGORIES, recent_rows=recent_rows)


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

    names = request.form.getlist("row_name")
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
        if not name:
            continue
        values = {
            "Name": name,
            "House": houses[i].strip() if i < len(houses) else "",
            "Class": classes[i].strip() if i < len(classes) else "",
            "Remarks": remarks_list[i] if i < len(remarks_list) else "",
        }
        append_log_entry(key, values, logged_by, "OCR")
        saved += 1

    flash(f"Saved {saved} entr{'y' if saved == 1 else 'ies'} from the scan into {CATEGORIES[key]['label']}.", "success")
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


if __name__ == "__main__":
    if not os.path.exists(DB_PATH):
        print(f"Database not found at {DB_PATH}.")
        print("Run: python build_excel_template.py")
    else:
        print(f"Using database: {DB_PATH}")
        print(f"OCR backend: {ocr_backend_label()} (ready: {ocr_ready()})")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
