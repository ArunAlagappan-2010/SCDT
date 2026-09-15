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
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash").strip()
GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/interactions"
GEMINI_OCR_PROMPT = (
    "This is a photo of a handwritten or printed school register page. "
    "Transcribe every line of text exactly as written, one line per line of "
    "the original. Output ONLY the transcribed text, no commentary, no markdown."
)
USING_GEMINI = bool(GEMINI_API_KEY)

# Must match the column order build_excel_template.py wrote into each sheet.
CATEGORIES = {
    "LateComers": {
        "label": "Late Comers",
        "color": "#2F5C8A",
        "light": "#DCE6F1",
        "extra_fields": [("Time", "text"), ("Reason", "text")],
    },
    "Defaulters": {
        "label": "Defaulters",
        "color": "#C0392B",
        "light": "#FADBD8",
        "extra_fields": [("Reason", "text")],
    },
    "Uniform": {
        "label": "Uniform",
        "color": "#D68910",
        "light": "#FDEBD0",
        "extra_fields": [("Issue", "text")],
    },
    "NailsHair": {
        "label": "Nails & Hair",
        "color": "#1E8449",
        "light": "#D5F5E3",
        "extra_fields": [("Issue", "text")],
    },
}

MAX_LOG_ROWS = 1000
MAX_MASTER_ROWS = 2000


def sheet_columns(key):
    cfg = CATEGORIES[key]
    return ["EntryID", "Date", "Name", "Class", "Section"] + [f[0] for f in cfg["extra_fields"]] + [
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


def add_student(name, klass, section, roll):
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
        wb.save(DB_PATH)
        return sid


def update_student(student_id, name, klass, section, roll):
    with wb_lock:
        wb = open_wb()
        ws = wb["MasterList"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == student_id:
                ws.cell(row=r, column=2, value=name)
                ws.cell(row=r, column=3, value=klass)
                ws.cell(row=r, column=4, value=section)
                ws.cell(row=r, column=5, value=roll)
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

    return summary, repeat_list, max_month, recent


# ---------------------------------------------------------------------------
# OCR helpers
# ---------------------------------------------------------------------------

def run_ocr(image_path):
    if USING_GEMINI:
        return run_ocr_gemini(image_path)
    return run_ocr_tesseract(image_path)


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
            # lower latency and cost (this API can otherwise take 20s+ even on
            # trivial requests).
            "generation_config": {"thinking_level": "low"},
        }
        resp = requests.post(
            GEMINI_ENDPOINT,
            headers={"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"},
            json=payload,
            timeout=90,
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


def parse_ocr_lines(raw_text, students):
    names = [s["name"] for s in students]
    lines = [ln.strip() for ln in raw_text.splitlines()]
    lines = [re.sub(r"^[\W_]+", "", ln) for ln in lines]  # strip leading bullets/numbers/punct
    lines = [ln for ln in lines if len(ln) >= 2]

    candidates = []
    for ln in lines:
        best = None
        if names:
            matches = difflib.get_close_matches(ln, names, n=1, cutoff=0.45)
            if matches:
                best = matches[0]
        candidates.append({
            "raw_text": ln,
            "guess_name": best or "",
            "confident": bool(best),
        })
    return candidates


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def dashboard():
    summary, repeat_list, max_month, recent = dashboard_data()
    return render_template(
        "dashboard.html",
        categories=CATEGORIES,
        summary=summary,
        repeat_list=repeat_list,
        max_month=max_month,
        recent=recent,
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


@app.route("/scan/<key>", methods=["GET", "POST"])
def scan(key):
    if key not in CATEGORIES:
        return "Unknown category", 404
    cfg = CATEGORIES[key]

    if request.method == "POST":
        file = request.files.get("photo")
        if not file or file.filename == "":
            flash("Please choose a photo of the register page.", "error")
            return redirect(url_for("scan", key=key))

        filename = f"{key}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}"
        save_path = os.path.join(UPLOAD_DIR, filename)
        file.save(save_path)

        raw_text, error = run_ocr(save_path)
        students = get_master_students()
        candidates = parse_ocr_lines(raw_text, students) if raw_text else []

        return render_template(
            "scan_review.html",
            key=key,
            cfg=cfg,
            categories=CATEGORIES,
            candidates=candidates,
            raw_text=raw_text,
            error=error,
            ocr_available=ocr_ready(),
            ocr_backend=ocr_backend_label(),
        )

    return render_template(
        "scan.html", key=key, cfg=cfg, categories=CATEGORIES,
        ocr_available=ocr_ready(), ocr_backend=ocr_backend_label(),
    )


@app.route("/scan/<key>/commit", methods=["POST"])
def scan_commit(key):
    if key not in CATEGORIES:
        return "Unknown category", 404
    cfg = CATEGORIES[key]

    names = request.form.getlist("row_name")
    includes = request.form.getlist("row_include")  # values are the row indices that were checked
    extra_lists = {field: request.form.getlist(f"row_{field}") for field, _ in cfg["extra_fields"]}
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
        values = {"Name": name, "Remarks": remarks_list[i] if i < len(remarks_list) else ""}
        for field, _ in cfg["extra_fields"]:
            lst = extra_lists[field]
            values[field] = lst[i] if i < len(lst) else ""
        append_log_entry(key, values, logged_by, "OCR")
        saved += 1

    flash(f"Saved {saved} entr{'y' if saved == 1 else 'ies'} from the scan.", "success")
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
                )
                flash(f"Added {name} to the master list.", "success")
        elif action == "edit_all":
            ids = request.form.getlist("student_id")
            names = request.form.getlist("name")
            klasses = request.form.getlist("klass")
            sections = request.form.getlist("section")
            rolls = request.form.getlist("roll")
            for i, sid in enumerate(ids):
                update_student(
                    sid,
                    names[i].strip() if i < len(names) else "",
                    klasses[i].strip() if i < len(klasses) else "",
                    sections[i].strip() if i < len(sections) else "",
                    rolls[i].strip() if i < len(rolls) else "",
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
