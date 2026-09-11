"""
Automated smoke tests for the Discipline Tracker app.

Runs against a throwaway COPY of the Excel database (never touches the real
data/SchoolDisciplineSystem.xlsx), using Flask's test client so no server
needs to be running.

Run:
    python tests/test_app.py
"""

import os
import sys
import shutil
import traceback
import tempfile
import io

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REAL_DB = os.path.join(BASE_DIR, "data", "SchoolDisciplineSystem.xlsx")

TMP_DIR = tempfile.mkdtemp(prefix="discipline_tracker_test_")
TEST_DB = os.path.join(TMP_DIR, "TestDatabase.xlsx")
shutil.copy(REAL_DB, TEST_DB)

import app as appmod  # noqa: E402

appmod.DB_PATH = TEST_DB
appmod.UPLOAD_DIR = TMP_DIR
appmod.app.config["TESTING"] = True

client = appmod.app.test_client()

results = []


def check(name, condition, detail=""):
    results.append((name, bool(condition), detail))
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {name}" + (f" -- {detail}" if detail and not condition else ""))


def run():
    initial_summary, _, _, _ = appmod.dashboard_data()
    initial_total = sum(s["total"] for s in initial_summary.values())

    # 1. Dashboard loads
    r = client.get("/")
    check("dashboard loads (200)", r.status_code == 200)
    check("dashboard shows title", b"Discipline &amp; Grooming Dashboard" in r.data or b"Discipline & Grooming Dashboard" in r.data)

    # 2. Student search API
    r = client.get("/api/students?q=")
    data = r.get_json()
    check("api/students empty query returns list", r.status_code == 200 and isinstance(data, list))
    check("api/students seed data has 5 students", len(data) == 5, f"got {len(data)}")

    r = client.get("/api/students?q=Rohan")
    data = r.get_json()
    check("api/students matches 'Rohan'", len(data) == 1 and data[0]["name"] == "Rohan Gupta", str(data))

    r = client.get("/api/students?q=zzz_no_such_name")
    check("api/students no-match returns empty list", r.get_json() == [])

    # 3. Unknown category returns 404
    r = client.get("/entry/NotACategory")
    check("unknown category -> 404", r.status_code == 404)

    # 4. Entry pages load for every category
    for key in appmod.CATEGORIES:
        r = client.get(f"/entry/{key}")
        check(f"GET /entry/{key} loads", r.status_code == 200)
        r = client.get(f"/scan/{key}")
        check(f"GET /scan/{key} loads", r.status_code == 200)

    # 5. Manual entry submission for each category, verifying category-specific fields.
    # Uses before/after deltas so this doesn't assume a pristine fixture.
    submissions = {
        "LateComers": {"name": "Rohan Gupta", "Time": "8:45 AM", "Reason": "Bus delay", "remarks": "", "logged_by": "Mrs. Rao"},
        "Defaulters": {"name": "Diya Patel", "Reason": "Missing homework", "remarks": "Second time", "logged_by": "Mr. Iyer"},
        "Uniform": {"name": "Kabir Singh", "Issue": "No tie", "remarks": "", "logged_by": ""},
        "NailsHair": {"name": "Meera Nair", "Issue": "Nail polish", "remarks": "", "logged_by": "Mrs. Rao"},
    }
    for key, form in submissions.items():
        before = len(appmod.read_log_rows(key, limit=10000))
        r = client.post(f"/entry/{key}", data=form, follow_redirects=True)
        check(f"POST /entry/{key} succeeds", r.status_code == 200)
        rows = appmod.read_log_rows(key, limit=10000)
        check(f"{key} row was written", len(rows) == before + 1, f"before={before} after={len(rows)}")
        new_row = next((row for row in rows if row.get("Name") == form["name"]), None)
        check(f"{key} new row found", new_row is not None, str(rows))
        if new_row:
            check(f"{key} row has correct class (auto-matched)", new_row.get("Class") in ("9", "10"), new_row.get("Class"))
            check(f"{key} row Source is Manual", new_row.get("Source") == "Manual")
            check(f"{key} no field left as Python None", all(v is not None for v in new_row.values()), str(new_row))

    # 6. Entry with missing name should not save and should redirect back
    before = len(appmod.read_log_rows("LateComers", limit=10000))
    r = client.post("/entry/LateComers", data={"name": "", "logged_by": "x"}, follow_redirects=True)
    rows = appmod.read_log_rows("LateComers", limit=10000)
    check("empty name is rejected (no extra row)", len(rows) == before, f"before={before} after={len(rows)}")

    # 7. Dashboard reflects the new entries
    r = client.get("/")
    check("dashboard reflects entries", b"Rohan Gupta" in r.data)

    # 8. Scan upload path: synthetic printed-text image, exercise OCR or graceful fallback
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (500, 200), "white")
    d = ImageDraw.Draw(img)
    d.text((10, 10), "Meera Nair", fill="black")
    d.text((10, 50), "Kabir Singh", fill="black")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)

    r = client.post(
        "/scan/Uniform",
        data={"photo": (buf, "test_register.png")},
        content_type="multipart/form-data",
    )
    check("POST /scan/<key> with image returns 200", r.status_code == 200)

    def tesseract_binary_available():
        if not appmod.OCR_AVAILABLE:
            return False
        try:
            appmod.pytesseract.get_tesseract_version()
            return True
        except Exception:
            return False

    # Gemini makes a real network call -- only assert the offline-fallback
    # behaviour deterministically; when Gemini is configured, just confirm the
    # request didn't crash the page (network/key issues surface as flash text).
    if appmod.USING_GEMINI:
        check("scan review renders when using Gemini backend", r.status_code == 200)
    elif tesseract_binary_available():
        check("scan review mentions OCR-unavailable warning is ABSENT (engine installed)", b"OCR could not run" not in r.data)
    else:
        check("scan review shows OCR-unavailable warning when engine missing", b"OCR could not run" in r.data or b"No text was detected" in r.data)

    # 9. Scan commit path (independent of what OCR actually produced -- simulates the
    #    human review/correction step, which is the important safety net).
    r = client.post(
        "/scan/Uniform/commit",
        data={
            "row_name": ["Aarav Sharma", "Diya Patel"],
            "row_include": ["0"],  # only the first row is checked
            "row_Issue": ["Shoes", "Blazer"],
            "row_remarks": ["", ""],
            "logged_by": "Scan-Test",
        },
        follow_redirects=True,
    )
    check("POST /scan/<key>/commit succeeds", r.status_code == 200)
    uniform_rows = appmod.read_log_rows("Uniform", limit=10)
    check("scan commit saved exactly the checked row", len(uniform_rows) == 2, f"rows={uniform_rows}")
    ocr_rows = [row for row in uniform_rows if row.get("Source") == "OCR"]
    check("scan commit row has Source=OCR", len(ocr_rows) == 1 and ocr_rows[0]["Name"] == "Aarav Sharma", str(ocr_rows))

    # 10. Master student list: add + edit_all
    r = client.get("/students")
    check("GET /students loads", r.status_code == 200)

    r = client.post("/students", data={"action": "add", "name": "Test Student", "klass": "8", "section": "C", "roll": "9"}, follow_redirects=True)
    students = appmod.get_master_students()
    check("student added", any(s["name"] == "Test Student" for s in students), str(students))

    new_student = next(s for s in students if s["name"] == "Test Student")
    ids = [s["id"] for s in students]
    names = [("Renamed Student" if s["id"] == new_student["id"] else s["name"]) for s in students]
    klasses = [s["class"] for s in students]
    sections = [s["section"] for s in students]
    rolls = [s["roll"] for s in students]
    r = client.post("/students", data={
        "action": "edit_all", "student_id": ids, "name": names,
        "klass": klasses, "section": sections, "roll": rolls,
    }, follow_redirects=True)
    students2 = appmod.get_master_students()
    check("edit_all renamed the right student", any(s["name"] == "Renamed Student" for s in students2), str(students2))
    check("edit_all did not affect other students", sum(1 for s in students2 if s["name"] == "Rohan Gupta") == 1)

    # 11. Autocomplete reflects newly added/renamed student
    r = client.get("/api/students?q=Renamed")
    check("new student searchable via autocomplete", len(r.get_json()) == 1)

    # 12. Workbook integrity after all writes: still loadable, named range + validations intact
    from openpyxl import load_workbook
    wb = load_workbook(TEST_DB)
    check("workbook still loads after all writes", True)
    check("StudentNames named range survived", "StudentNames" in wb.defined_names)
    check("Dashboard sheet present", "Dashboard" in wb.sheetnames)
    for key in appmod.CATEGORIES:
        ws = wb[key]
        check(f"{key} sheet has data validations", len(ws.data_validations.dataValidation) >= 1)

    # 13. Gemini response parsing (pure function, no network needed) -- shape
    # taken from Google's documented /v1beta/interactions example response.
    sample_response = {
        "created": "2025-11-26T12:25:15Z",
        "model": "gemini-3.8-flash",
        "status": "completed",
        "steps": [
            {
                "type": "model_output",
                "content": [{"type": "text", "text": "Diya Patel\nKabir Singh"}],
            }
        ],
    }
    check("parse_gemini_response extracts text", appmod.parse_gemini_response(sample_response) == "Diya Patel\nKabir Singh")
    check("parse_gemini_response handles empty steps", appmod.parse_gemini_response({"steps": []}) == "")
    check("parse_gemini_response handles missing steps key", appmod.parse_gemini_response({}) == "")
    non_output_step = {"steps": [{"type": "tool_call", "content": [{"type": "text", "text": "ignored"}]}]}
    check("parse_gemini_response ignores non-model_output steps", appmod.parse_gemini_response(non_output_step) == "")

    # 14. Repeat-offender aggregation math sanity check
    # 4 manual entries (one per category) + 1 checked row from the scan commit = 5 new rows.
    summary, repeat_list, max_month, recent = appmod.dashboard_data()
    total_entries = sum(s["total"] for s in summary.values())
    check("dashboard total matches rows written", total_entries == initial_total + 5, f"initial={initial_total} total={total_entries}")


try:
    run()
except Exception:
    print("\n--- UNEXPECTED EXCEPTION DURING TESTS ---")
    traceback.print_exc()
    results.append(("unexpected exception", False, "see traceback above"))
finally:
    shutil.rmtree(TMP_DIR, ignore_errors=True)

passed = sum(1 for _, ok, _ in results if ok)
failed = [name for name, ok, _ in results if not ok]

print("\n" + "=" * 60)
print(f"{passed}/{len(results)} checks passed")
if failed:
    print("FAILED:")
    for name in failed:
        print(f"  - {name}")
    sys.exit(1)
else:
    print("ALL TESTS PASSED")
    sys.exit(0)
