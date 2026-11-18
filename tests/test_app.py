"""
Automated smoke tests for the Discipline Tracker app.

Runs against a throwaway workbook freshly built from build_excel_template.py
(with its 5 example students), so it never reads or touches the real
data/SchoolDisciplineSystem.xlsx. Uses Flask's test client, so no server
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
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

import build_excel_template as bet  # noqa: E402

TMP_DIR = tempfile.mkdtemp(prefix="discipline_tracker_test_")
TEST_DB = os.path.join(TMP_DIR, "TestDatabase.xlsx")
bet.build_workbook().save(TEST_DB)

import app as appmod  # noqa: E402

appmod.DB_PATH = TEST_DB
appmod.UPLOAD_DIR = TMP_DIR
appmod.SETTINGS_PATH = os.path.join(TMP_DIR, "settings.json")
appmod.DEFAULT_SCAN_FOLDER = os.path.join(TMP_DIR, "scan-folder")
appmod.SCAN_JOBS_DIR = os.path.join(TMP_DIR, "scan-jobs")
appmod.app.config["TESTING"] = True

# Force the deterministic offline OCR path regardless of whatever real Gemini
# key is sitting in .env on this machine -- an automated smoke suite should
# never depend on a live third-party API (slow, and flaky by nature, as
# Gemini's own "high demand" 500s demonstrate). Gemini-specific logic
# (parse_gemini_response, the run_ocr fallback) is still covered separately
# below without making a real network call.
appmod.USING_GEMINI = False

client = appmod.app.test_client()

results = []


def check(name, condition, detail=""):
    results.append((name, bool(condition), detail))
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {name}" + (f" -- {detail}" if detail and not condition else ""))


def run():
    initial_summary, _ = appmod.dashboard_data()
    initial_total = sum(s["year"] for s in initial_summary.values())

    # 1. Dashboard loads
    r = client.get("/")
    check("dashboard loads (200)", r.status_code == 200)
    check("dashboard shows title", b"MCTM Dashboard" in r.data)

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

    # 4. Entry pages load for every category; the unified scan page loads
    # (optionally hinting a category, which it should accept without error).
    for key in appmod.CATEGORIES:
        r = client.get(f"/entry/{key}")
        check(f"GET /entry/{key} loads", r.status_code == 200)

    r = client.get("/scan")
    check("GET /scan loads", r.status_code == 200)
    r = client.get("/scan?hint=Uniform")
    check("GET /scan?hint=Uniform loads", r.status_code == 200)
    r = client.get("/scan?hint=NotARealCategory")
    check("GET /scan with a bogus hint doesn't error", r.status_code == 200)

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

    # 7. Dashboard reflects the new entries (the simplified dashboard no longer
    # lists individual names, so check the category's month count instead).
    r = client.get("/")
    expected_summary, _ = appmod.dashboard_data()
    check(
        "dashboard reflects entries",
        str(expected_summary["LateComers"]["month"]).encode() in r.data,
        f"expected LateComers month count {expected_summary['LateComers']['month']} to appear on page",
    )

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
        "/scan",
        data={"photo": (buf, "test_register.png"), "hint": "Uniform"},
        content_type="multipart/form-data",
    )
    check("POST /scan with image returns 200", r.status_code == 200)
    check("scan review falls back to the hint when the title can't be detected", b'value="Uniform"' in r.data, r.data[:2000])

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

    # 8b. Scan commit with no category chosen should be rejected, not silently
    # filed somewhere -- this is the safety net for the "which register is
    # this" step that replaced always-know-the-category-first navigation.
    before_commit_without_category = len(appmod.read_log_rows("Uniform", limit=10000))
    r = client.post(
        "/scan/commit",
        data={"category": "", "row_name": ["Aarav Sharma"], "row_include": ["0"]},
        follow_redirects=True,
    )
    check("scan commit without a category doesn't crash", r.status_code == 200)
    check(
        "scan commit without a category doesn't write anything",
        len(appmod.read_log_rows("Uniform", limit=10000)) == before_commit_without_category,
    )

    # 9. Scan commit path (independent of what OCR actually produced -- simulates the
    #    human review/correction step, which is the important safety net).
    # Uses a before/after delta -- the copied fixture may already contain real
    # OCR entries from live use of the app, not just what this test writes.
    before_scan_commit = len(appmod.read_log_rows("Uniform", limit=10000))
    r = client.post(
        "/scan/commit",
        data={
            "category": "Uniform",
            "row_name": ["Rohan Gupta", "Diya Patel"],
            "row_house": ["Wrong Guess", ""],  # OCR's raw guess -- matched student's real House should win
            "row_class": ["10", ""],
            "row_include": ["0"],  # only the first row is checked
            "row_remarks": ["", ""],
            "logged_by": "Scan-Test",
        },
        follow_redirects=True,
    )
    check("POST /scan/commit succeeds", r.status_code == 200)
    uniform_rows = appmod.read_log_rows("Uniform", limit=10000)
    check("scan commit saved exactly the checked row", len(uniform_rows) == before_scan_commit + 1, f"before={before_scan_commit} rows={uniform_rows}")
    ocr_rows = [row for row in uniform_rows if row.get("Source") == "OCR" and row.get("LoggedBy") == "Scan-Test"]
    check("scan commit row has Source=OCR", len(ocr_rows) == 1 and ocr_rows[0]["Name"] == "Rohan Gupta", str(ocr_rows))
    check(
        "scan commit uses the matched student's real House, not OCR's raw guess",
        ocr_rows and ocr_rows[0]["House"] == "Kailash",
        str(ocr_rows),
    )

    # 10. Master student list: add + edit_all
    r = client.get("/students")
    check("GET /students loads", r.status_code == 200)

    r = client.post("/students", data={"action": "add", "name": "Test Student", "klass": "8", "section": "C", "roll": "9", "house": "Purple"}, follow_redirects=True)
    students = appmod.get_master_students()
    check("student added", any(s["name"] == "Test Student" for s in students), str(students))
    check("student added with House", next(s for s in students if s["name"] == "Test Student")["house"] == "Purple")

    new_student = next(s for s in students if s["name"] == "Test Student")
    ids = [s["id"] for s in students]
    names = [("Renamed Student" if s["id"] == new_student["id"] else s["name"]) for s in students]
    klasses = [s["class"] for s in students]
    sections = [s["section"] for s in students]
    rolls = [s["roll"] for s in students]
    houses = [("Orange" if s["id"] == new_student["id"] else s["house"]) for s in students]
    r = client.post("/students", data={
        "action": "edit_all", "student_id": ids, "name": names,
        "klass": klasses, "section": sections, "roll": rolls, "house": houses,
    }, follow_redirects=True)
    students2 = appmod.get_master_students()
    check("edit_all renamed the right student", any(s["name"] == "Renamed Student" for s in students2), str(students2))
    check("edit_all did not affect other students", sum(1 for s in students2 if s["name"] == "Rohan Gupta") == 1)
    check("edit_all updated House too", next(s for s in students2 if s["name"] == "Renamed Student")["house"] == "Orange")

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

    # 13b. run_ocr fallback behaviour (Gemini down -> Tesseract picks up the
    # slack automatically) -- monkeypatched so it's deterministic and makes no
    # network call, exercising exactly the bug fixed after Gemini returned a
    # real "high demand" 500 in production.
    orig_gemini, orig_tesseract = appmod.run_ocr_gemini, appmod.run_ocr_tesseract
    orig_using_gemini, orig_ocr_available = appmod.USING_GEMINI, appmod.OCR_AVAILABLE
    try:
        appmod.USING_GEMINI = True
        appmod.OCR_AVAILABLE = True

        appmod.run_ocr_gemini = lambda path: ("Gemini text", None)
        text, error, warning, backend = appmod.run_ocr("fake-path")
        check("run_ocr uses Gemini text when it succeeds", text == "Gemini text" and error is None and warning is None)
        check("run_ocr reports backend=gemini on success", backend == "gemini")

        appmod.run_ocr_gemini = lambda path: ("", "high demand, try again later")
        appmod.run_ocr_tesseract = lambda path: ("Tesseract fallback text", None)
        text, error, warning, backend = appmod.run_ocr("fake-path")
        check("run_ocr falls back to Tesseract when Gemini fails", text == "Tesseract fallback text" and error is None)
        check("run_ocr surfaces a warning (not a hard error) on fallback", warning is not None and "Tesseract" in warning)
        check("run_ocr reports backend=tesseract after fallback", backend == "tesseract")

        appmod.run_ocr_gemini = lambda path: ("", "high demand, try again later")
        appmod.run_ocr_tesseract = lambda path: ("", "tesseract also broken")
        text, error, warning, backend = appmod.run_ocr("fake-path")
        check("run_ocr reports a hard error when both backends fail", error is not None and "also failed" in error)
    finally:
        appmod.run_ocr_gemini = orig_gemini
        appmod.run_ocr_tesseract = orig_tesseract
        appmod.USING_GEMINI = orig_using_gemini
        appmod.OCR_AVAILABLE = orig_ocr_available

    # 13b-2. Gemini model fallback chain (model A out of quota -> model B ->
    # ...) -- _gemini_request is monkeypatched, so no network call is made.
    from PIL import Image as _Image
    chain_img = os.path.join(TMP_DIR, "chain_test.png")
    _Image.new("RGB", (60, 40), "white").save(chain_img)
    orig_request, orig_models = appmod._gemini_request, appmod.GEMINI_MODELS
    try:
        appmod.GEMINI_MODELS = ["model-a", "model-b", "model-c"]
        calls = []

        def fake_request(responses):
            def _fake(model, b64_image, timeout):
                calls.append(model)
                return responses[model]
            return _fake

        def reset():
            calls.clear()
            appmod._gemini_cooldown_until.clear()

        reset()
        appmod._gemini_request = fake_request({
            "model-a": ("", "HTTP 429: quota used up", "quota"),
            "model-b": ("text from b", None, None),
            "model-c": ("text from c", None, None),
        })
        text, error = appmod.run_ocr_gemini(chain_img)
        check("model chain moves to the next model on a quota error", text == "text from b" and error is None, f"{text!r} {error!r}")
        check("model chain stops at the first model that works", calls == ["model-a", "model-b"], str(calls))

        calls.clear()
        text, error = appmod.run_ocr_gemini(chain_img)
        check("a model that just failed is skipped on the next scan", calls == ["model-b"], str(calls))

        reset()
        appmod._gemini_request = fake_request({
            "model-a": ("", "HTTP 404: retired", "gone"),
            "model-b": ("", "HTTP 503: overloaded", "busy"),
            "model-c": ("", "no answer within 45s", "busy"),
        })
        text, error = appmod.run_ocr_gemini(chain_img)
        check("model chain tries every model before giving up", calls == ["model-a", "model-b", "model-c"], str(calls))
        check("model chain reports an error naming each failed model", error is not None and all(m in error for m in appmod.GEMINI_MODELS), str(error))

        calls.clear()
        appmod._gemini_request = fake_request({
            "model-a": ("back again", None, None),
            "model-b": ("text from b", None, None),
            "model-c": ("text from c", None, None),
        })
        text, error = appmod.run_ocr_gemini(chain_img)
        check("when every model is cooling down they are all retried anyway", text == "back again" and calls == ["model-a"], f"{text!r} {calls}")

        reset()
        appmod._gemini_request = fake_request({
            "model-a": ("", "HTTP 403: API key not valid", "auth"),
            "model-b": ("text from b", None, None),
            "model-c": ("text from c", None, None),
        })
        text, error = appmod.run_ocr_gemini(chain_img)
        check("a rejected API key stops the chain (other models can't help)", calls == ["model-a"] and error is not None and "API key" in error, f"{calls} {error!r}")

        reset()
        appmod._gemini_request = fake_request({
            "model-a": ("", "could not reach Gemini", "network"),
            "model-b": ("text from b", None, None),
            "model-c": ("text from c", None, None),
        })
        text, error = appmod.run_ocr_gemini(chain_img)
        check("no internet stops the chain (other models can't help)", calls == ["model-a"] and error is not None, f"{calls} {error!r}")
        check("a network failure doesn't put the model on cooldown", "model-a" not in appmod._gemini_cooldown_until)
    finally:
        appmod._gemini_request = orig_request
        appmod.GEMINI_MODELS = orig_models
        appmod._gemini_cooldown_until.clear()

    check("gemini-4-argon is not in the default free-tier chain", "gemini-4-argon" not in appmod.GEMINI_FALLBACK_MODELS)

    # 13c. Structured House/Name/Class extraction (the Gemini path) -- this is
    # the actual feature requested: pull 3 fields per row, and never treat the
    # register's own column-heading row as if it were a student.
    students_for_ocr = appmod.get_master_students()

    gemini_json_text = (
        '[{"house": "House", "name": "Name", "class": "Class"},'
        ' {"house": "Red", "name": "Rohan Gupta", "class": "10"},'
        ' {"house": "Blue", "name": "Totally Unknown Kid", "class": "7"}]'
    )
    structured = appmod.parse_structured_ocr_json(gemini_json_text)
    check("parse_structured_ocr_json parses a plain JSON array", structured is not None and len(structured) == 3)

    candidates = appmod.build_candidates_from_structured(structured, students_for_ocr)
    check(
        "build_candidates_from_structured drops the header row even if the model included it",
        all(c["guess_name"] != "Name" for c in candidates),
        str(candidates),
    )
    check("build_candidates_from_structured keeps the 2 real rows", len(candidates) == 2, str(candidates))
    rohan_candidate = next((c for c in candidates if "Rohan" in c["guess_name"]), None)
    check("build_candidates_from_structured matches a known student by name", rohan_candidate is not None and rohan_candidate["confident"])
    rohan_master = appmod.find_student_by_name("Rohan Gupta", students_for_ocr)
    check(
        "a matched student's House comes from the master list, not OCR",
        rohan_candidate is not None and rohan_candidate["guess_house"] == rohan_master["house"],
        str(rohan_candidate),
    )
    check(
        "a matched student's Class comes from the master list, not OCR",
        rohan_candidate is not None and rohan_candidate["guess_class"] == appmod.class_label(rohan_master),
        str(rohan_candidate),
    )
    check("a matched student carries their student ID", rohan_candidate is not None and rohan_candidate["student_id"] == rohan_master["id"])
    unknown_candidate = next((c for c in candidates if "Unknown" in c["guess_name"]), None)
    check("unmatched name still comes through unconfident (for manual review)", unknown_candidate is not None and not unknown_candidate["confident"])
    check(
        "an unmatched student keeps what OCR read for House/Class, with no ID",
        unknown_candidate is not None and unknown_candidate["guess_house"] == "Blue"
        and unknown_candidate["guess_class"] == "7" and unknown_candidate["student_id"] == "",
        str(unknown_candidate),
    )

    # Gemini wrapping its JSON in a markdown code fence (common LLM habit) should still parse.
    fenced = "```json\n" + gemini_json_text + "\n```"
    check("parse_structured_ocr_json strips markdown code fences", appmod.parse_structured_ocr_json(fenced) is not None)

    # Plain prose (Gemini ignored the JSON instruction) should signal "fall back", not crash.
    check("parse_structured_ocr_json returns None for non-JSON text", appmod.parse_structured_ocr_json("Rohan Gupta\nDiya Patel") is None)

    # 13d. Heuristic column split + header filtering (the Tesseract plain-text path)
    check("is_header_like flags a real header row", appmod.is_header_like(["House", "Name", "Class"]))
    check("is_header_like does not flag a real student row", not appmod.is_header_like(["Red", "Rohan Gupta", "10"]))

    tesseract_lines = appmod.parse_ocr_lines(
        "House    Name    Class\nRed    Xyzzy Qwerty    10\nBlue    Diya Patel    9",
        students_for_ocr,
    )
    check("parse_ocr_lines (Tesseract path) also drops the header line", len(tesseract_lines) == 2, str(tesseract_lines))
    check(
        "parse_ocr_lines splits House/Name/Class from spaced columns",
        tesseract_lines[0]["guess_house"] == "Red" and tesseract_lines[0]["guess_class"] == "10",
        str(tesseract_lines),
    )
    diya_master = appmod.find_student_by_name("Diya Patel", students_for_ocr)
    check(
        "parse_ocr_lines fills a matched student's House/Class from the master list",
        tesseract_lines[1]["guess_house"] == diya_master["house"] and tesseract_lines[1]["guess_class"] == appmod.class_label(diya_master),
        str(tesseract_lines),
    )

    # 13e. extract_candidates dispatches correctly by backend, and returns
    # (candidates, category_guess) now that scanning is a single unified flow.
    cands, cat_guess = appmod.extract_candidates(gemini_json_text, students_for_ocr, "gemini")
    check("extract_candidates uses structured parsing for backend='gemini'", len(cands) == 2)
    check("extract_candidates has no category guess for a bare array (no category_guess key)", cat_guess is None)

    cands, cat_guess = appmod.extract_candidates("Rohan Gupta", students_for_ocr, "gemini")
    check("extract_candidates falls back to line parsing if Gemini text wasn't JSON", len(cands) == 1)

    cands, cat_guess = appmod.extract_candidates("Rohan Gupta", students_for_ocr, "tesseract")
    check("extract_candidates uses line parsing for backend='tesseract'", len(cands) == 1)

    cands, cat_guess = appmod.extract_candidates("", students_for_ocr, "gemini")
    check("extract_candidates returns nothing for empty text", cands == [] and cat_guess is None)

    # 13f. The actual feature: detecting which register a scan belongs to
    # from its title, so the user doesn't have to navigate there first.
    check("match_category_label matches an exact label", appmod.match_category_label("Uniform") == "Uniform")
    check("match_category_label matches a category key directly", appmod.match_category_label("NailsHair") == "NailsHair")
    check("match_category_label is case-insensitive", appmod.match_category_label("late comers") == "LateComers")
    check("match_category_label is tolerant of near misses", appmod.match_category_label("Defaultrs") == "Defaulters")
    check("match_category_label returns None for unrelated text", appmod.match_category_label("Rohan Gupta") is None)
    check("match_category_label returns None for empty text", appmod.match_category_label("") is None)

    object_shape_text = (
        '{"category_guess": "Uniform", "rows": ['
        '{"house": "House", "name": "Name", "class": "Class"},'
        '{"house": "Red", "name": "Rohan Gupta", "class": "10"}]}'
    )
    cat_text, rows = appmod.parse_gemini_extraction(object_shape_text)
    check("parse_gemini_extraction reads category_guess from the object shape", cat_text == "Uniform")
    check("parse_gemini_extraction still extracts rows from the object shape", rows is not None and len(rows) == 2)

    cands, cat_guess = appmod.extract_candidates(object_shape_text, students_for_ocr, "gemini")
    check("extract_candidates surfaces the detected category from Gemini's guess", cat_guess == "Uniform")
    check("extract_candidates still drops the header row in the object shape", len(cands) == 1, str(cands))

    check(
        "guess_category_from_text_lines finds a title among the first lines (Tesseract path)",
        appmod.guess_category_from_text_lines("UNIFORM REGISTER\nRed   Rohan Gupta   10") == "Uniform",
    )
    check(
        "guess_category_from_text_lines returns None when no title is present",
        appmod.guess_category_from_text_lines("Red   Rohan Gupta   10\nBlue   Diya Patel   9") is None,
    )

    # 14. dashboard_data() shape: each category has month/year/months (12 entries)
    summary, stats = appmod.dashboard_data()
    total_year_entries = sum(s["year"] for s in summary.values())
    check("dashboard year total matches rows written", total_year_entries == initial_total + 5, f"initial={initial_total} total={total_year_entries}")
    check("each category summary has month/year/months", all("month" in s and "year" in s and "months" in s for s in summary.values()))
    check(
        "each category's months breakdown covers Jan-Dec",
        all([m["label"] for m in s["months"]] == appmod.MONTH_NAMES for s in summary.values()),
    )
    check(
        "a category's months breakdown sums to its year total",
        all(sum(m["count"] for m in s["months"]) == s["year"] for s in summary.values()),
    )
    check("stats has max_month/max_year_month", stats["max_month"] >= 0 and stats["max_year_month"] >= 0)

    # 15. House breakdown is well-formed
    check("stats house_breakdown is a list of label/count dicts", all("label" in h and "count" in h for h in stats["house_breakdown"]))

    # 16. Weekly repeat-flag logic (3+ times in the same ISO week) -- built on
    # synthetic data anchored via date.fromisocalendar so it's deterministic
    # regardless of what today's actual date happens to be.
    week_monday = date.fromisocalendar(2026, 10, 1)
    d1, d2, d3 = week_monday, week_monday + timedelta(days=1), week_monday + timedelta(days=2)
    synthetic_rows = {
        "LateComers": [{"Name": "Test Kid", "Class": "9", "Date": d1}, {"Name": "Test Kid", "Class": "9", "Date": d2}],
        "Defaulters": [{"Name": "Test Kid", "Class": "9", "Date": d3}],
        "Uniform": [{"Name": "Only Twice", "Class": "9", "Date": d1}, {"Name": "Only Twice", "Class": "9", "Date": d2}],
        "NailsHair": [],
    }
    flags = appmod.compute_weekly_repeat_flags(synthetic_rows, threshold=3)
    check("compute_weekly_repeat_flags flags a student with 3 entries in one week", any(f["name"] == "Test Kid" and f["count"] == 3 for f in flags))
    check("compute_weekly_repeat_flags does not flag a student with only 2 entries in one week", not any(f["name"] == "Only Twice" for f in flags))
    test_kid_flag = next((f for f in flags if f["name"] == "Test Kid"), None)
    check(
        "weekly flag records per-category breakdown",
        test_kid_flag is not None and test_kid_flag["categories"]["LateComers"] == 2 and test_kid_flag["categories"]["Defaulters"] == 1,
        str(test_kid_flag),
    )
    check("weekly flag has a human week_label", test_kid_flag is not None and test_kid_flag["week_label"].startswith("Week of"))

    chart = appmod.weekly_flag_chart_series(flags, num_weeks=10)
    check("weekly_flag_chart_series returns num_weeks entries", len(chart) == 10)
    check("weekly_flag_chart_series entries have label/count", all("label" in c and "count" in c for c in chart))

    # 17. Entry date input
    today = date.today()
    d, err = appmod.parse_entry_date("")
    check("blank entry date means today", d == today and err is None)
    d, err = appmod.parse_entry_date((today - timedelta(days=40)).isoformat())
    check("a past entry date is accepted", d == today - timedelta(days=40) and err is None)
    d, err = appmod.parse_entry_date("not-a-date")
    check("a malformed entry date is rejected", d is None and err)
    d, err = appmod.parse_entry_date((today + timedelta(days=1)).isoformat())
    check("a future entry date is rejected", d is None and err and "future" in err)

    # 18. Finding a student by ID or name
    master = appmod.get_master_students()
    kabir = appmod.find_student_by_name("Kabir Singh", master)
    meera = appmod.find_student_by_name("Meera Nair", master)
    check("resolve_student finds a student by ID (any case)", appmod.resolve_student("", kabir["id"].lower(), master) == kabir)
    check("resolve_student finds a student by an ID typed in the name box", appmod.resolve_student(meera["id"], "", master) == meera)
    check("resolve_student finds a student by exact name", appmod.resolve_student("kabir singh", "", master) == kabir)
    check("resolve_student returns None for an unknown student", appmod.resolve_student("Nobody Atall", "STU999", master) is None)
    check("lookup API returns the student for an ID", client.get(f"/api/students/lookup?q={kabir['id']}").get_json()["name"] == "Kabir Singh")
    check("lookup API returns null for an unknown student", client.get("/api/students/lookup?q=zzz").get_json() is None)
    check("search API also matches on student ID", any(s["id"] == kabir["id"] for s in client.get(f"/api/students?q={kabir['id']}").get_json()))

    # 19. Manual entry: backdated, by ID only, and ID typed in the name box
    backdate = today - timedelta(days=40)
    before_ids = {row["EntryID"] for row in appmod.read_log_rows("Defaulters", limit=100000)}
    r = client.post("/entry/Defaulters", data={
        "date": backdate.isoformat(), "name": "", "student_id": kabir["id"],
        "Reason": "No homework", "remarks": "", "logged_by": "Test",
    }, follow_redirects=True)
    new_rows = [row for row in appmod.read_log_rows("Defaulters", limit=100000) if row["EntryID"] not in before_ids]
    check("entry by student ID only is saved", r.status_code == 200 and len(new_rows) == 1, str(new_rows))
    by_id = new_rows[0] if new_rows else {}
    check("entry by ID fills name/class/house from the master list",
          by_id.get("Name") == "Kabir Singh" and str(by_id.get("Class")) == str(kabir["class"]) and by_id.get("House") == kabir["house"], str(by_id))
    check("entry keeps the date that was typed in", appmod._row_date(by_id) == backdate, str(by_id.get("Date")))

    before_ids = {row["EntryID"] for row in appmod.read_log_rows("Defaulters", limit=100000)}
    client.post("/entry/Defaulters", data={"name": meera["id"], "Reason": "Late work"}, follow_redirects=True)
    new_rows = [row for row in appmod.read_log_rows("Defaulters", limit=100000) if row["EntryID"] not in before_ids]
    check("an ID typed in the name box saves under the student's real name",
          len(new_rows) == 1 and new_rows[0]["Name"] == "Meera Nair" and new_rows[0]["House"] == meera["house"], str(new_rows))

    count_before = len(appmod.read_log_rows("Defaulters", limit=100000))
    r = client.post("/entry/Defaulters", data={"date": (today + timedelta(days=3)).isoformat(), "name": "Kabir Singh"})
    check("a future-dated entry is refused", r.status_code == 400 and len(appmod.read_log_rows("Defaulters", limit=100000)) == count_before)
    r = client.post("/entry/Defaulters", data={"name": "", "student_id": "STU999"})
    check("an unknown student ID alone is refused", r.status_code == 400 and len(appmod.read_log_rows("Defaulters", limit=100000)) == count_before)

    # 20. Edit an entry
    entry_id = by_id.get("EntryID")
    r = client.get(f"/entry/Defaulters/{entry_id}/edit")
    check("edit page loads", r.status_code == 200 and entry_id.encode() in r.data)
    new_date = today - timedelta(days=2)
    r = client.post(f"/entry/Defaulters/{entry_id}/edit", data={
        "date": new_date.isoformat(), "name": "Kabir Singh", "student_id": kabir["id"],
        "Reason": "Incomplete project", "remarks": "Parent informed", "logged_by": "Editor",
    }, follow_redirects=True)
    edited = appmod.get_log_entry("Defaulters", entry_id)
    check("editing changes the saved values",
          r.status_code == 200 and edited["Reason"] == "Incomplete project" and edited["Remarks"] == "Parent informed"
          and appmod._row_date(edited) == new_date and edited["LoggedBy"] == "Editor", str(edited))
    check("editing keeps the entry ID and source", edited["EntryID"] == entry_id and edited["Source"] == "Manual")
    r = client.get("/entry/Defaulters/DE-99999/edit", follow_redirects=True)
    check("editing a missing entry says so", r.status_code == 200 and b"doesn&#39;t exist" in r.data)

    # 21. Delete an entry
    count_before = len(appmod.read_log_rows("Defaulters", limit=100000))
    r = client.post(f"/entry/Defaulters/{entry_id}/delete", follow_redirects=True)
    check("deleting removes the entry", r.status_code == 200 and appmod.get_log_entry("Defaulters", entry_id) is None
          and len(appmod.read_log_rows("Defaulters", limit=100000)) == count_before - 1)
    from openpyxl import load_workbook as _load
    ws = _load(TEST_DB)["Defaulters"]
    blank_rows = [r_ for r_ in range(2, ws.max_row + 1) if not ws.cell(row=r_, column=1).value]
    first_blank = blank_rows[0] if blank_rows else None
    check("a deleted row gets its auto-fill formulas back",
          first_blank is not None and str(ws.cell(row=first_blank, column=4).value).startswith("=IFERROR(VLOOKUP(C"))
    r = client.post(f"/entry/Defaulters/{entry_id}/delete", follow_redirects=True)
    check("deleting it twice is handled", r.status_code == 200 and b"already gone" in r.data)

    # 22. Open Excel button
    opened = []
    orig_open = appmod._open_file_on_server
    try:
        appmod._open_file_on_server = lambda path: opened.append(path)
        r = client.get("/open-excel")
        check("on the server PC, Open Excel opens the live file", r.status_code == 302 and opened == [appmod.DB_PATH], str(opened))
        r = client.get("/open-excel", environ_base={"REMOTE_ADDR": "192.168.1.77"})
        check("from another PC, Open Excel downloads a copy",
              r.status_code == 200 and "attachment" in r.headers.get("Content-Disposition", "") and len(opened) == 1)
        r = client.get("/open-excel?download=1")
        check("?download=1 always downloads", r.status_code == 200 and "attachment" in r.headers.get("Content-Disposition", ""))
    finally:
        appmod._open_file_on_server = orig_open

    # 23. Saving while Excel has the file open shows a message instead of crashing
    orig_save = appmod.save_wb
    try:
        def locked(_wb):
            raise appmod.WorkbookLocked()
        appmod.save_wb = locked
        r = client.post("/entry/Defaulters", data={"name": "Kabir Singh"}, headers={"Referer": "/entry/Defaulters"}, follow_redirects=True)
        check("a locked Excel file gives a clear message", r.status_code == 200 and b"Excel file is open" in r.data)
    finally:
        appmod.save_wb = orig_save

    # 24. Dashboard patterns (synthetic data with a fixed "today")
    pt_today = date(2026, 3, 18)  # a Wednesday
    def row(name, d, klass="9", section="A", house="Kailash", **extra):
        base = {"Name": name, "Date": d, "Class": klass, "Section": section, "House": house}
        base.update(extra)
        return base
    synthetic = {
        "LateComers": [
            row("A Kid", date(2026, 3, 16), Time="8:40 AM", Reason="Bus late"),
            row("A Kid", date(2026, 3, 17), Time="8:35", Reason="bus late"),
            row("A Kid", date(2026, 3, 18), Time="8:44 am", Reason="Overslept"),
            row("B Kid", date(2026, 3, 9), klass="10", section="B", house="Aravali", Time="8:31", Reason="Bus late"),
            row("C Kid", date(2026, 3, 2), klass="10", section="B", house="Aravali", Time="9:10", Reason=""),
            row("D Kid", date(2026, 2, 9), Time="8:40", Reason=""),
        ],
        "Defaulters": [],
        "Uniform": [row("A Kid", date(2026, 3, 10), Issue="No tie")],
        "NailsHair": [],
    }
    pt_flags = appmod.compute_weekly_repeat_flags(synthetic, threshold=3)
    patterns = appmod.build_patterns(synthetic, pt_flags, pt_today)
    texts = [p["text"] for p in patterns]
    joined = " | ".join(texts)
    check("patterns lead with students already at 3+ this week",
          patterns and patterns[0]["tone"] == "alert" and "A Kid" in patterns[0]["text"] and "this week" in patterns[0]["text"], joined)
    check("patterns compare this month with last month",
          any(p["tone"] == "up" and "Late Comers" in p["text"] and "400%" in p["text"] for p in patterns), joined)
    check("patterns spot the peak weekday", "Mondays" in joined, joined)
    check("patterns spot the usual late-arrival time", "8:30 and 8:45" in joined, joined)
    check("patterns spot the most common reason", "“Bus late” (3 times)" in joined, joined)
    check("patterns spot students in more than one register", "more than one register" in joined and "A Kid (Late Comers, Uniform)" in joined, joined)
    check("patterns are capped", len(patterns) <= appmod.MAX_PATTERNS)
    check("no data means no patterns", appmod.build_patterns({k: [] for k in appmod.CATEGORIES}, [], pt_today) == [])
    r = client.get("/")
    check("dashboard shows the Patterns panel", r.status_code == 200 and b"Patterns" in r.data)

    run_find_tests()
    run_roster_tests()
    run_multi_scan_tests()
    run_folder_scan_tests()


def run_folder_scan_tests():
    import json as _json
    import re as _re
    import time as _time
    from PIL import Image as _Im

    real_today = date.today()
    folder = os.path.join(TMP_DIR, "my register photos")
    os.makedirs(os.path.join(folder, "Scanned"), exist_ok=True)

    def make_photo(path):
        _Im.new("RGB", (40, 50), "white").save(path)

    # The default folder is created on first use and is empty.
    r = client.get("/scan")
    check("folder: scan page shows the folder panel", b"Scan a whole folder" in r.data and b"No photos in it yet" in r.data)
    check("folder: default folder created", os.path.isdir(appmod.DEFAULT_SCAN_FOLDER))
    r = client.post("/scan/folder/start", follow_redirects=True)
    check("folder: empty folder -> clear message, no job", b"There are no photos in" in r.data and not appmod.SCAN_JOBS)

    # Choose another folder.
    r = client.post("/scan/folder/settings", data={"folder": os.path.join(TMP_DIR, "nope")}, follow_redirects=True)
    check("folder: a missing folder is refused", b"no folder at" in r.data and appmod.scan_folder_path() != os.path.join(TMP_DIR, "nope"))
    r = client.post("/scan/folder/settings", data={"folder": f'"{folder}"'}, follow_redirects=True)
    check("folder: a real folder is remembered (quotes from Explorer's 'Copy as path' are fine)",
          appmod.scan_folder_path() == os.path.abspath(folder))

    for name in ["c_late.jpg", "a_late.png", "b_uniform.jpg", "d_blurry.jpg"]:
        make_photo(os.path.join(folder, name))
        _time.sleep(0.02)  # distinct modified times: oldest first
    open(os.path.join(folder, "notes.txt"), "w").write("not a photo")
    make_photo(os.path.join(folder, "Scanned", "old.jpg"))
    check("folder: only photos directly in the folder count",
          [os.path.basename(p) for p in appmod.list_folder_photos(folder)] == ["c_late.jpg", "a_late.png", "b_uniform.jpg", "d_blurry.jpg"])
    r = client.get("/scan")
    check("folder: scan page counts waiting photos", b"<b>4</b> photos waiting" in r.data and b"Scan 4 photos" in r.data)

    iso = lambda days: (real_today - timedelta(days=days)).isoformat()  # noqa: E731
    by_source = {
        "c_late": {"category_guess": "Late Comers", "page_date": iso(1),
                   "rows": [{"house": "", "name": "Rohan Gupta", "class": "10", "date": ""}]},
        "a_late": {"category_guess": "Late Comers", "page_date": iso(2),
                   "rows": [{"house": "", "name": "Diya Patel", "class": "9", "date": ""}]},
        "b_uniform": {"category_guess": "Uniform", "page_date": iso(3),
                      "rows": [{"house": "", "name": "Kabir Singh", "class": "9", "date": ""}]},
    }

    def fake_run_ocr(path):
        _time.sleep(0.05)
        key = _re.search(r"_\d\d_(.+)\.\w+$", os.path.basename(path)).group(1)
        page = by_source.get(key)
        return (_json.dumps(page), None, None, "gemini") if page else ("", "could not read the photo", None, None)

    real_run_ocr = appmod.run_ocr
    appmod.run_ocr = fake_run_ocr
    try:
        r = client.post("/scan/folder/start", data={"hint": ""})
        m = _re.search(r"/scan/job/([^/\"?]+)", r.headers.get("Location", ""))
        check("folder: starting a scan opens its progress page", r.status_code == 302 and m, r.headers.get("Location"))
        job_id = m.group(1) if m else ""
        r = client.get(f"/scan/job/{job_id}")
        check("folder: progress page while reading", r.status_code == 200 and (b"Reading 4 photos" in r.data or b"scan-group" in r.data))
        r = client.post("/scan/folder/start")
        check("folder: a second start joins the running batch",
              appmod.get_job(job_id)["status"] != "running" or job_id in r.headers.get("Location", ""))
        for _ in range(100):
            st = client.get(f"/scan/job/{job_id}/status").get_json()
            if st["status"] != "running":
                break
            _time.sleep(0.05)
        check("folder: status reports progress to the end", st == {"status": "done", "done": 4, "total": 4}, str(st))
        r = client.get(f"/scan/job/{job_id}")
        html = r.data.decode()
        check("folder: review shows every photo, oldest first",
              html.count('class="panel scan-group"') == 4 and html.index("c_late.jpg") < html.index("a_late.png") < html.index("b_uniform.jpg"))
        check("folder: review carries the batch id", f'name="job_id" value="{job_id}"' in html)
        check("folder: the folder itself is untouched until saving", len(appmod.list_folder_photos(folder)) == 4)
        # The server restarts mid-review: the batch comes back from disk.
        check("folder: a finished batch is saved to disk", os.path.exists(os.path.join(appmod.SCAN_JOBS_DIR, f"{job_id}.json")))
        with appmod.scan_jobs_lock:
            appmod.SCAN_JOBS.clear()
        r = client.get(f"/scan/job/{job_id}")
        reloaded = appmod.get_job(job_id)
        check("folder: after a restart the review still opens", r.status_code == 200
              and r.data.decode().count('class="panel scan-group"') == 4, r.data[:300])
        check("folder: reloaded batch keeps dates as dates",
              reloaded and isinstance(reloaded["groups"][0]["page_date"], date)
              and all(isinstance(c["row_date"], (date, type(None))) for g in reloaded["groups"] for c in g["candidates"]))
        check("folder: unknown batch id -> back to Scan with a message",
              b"no longer available" in client.get("/scan/job/nope", follow_redirects=True).data)
    finally:
        appmod.run_ocr = real_run_ocr

    # A bad save keeps the batch and returns to it.
    r = client.post("/scan/commit", data={"job_id": job_id, "group_category_0": "",
                                          "row_name": ["Rohan Gupta"], "row_group": ["0"], "row_include": ["0"]})
    check("folder: a bad save returns to the batch, not a rescan",
          r.status_code == 302 and f"/scan/job/{job_id}" in r.headers.get("Location", "") and appmod.get_job(job_id))

    # Save photos 0 and 2 (c_late, b_uniform); leave a_late unticked; d_blurry had nothing.
    before = len(appmod.read_log_rows("LateComers", limit=100000))
    r = client.post("/scan/commit", data={
        "job_id": job_id, "logged_by": "FolderTest",
        "group_category_0": "LateComers", "group_date_0": iso(1),
        "group_category_1": "LateComers", "group_date_1": iso(2),
        "group_category_2": "Uniform", "group_date_2": iso(3),
        "row_name": ["Rohan Gupta", "Diya Patel", "Kabir Singh"], "row_group": ["0", "1", "2"],
        "row_include": ["0", "2"],
    }, follow_redirects=True)
    check("folder: saved rows land in their registers",
          len(appmod.read_log_rows("LateComers", limit=100000)) == before + 1
          and any(x.get("LoggedBy") == "FolderTest" for x in appmod.read_log_rows("Uniform", limit=100000)))
    left = sorted(os.path.basename(p) for p in appmod.list_folder_photos(folder))
    check("folder: saved photos move to Scanned; the rest wait", left == ["a_late.png", "d_blurry.jpg"], str(left))
    check("folder: Scanned keeps old files and avoids name clashes",
          sorted(os.listdir(os.path.join(folder, "Scanned"))) == ["b_uniform.jpg", "c_late.jpg", "old.jpg"])
    check("folder: message says what moved and what's left", b"2 photos moved to the Scanned folder; 2 still waiting" in r.data)
    check("folder: a saved batch is closed, on disk too",
          appmod.get_job(job_id) is None and not os.path.exists(os.path.join(appmod.SCAN_JOBS_DIR, f"{job_id}.json")))
    check("batch ids can't reach outside the batches folder", appmod.get_job("../../app") is None and appmod._job_file("../x") is None)

    # More than 100 photos: the oldest 100 are read.
    for i in range(103):
        make_photo(os.path.join(folder, f"bulk_{i:03d}.jpg"))
    appmod.run_ocr = lambda path: ("", "skipped in test", None, None)
    try:
        r = client.post("/scan/folder/start", follow_redirects=False)
        job = appmod.get_job(_re.search(r"/scan/job/([^/\"?]+)", r.headers["Location"]).group(1))
        check("folder: at most 100 photos per batch", job["total"] == 100)
        for _ in range(200):
            if job["status"] != "running":
                break
            _time.sleep(0.05)
        check("folder: a 100-photo batch finishes", job["status"] == "done" and job["done"] == 100)
    finally:
        appmod.run_ocr = real_run_ocr

    client.post("/scan/folder/settings", data={"reset": "1"})
    check("folder: back to the default folder", appmod.scan_folder_path() == appmod.DEFAULT_SCAN_FOLDER)


def run_multi_scan_tests():
    import json as _json
    import re as _re
    T = date(2026, 11, 18)
    prd = lambda s: appmod.parse_register_date(s, today=T)  # noqa: E731

    # --- reading dates ---------------------------------------------------
    check("date: 18/11/2026", prd("18/11/2026") == date(2026, 11, 18))
    check("date: day first (05/11/2026 is 5 Nov)", prd("05/11/2026") == date(2026, 11, 5))
    check("date: 2-digit year 18-11-26", prd("18-11-26") == date(2026, 11, 18))
    check("date: dots 18.11.2026", prd("18.11.2026") == date(2026, 11, 18))
    check("date: ISO 2026-11-18", prd("2026-11-18") == date(2026, 11, 18))
    check("date: 18th November", prd("18th November") == date(2026, 11, 18))
    check("date: 18 Nov 2026", prd("Date: 18 Nov 2026") == date(2026, 11, 18))
    check("date: Nov 18, 2026", prd("Nov 18, 2026") == date(2026, 11, 18))
    check("date: no year 18/11", prd("18/11") == date(2026, 11, 18))
    check("date: no year in the future means last year", prd("20/12") == date(2025, 12, 20))
    check("date: a future date is a misread", prd("25/12/2026") is None)
    check("date: an arrival time is not a date", prd("8.45") is None and prd("8:45 AM") is None)
    check("date: nonsense is None", prd("Rohan Gupta") is None and prd("") is None and prd(None) is None)

    # --- Tesseract text: date headings and per-row dates -------------------
    students = appmod.get_master_students()
    real_today = date.today()
    d1 = (real_today - timedelta(days=3)).strftime("%d/%m/%Y")
    d2 = (real_today - timedelta(days=1)).strftime("%d/%m/%Y")
    lines = appmod.parse_ocr_lines(f"Date: {d1}\nKailash   Rohan Gupta   10\n{d2}\nAravali   Diya Patel   9", students)
    check("tesseract: date-only lines are not students", len(lines) == 2, str(lines))
    check("tesseract: rows take the date heading above them",
          lines and lines[0]["row_date"] == real_today - timedelta(days=3) and lines[1]["row_date"] == real_today - timedelta(days=1), str(lines))
    check("tesseract: page date from the top of the page",
          appmod.page_date_from_text(f"UNIFORM REGISTER\nDate: {d1}\nRohan Gupta") == real_today - timedelta(days=3))
    lines = appmod.parse_ocr_lines(f"Kailash   Rohan Gupta   10   {d2}", students)
    check("tesseract: a date in the row is that row's date", lines and lines[0]["row_date"] == real_today - timedelta(days=1)
          and lines[0]["guess_name"] == "Rohan Gupta", str(lines))

    # --- several photos at once, with a fake OCR engine --------------------
    iso = lambda days: (real_today - timedelta(days=days)).isoformat()  # noqa: E731
    pages = {
        "01": {"category_guess": "Late Comers", "page_date": iso(2),
               "rows": [{"house": "", "name": "Rohan Gupta", "class": "10", "date": ""},
                        {"house": "", "name": "Diya Patel", "class": "9", "date": iso(4)}]},
        "02": {"category_guess": "Uniform", "page_date": "",
               "rows": [{"house": "", "name": "Kabir Singh", "class": "9", "date": iso(5)}]},
        "03": None,  # unreadable photo
    }

    def fake_run_ocr(path):
        n = _re.search(r"scan_\d+_\d+_(\d\d)_", os.path.basename(path)).group(1)
        page = pages.get(n)
        if page is None:
            return "", "could not read the photo", None, None
        return _json.dumps(page), None, None, "gemini"

    def photo(name):
        from PIL import Image as _Im
        b = io.BytesIO()
        _Im.new("RGB", (40, 50), "white").save(b, format="PNG")
        b.seek(0)
        return (b, name)

    real_run_ocr = appmod.run_ocr
    appmod.run_ocr = fake_run_ocr
    try:
        r = client.post("/scan", data={"photo": [photo("a.png"), photo("b.png"), photo("c.png")], "hint": "Defaulters"},
                        content_type="multipart/form-data")
        html = r.data.decode()
        check("multi scan: page loads", r.status_code == 200)
        check("multi scan: one section per photo", html.count('class="panel scan-group"') == 3, html[-1500:])
        check("multi scan: register read per photo", 'name="group_category_0" value="LateComers"' in html
              and 'name="group_category_1" value="Uniform"' in html)
        check("multi scan: page date read from the photo", f'name="group_date_0" class="group-date-input" value="{iso(2)}"' in html)
        check("multi scan: no page date -> most common row date",
              f'name="group_date_1" class="group-date-input" value="{iso(5)}"' in html)
        check("multi scan: a row's own date is filled in", f'name="row_date" value="{iso(4)}"' in html)
        check("multi scan: unreadable photo says so, others still shown",
              "Nothing could be read from this photo" in html and "could not read the photo" in html)
        check("multi scan: rows are numbered across photos", 'name="row_include" value="2"' in html)
        thumbs = _re.findall(r'src="(/scan/photo/[^"]+)"', html)
        check("multi scan: thumbnails served", thumbs and client.get(thumbs[0]).status_code == 200, str(thumbs))

        many = [photo(f"p{i}.png") for i in range(12)]
        r = client.post("/scan", data={"photo": many}, content_type="multipart/form-data")
        check("multi scan: more than 10 photos -> only 10 read, with a message",
              r.data.decode().count('class="panel scan-group"') == 10 and b"only the first 10 were read" in r.data)
    finally:
        appmod.run_ocr = real_run_ocr

    # --- saving several photos ---------------------------------------------
    before = {k: len(appmod.read_log_rows(k, limit=100000)) for k in ("LateComers", "Uniform")}
    r = client.post("/scan/commit", data={
        "logged_by": "MultiTest",
        "group_category_0": "LateComers", "group_date_0": iso(2),
        "group_category_1": "Uniform", "group_date_1": iso(5),
        "row_name": ["Rohan Gupta", "Diya Patel", "Kabir Singh"],
        "row_student_id": ["", "", ""], "row_house": ["", "", ""], "row_class": ["", "", ""], "row_remarks": ["", "", ""],
        "row_group": ["0", "0", "1"],
        "row_date": ["", iso(4), ""],
        "row_include": ["0", "1", "2"],
    }, follow_redirects=True)
    late = [x for x in appmod.read_log_rows("LateComers", limit=100000) if x.get("LoggedBy") == "MultiTest"]
    uni = [x for x in appmod.read_log_rows("Uniform", limit=100000) if x.get("LoggedBy") == "MultiTest"]
    check("multi save: rows go to each photo's register",
          len(late) == 2 and len(uni) == 1 and len(appmod.read_log_rows("Uniform", limit=100000)) == before["Uniform"] + 1)
    check("multi save: page date and row date both used",
          {(x["Name"], appmod._row_date(x)) for x in late} == {("Rohan Gupta", real_today - timedelta(days=2)), ("Diya Patel", real_today - timedelta(days=4))}
          and appmod._row_date(uni[0]) == real_today - timedelta(days=5), str(late + uni))
    check("multi save: message lists both registers", b"in Late Comers" in r.data and b"in Uniform" in r.data)

    n_late = len(appmod.read_log_rows("LateComers", limit=100000))
    r = client.post("/scan/commit", data={
        "group_category_0": "LateComers", "group_date_0": iso(1), "group_category_1": "",
        "row_name": ["Rohan Gupta", "Kabir Singh"], "row_group": ["0", "1"], "row_include": ["0", "1"],
    }, follow_redirects=True)
    check("multi save: a photo with no register -> nothing saved at all",
          len(appmod.read_log_rows("LateComers", limit=100000)) == n_late and b"Nothing was saved" in r.data)
    future = (real_today + timedelta(days=3)).isoformat()
    r = client.post("/scan/commit", data={
        "group_category_0": "LateComers", "group_date_0": iso(1),
        "row_name": ["Rohan Gupta", "Diya Patel"], "row_group": ["0", "0"], "row_date": ["", future], "row_include": ["0", "1"],
    }, follow_redirects=True)
    check("multi save: a future row date -> nothing saved",
          len(appmod.read_log_rows("LateComers", limit=100000)) == n_late and b"Diya Patel" in r.data)

    # --- matching scanned names to students --------------------------------
    roster = [
        {"id": "S1", "name": "Samyuktha (7)", "class": "7", "section": "", "house": "Kailash"},
        {"id": "S2", "name": "Samyuktha (9)", "class": "9", "section": "", "house": "Vindhyas"},
        {"id": "S3", "name": "Arun Alagappan", "class": "IB-1", "section": "A", "house": "Kailash"},
        {"id": "S4", "name": "Tarun", "class": "6", "section": "", "house": "Aravallis"},
        {"id": "S5", "name": "Aditya Prasanna", "class": "IB-2", "section": "", "house": "Vindhyas"},
    ]
    ms = lambda n, k="": (appmod.match_student(n, k, roster) or {}).get("id")  # noqa: E731
    check("match: small misread still matches", ms("Arun Alagapan") == "S3")
    check("match: a stranger matches nobody", ms("Mystery Person", "8") is None)
    check("match: same name, class on the page decides", ms("Samyuktha", "9") == "S2" and ms("Samyuktha", "7") == "S1")
    check("match: same name, no class -> left for a person", ms("Samyuktha") is None)
    check("match: IB class written as IB-1A", ms("Arun Alagappan", "IB-1A") == "S3")
    check("match: a short name one letter off is a different child", ms("Arun") is None and ms("Tarun") == "S4")
    check("match: unmatched rows are left unticked",
          appmod._make_candidate("x", "", "Mystery Person", "8", roster)["confident"] is False)

    # --- Tesseract text where the column gaps collapsed to single spaces ---
    roster2 = roster + [
        {"id": "S6", "name": "Akileash", "class": "9", "section": "", "house": "Kailash"},
        {"id": "S7", "name": "Keyaashini", "class": "6", "section": "", "house": "Aravallis"},
    ]
    one = lambda ln: appmod.parse_ocr_lines(ln, roster2)[0]  # noqa: E731
    c = one("Kailash ArunAlagappan IB-1A")
    check("tesseract: a merged name is found inside the line", c["student_id"] == "S3" and c["guess_house"] == "Kailash", str(c))
    check("tesseract: 'Samyuktha 9' run together is Samyuktha (9)", one("Vindhyas Samyuktha 9")["student_id"] == "S2")
    check("tesseract: the class after the name decides", one("Kailash Samyuktha 7")["student_id"] == "S1")
    check("tesseract: a house is never taken for a name (Kailash ~ Akileash)",
          one("Kailash Mystery Person 8")["student_id"] == "" and one("Kailash Akileash 9")["student_id"] == "S6")
    check("tesseract: OCR specks don't break the split", one("Aravallis Keyaashini | 6")["student_id"] == "S7")
    check("tesseract: title with a date and full stop is recognised",
          appmod.guess_category_from_text_lines("LATE COMERS. Date: 17/11/2026\nKailash Arun 9") == "LateComers")

    # --- capitals, sound-alike spellings, houses -------------------------------
    roster3 = roster2 + [
        {"id": "S8", "name": "Achyut", "class": "6", "section": "", "house": "Kailash"},
        {"id": "S9", "name": "Gowtham Karthik", "class": "IB-1", "section": "A", "house": ""},
        {"id": "S10", "name": "Siddharth Kannan", "class": "IB-1", "section": "B", "house": "Aravallis"},
        {"id": "S11", "name": "Siddarth Kannan", "class": "IB-1", "section": "B", "house": "Aravallis"},
        {"id": "S12", "name": "Kabir", "class": "9", "section": "", "house": "Nilgiris"},
        {"id": "S13", "name": "Vindhya Rao", "class": "8", "section": "", "house": "Vindhyas"},
    ]
    m3 = lambda n, k="": (appmod.match_student(n, k, roster3) or {}).get("id")  # noqa: E731
    check("capitals never matter", m3("ACHYUT") == "S8" and m3("arun alagappan") == "S3")
    check("sound-alike at 90%+: Achit -> Achyut", m3("ACHIT", "6") == "S8" and appmod.sound_score("achit", "achyut") >= 0.9)
    check("sound-alike: Gautham -> Gowtham", m3("GAUTHAM KARTHIK") == "S9" and m3("Gautham Karthik", "IB-1A") == "S9")
    check("sound-alike: Kabeer -> Kabir", m3("KABEER") == "S12")
    check("not sound-alike enough: Arun is not Tarun", appmod.sound_score("arun", "tarun") < 0.9 and m3("Arun") is None)
    check("exact spelling wins over a one-letter classmate",
          m3("Siddarth Kannan") == "S11" and m3("SIDDHARTH KANNAN") == "S10")
    check("a misread between two near-identical classmates is left for a person", m3("Sidarth Kannan", "IB-1B") is None)
    check("a house word in the name column matches nobody", m3("Kailash") is None and m3("VINDYAS") is None)
    check("houses: spelling variants become the real house",
          [appmod.match_house(h, roster3) for h in ("VINDYAS", "Aravali", "NILGIRI", "kailas")] == ["Vindhyas", "Aravallis", "Nilgiris", "Kailash"])
    check("houses: an unknown house is kept, capitals tidied", appmod.match_house("RED HOUSE", roster3) == "Red House")
    c = appmod._make_candidate("x", "VINDYAS", "MYSTERY PERSON", "8", roster3)
    check("unmatched rows show tidy capitals and the real house, unticked",
          c["guess_name"] == "Mystery Person" and c["guess_house"] == "Vindhyas" and not c["confident"], str(c))
    check("names with lower case are left as written", appmod.tidy_caps("McDonald") == "McDonald")

    # --- Gemini first, Tesseract as the fallback ----------------------------
    saved = (appmod.USING_GEMINI, appmod.OCR_AVAILABLE)
    try:
        appmod.USING_GEMINI, appmod.OCR_AVAILABLE = True, True
        check("engine label: Gemini primary with Tesseract backup", appmod.ocr_backend_label() == "Gemini (cloud) · Tesseract backup")
        r = client.get("/scan")
        check("scan page: no 'Add Gemini key' button once a key is set", b"Add Gemini key" not in r.data)
        appmod.USING_GEMINI = False
        check("engine label: Tesseract only without a key", appmod.ocr_backend_label() == "Tesseract (offline)")
        r = client.get("/scan")
        check("scan page: offers 'Add Gemini key' on the server PC", b"Add Gemini key" in r.data)
        opened = []
        real_open = appmod._open_file_on_server
        appmod._open_file_on_server = lambda path: opened.append(path)
        try:
            r = client.post("/scan/gemini-key", follow_redirects=True)
        finally:
            appmod._open_file_on_server = real_open
        check("'Add Gemini key' opens SETUP_GEMINI_KEY.bat on the server PC",
              opened and opened[0].endswith("SETUP_GEMINI_KEY.bat") and b"Paste the key" in r.data, str(opened))
    finally:
        appmod.USING_GEMINI, appmod.OCR_AVAILABLE = saved

    # --- saving a whole scan in one write ------------------------------------
    before = {k: len(appmod.read_log_rows(k, 100000)) for k in ("LateComers", "Uniform")}
    batch = [("LateComers", {"Date": real_today, "Name": "Rohan Gupta"}),
             ("Uniform", {"Date": real_today, "Name": "Diya Patel", "Issue": "No tie"}),
             ("LateComers", {"Date": real_today, "Name": "Kabir Singh"})]
    ids = appmod.append_log_entries(batch, "BatchTest", "OCR")
    after = {k: len(appmod.read_log_rows(k, 100000)) for k in ("LateComers", "Uniform")}
    check("batch save: every row written, ids unique",
          len(set(ids)) == 3 and after == {"LateComers": before["LateComers"] + 2, "Uniform": before["Uniform"] + 1}, str(ids))
    real_save = appmod.save_wb
    def locked(_wb):
        raise appmod.WorkbookLocked()
    appmod.save_wb = locked
    try:
        try:
            appmod.append_log_entries(batch, "BatchTest", "OCR")
            raised = False
        except appmod.WorkbookLocked:
            raised = True
    finally:
        appmod.save_wb = real_save
    check("batch save: Excel holding the file -> nothing written at all",
          raised and {k: len(appmod.read_log_rows(k, 100000)) for k in ("LateComers", "Uniform")} == after)

    # --- upload safety -----------------------------------------------------
    name = appmod.safe_upload_name("../../windows/evil name.exe", 0)
    check("upload name can't leave the scans folder", "/" not in name and "\\" not in name and ".." not in name
          and name.endswith(".jpg"), name)
    check("photo route refuses odd names", client.get("/scan/photo/..%2Fapp.py").status_code == 404
          and client.get("/scan/photo/app.py").status_code == 404)


def run_roster_tests():
    import import_roster as ir
    from openpyxl import Workbook, load_workbook as _lw

    # A small roster shaped like the school's file: one sheet per grade.
    roster_path = os.path.join(TMP_DIR, "roster.xlsx")
    wb = Workbook()
    wb.remove(wb.active)
    sheets = {
        "Grade 7": [("Samyuktha", "Kailash"), ("Tarun ", "Vindyas")],
        "Grade 9": [("Samyuktha", "Vindhyas "), ("Advaith Sriram", "Aravalis")],
        "Grade IB-1A": [("Advaith Sriram", "Vindhyas"), ("Gowtham Karthik", None), ("Aarav Sharma", "Kailash")],
        "Grade IB-2": [("Drishya Chordia", "Nilgiris")],
    }
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        ws.append(["Student Name ", "House "])
        for r in rows:
            ws.append(list(r))
    wb.save(roster_path)

    check("grade sheet -> class/section", ir.class_from_sheet("Grade IB-1A") == ("IB-1", "A")
          and ir.class_from_sheet("Grade 10") == ("10", "") and ir.class_from_sheet("Grade IB-2") == ("IB-2", ""))

    # Fresh database whose example students have no entries: they get replaced.
    db = os.path.join(TMP_DIR, "roster_db.xlsx")
    bet.build_workbook().save(db)
    check("roster import runs", ir.main([roster_path, "--db", db]) == 0)
    rows = ir.read_master(_lw(db)["MasterList"])
    by_name = {r["name"]: r for r in rows}
    check("roster: numbered from STU001 in file order", [r["id"] for r in rows] == [f"STU{i:03d}" for i in range(1, 9)]
          and rows[0]["name"] == "Samyuktha (7)", str(rows))
    check("roster: duplicate names get their class", {"Samyuktha (7)", "Samyuktha (9)", "Advaith Sriram (9)", "Advaith Sriram (IB-1A)"} <= set(by_name))
    check("roster: houses cleaned", by_name["Tarun"]["house"] == "Vindhyas" and by_name["Advaith Sriram (9)"]["house"] == "Aravallis"
          and by_name["Samyuktha (9)"]["house"] == "Vindhyas")
    check("roster: IB class and section", by_name["Gowtham Karthik"]["class"] == "IB-1" and by_name["Gowtham Karthik"]["section"] == "A"
          and by_name["Drishya Chordia"]["class"] == "IB-2" and by_name["Drishya Chordia"]["section"] == "")
    check("roster: example students removed", "Rohan Gupta" not in by_name)
    dash = _lw(db)["Dashboard"]
    hdr = next(r for r in range(1, dash.max_row + 1) if dash.cell(row=r, column=2).value == "House")
    check("roster: dashboard houses set", [dash.cell(row=hdr + i, column=2).value for i in range(1, 5)]
          == ["Aravallis", "Kailash", "Nilgiris", "Vindhyas"])

    # Re-running keeps every number, and a student with entries is never dropped.
    appmod_db = appmod.DB_PATH
    appmod.DB_PATH = db
    appmod.append_log_entry("LateComers", {"Date": date.today(), "Name": "Tarun", "Time": "8:50"}, "T", "Manual")
    sid_extra = appmod.add_student("Left School", "8", "", "", "Kailash")
    appmod.DB_PATH = appmod_db
    check("roster import re-run", ir.main([roster_path, "--db", db]) == 0)
    rows2 = ir.read_master(_lw(db)["MasterList"])
    check("roster: numbers stable on re-run", {r["name"]: r["id"] for r in rows2 if r["name"] in by_name}
          == {n: r["id"] for n, r in by_name.items()})
    check("roster: students not in roster are kept", any(r["id"] == sid_extra for r in rows2))

    before = open(db, "rb").read()
    check("roster dry run writes nothing", ir.main([roster_path, "--db", db, "--dry-run"]) == 0 and open(db, "rb").read() == before)


def run_find_tests():
    # A student nobody else's tests touch, with a known record.
    sid = appmod.add_student("Zara Findtest", "11", "C", "7", "Testhouse")
    today = date.today()
    for i, t in enumerate(["8:50 AM", "8:55 AM", "9:00 AM"]):
        appmod.append_log_entry("LateComers", {"Date": today - timedelta(days=i), "Name": "Zara Findtest",
                                               "Time": t, "Reason": "Missed bus", "Remarks": ""}, "Tester", "Manual")
    appmod.append_log_entry("Uniform", {"Date": today, "Name": "Zara Findtest", "Issue": "No tie"}, "Tester", "Manual")

    r = client.get("/find")
    check("find page loads", r.status_code == 200 and b"Find who was late" in r.data)
    check("sidebar links to find", b'href="/find"' in r.data)

    r = client.get("/find?q=findtest")
    check("find by part of a name", r.status_code == 200 and b"Zara Findtest" in r.data and b"1 student<" in r.data, r.data[-3000:])
    r = client.get(f"/find?q={sid}")
    check("find by student ID", b"Zara Findtest" in r.data)
    r = client.get("/find?q=findtest&klass=11&section=C&house=Testhouse")
    check("find with class, section and house", b"Zara Findtest" in r.data)
    r = client.get("/find?q=findtest&house=Kailash")
    check("house filter excludes other houses", b"Zara Findtest" not in r.data.split(b"<section class=\"panel\">")[-1])

    res = appmod.find_students(appmod.read_find_filters({"q": "findtest"}))
    check("late count is 3 in Late Comers", len(res) == 1 and res[0]["count"] == 3, str(res))
    check("usual arrival is the median time", res and res[0]["usual_time"] == "8:55 AM", str(res))
    res = appmod.find_students(appmod.read_find_filters({"q": "findtest", "register": "all"}))
    check("all registers counts every entry", res and res[0]["count"] == 4 and res[0]["by_register"]["Uniform"] == 1, str(res))
    check("repeat filter keeps 3+", len(appmod.find_students(appmod.read_find_filters({"q": "findtest", "min": "3"}))) == 1)
    check("repeat filter drops below 5", appmod.find_students(appmod.read_find_filters({"q": "findtest", "min": "5"})) == [])
    check("today filter counts only today",
          appmod.find_students(appmod.read_find_filters({"q": "findtest", "period": "today"}))[0]["count"] == 1)
    check("bad filter values fall back safely", appmod.read_find_filters({"register": "Nope", "period": "x", "min": "abc"})
          == {"q": "", "klass": "", "section": "", "house": "", "register": "LateComers", "period": "all", "min": 1})

    r = client.get("/find/export.xlsx?q=findtest&register=all")
    from openpyxl import load_workbook as _lw
    ok = r.status_code == 200 and r.data[:2] == b"PK"
    if ok:
        ws = _lw(io.BytesIO(r.data)).active
        names = [ws.cell(row=rr, column=1).value for rr in range(5, ws.max_row + 1)]
        ok = names == ["Zara Findtest"] and ws.cell(row=5, column=5).value == 4
    check("find Excel export has the filtered rows", ok)

    r = client.get("/student?name=Zara+Findtest")
    check("student profile loads", r.status_code == 200 and b"Zara Findtest" in r.data and sid.encode() in r.data)
    check("profile shows history and patterns", b"Missed bus" in r.data and b"No tie" in r.data and b"Usually arrives around 8:55 AM" in r.data)
    r = client.get(f"/student?name={sid}")
    check("profile opens by student ID", r.status_code == 200 and b"Zara Findtest" in r.data)
    r = client.get("/student?name=Nobody+Atall", follow_redirects=True)
    check("unknown student redirects with a message", r.status_code == 200 and b"No student called" in r.data)
    r = client.get("/student/export.xlsx?name=Zara+Findtest")
    ok = r.status_code == 200 and r.data[:2] == b"PK"
    if ok:
        ws = _lw(io.BytesIO(r.data)).active
        ok = ws.max_row - 4 == 4  # four entries under the header row
    check("student Excel export has every entry", ok)

    r = client.get("/entry/LateComers")
    check("register names link to the profile", b'/student?name=Zara+Findtest' in r.data)


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
